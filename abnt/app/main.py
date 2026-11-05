"""
Servidor FastAPI para o site "Formatador ABNT".
"""

import os
import shutil
import tempfile
import time
import uuid
import zipfile
from pathlib import Path
from typing import Dict

from fastapi import FastAPI, File, Form, HTTPException, UploadFile, BackgroundTasks
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from app.abnt_rules import ABNT_RULES
from app.formatter import format_docx_abnt
from app.pdf_converter import is_libreoffice_available, convert_docx_to_pdf


app = FastAPI(
    title="Formatador ABNT",
    description="Formatação automática de documentos .docx nas normas ABNT com 100% de preservação do texto.",
    version="1.0.0",
)

BASE_DIR = Path(__file__).resolve().parent.parent
STATIC_DIR = BASE_DIR / "static"
TEMP_STORAGE_DIR = Path(tempfile.gettempdir()) / "formatador_abnt_jobs"
TEMP_STORAGE_DIR.mkdir(parents=True, exist_ok=True)

JOBS_STORE: Dict[str, dict] = {}
JOB_TTL_SECONDS = 600


def cleanup_expired_jobs() -> None:
    now = time.time()
    expired_ids = [
        jid for jid, meta in list(JOBS_STORE.items())
        if now - meta.get("created_at", 0) > JOB_TTL_SECONDS
    ]
    for jid in expired_ids:
        remove_job_artifacts(jid)


def remove_job_artifacts(job_id: str) -> None:
    meta = JOBS_STORE.pop(job_id, None)
    if meta and meta.get("job_dir"):
        shutil.rmtree(meta["job_dir"], ignore_errors=True)


def validate_uploaded_docx(file_bytes: bytes, filename: str) -> None:
    if not filename.lower().endswith(".docx"):
        raise HTTPException(
            status_code=400,
            detail="Formato de arquivo inválido. Por favor, envie apenas arquivos do Word no formato .docx.",
        )

    max_bytes = ABNT_RULES.max_upload_size_mb * 1024 * 1024
    if len(file_bytes) > max_bytes:
        raise HTTPException(
            status_code=413,
            detail=f"O arquivo excede o tamanho máximo permitido de {ABNT_RULES.max_upload_size_mb} MB.",
        )

    if len(file_bytes) == 0:
        raise HTTPException(
            status_code=400,
            detail="O arquivo enviado está vazio.",
        )

    if file_bytes[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        raise HTTPException(
            status_code=422,
            detail="Este arquivo .docx está protegido por senha ou salvo no formato binário antigo (.doc). Remova a senha no Word e tente novamente.",
        )


@app.get("/api/capabilities")
async def get_capabilities():
    cleanup_expired_jobs()
    return {
        "pdf_available": is_libreoffice_available(),
        "max_upload_size_mb": ABNT_RULES.max_upload_size_mb,
        "allowed_fonts": ABNT_RULES.allowed_fonts,
        "default_font": ABNT_RULES.default_font,
        "work_types": {
            k: {"name": v.name, "description": v.description}
            for k, v in ABNT_RULES.work_types.items()
        },
    }


@app.post("/api/format")
async def format_document_endpoint(
    file: UploadFile = File(...),
    font_name: str = Form("Times New Roman"),
    font_size: int = Form(12),
    work_type: str = Form("tcc"),
    insert_auto_toc: bool = Form(False),
):
    cleanup_expired_jobs()

    raw_bytes = await file.read()
    original_filename = file.filename or "documento.docx"
    validate_uploaded_docx(raw_bytes, original_filename)

    job_id = uuid.uuid4().hex
    job_dir = TEMP_STORAGE_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=True)

    input_path = job_dir / "original_upload.docx"
    output_docx_path = job_dir / "documento_formatado_abnt.docx"

    try:
        input_path.write_bytes(raw_bytes)

        if not zipfile.is_zipfile(input_path):
            raise HTTPException(
                status_code=422,
                detail="O arquivo enviado está corrompido ou não é um pacote .docx válido.",
            )

        with zipfile.ZipFile(input_path, "r") as zf:
            names = zf.namelist()
            if "EncryptedPackage" in names or "EncryptionInfo" in names:
                raise HTTPException(
                    status_code=422,
                    detail="Este documento está protegido por senha. Remova a proteção no Word antes de formatar.",
                )
            if "word/document.xml" not in names:
                raise HTTPException(
                    status_code=422,
                    detail="O arquivo .docx não possui a estrutura interna válida do Word (word/document.xml ausente).",
                )

        report = format_docx_abnt(
            input_path=str(input_path),
            output_path=str(output_docx_path),
            font_name=font_name,
            font_size_pt=font_size,
            work_type=work_type,
            insert_auto_toc=insert_auto_toc,
        )

    except HTTPException:
        shutil.rmtree(job_dir, ignore_errors=True)
        raise
    except Exception as exc:
        shutil.rmtree(job_dir, ignore_errors=True)
        raise HTTPException(
            status_code=500,
            detail=f"Não foi possível processar o documento: {str(exc)}",
        )
    finally:
        if input_path.exists():
            try:
                input_path.unlink()
            except OSError:
                pass

    safe_stem = Path(original_filename).stem
    JOBS_STORE[job_id] = {
        "job_dir": str(job_dir),
        "docx_path": str(output_docx_path),
        "pdf_path": None,
        "download_stem": f"{safe_stem}_ABNT",
        "created_at": time.time(),
        "pdf_available": is_libreoffice_available(),
    }

    return JSONResponse({
        "job_id": job_id,
        "download_docx_url": f"/api/download/{job_id}/docx",
        "download_pdf_url": f"/api/download/{job_id}/pdf" if is_libreoffice_available() else None,
        "pdf_available": is_libreoffice_available(),
        "report": report.to_dict(),
    })


@app.get("/api/download/{job_id}/docx")
async def download_formatted_docx(job_id: str, background_tasks: BackgroundTasks):
    meta = JOBS_STORE.get(job_id)
    if not meta or not os.path.isfile(meta["docx_path"]):
        raise HTTPException(
            status_code=404,
            detail="O arquivo solicitado expirou ou já foi removido do servidor por privacidade. Envie o documento novamente.",
        )

    filename = f"{meta['download_stem']}.docx"
    if not meta.get("pdf_available"):
        background_tasks.add_task(remove_job_artifacts, job_id)

    return FileResponse(
        path=meta["docx_path"],
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=filename,
    )


@app.get("/api/download/{job_id}/pdf")
async def download_formatted_pdf(job_id: str, background_tasks: BackgroundTasks):
    meta = JOBS_STORE.get(job_id)
    if not meta or not os.path.isfile(meta["docx_path"]):
        raise HTTPException(
            status_code=404,
            detail="O arquivo solicitado expirou ou já foi removido do servidor.",
        )

    if not is_libreoffice_available():
        raise HTTPException(
            status_code=501,
            detail="O LibreOffice não está instalado no servidor para conversão em PDF.",
        )

    pdf_path = meta.get("pdf_path")
    if not pdf_path or not os.path.isfile(pdf_path):
        pdf_path = convert_docx_to_pdf(meta["docx_path"], meta["job_dir"])
        if not pdf_path:
            raise HTTPException(
                status_code=500,
                detail="Falha ao converter o documento para PDF usando LibreOffice.",
            )
        meta["pdf_path"] = pdf_path

    filename = f"{meta['download_stem']}.pdf"
    return FileResponse(
        path=pdf_path,
        media_type="application/pdf",
        filename=filename,
    )


@app.delete("/api/jobs/{job_id}")
async def delete_job_now(job_id: str):
    remove_job_artifacts(job_id)
    return {"status": "deleted"}


app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")


@app.get("/", response_class=HTMLResponse)
async def serve_index():
    index_file = STATIC_DIR / "index.html"
    return HTMLResponse(content=index_file.read_text(encoding="utf-8"))

