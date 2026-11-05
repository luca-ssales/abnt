"""
Testes dos endpoints da API FastAPI:
- Validação de formato (.docx obrigatório)
- Rejeição de arquivo vazio ou inválido
- Endpoint de formatação e download do documento formatado
"""

from pathlib import Path
import pytest
from httpx import AsyncClient, ASGITransport
from app.main import app
from tests.generate_fixtures import create_sample_documents


@pytest.fixture(scope="session")
def sample_docx():
    fixtures_dir = Path(__file__).resolve().parent / "fixtures"
    bad_path, _ = create_sample_documents(fixtures_dir)
    return bad_path


@pytest.mark.anyio
async def test_capabilities_endpoint():
    transport = ASGITransport(app=app)
    async with AsyncClient(transport=transport, base_url="http://test") as client:
        res = await client.get("/api/capabilities")
        assert res.status_code == 200
        data = res.json()
        assert "max_upload_size_mb" in data
        assert "allowed_fonts" in data
        assert "work_types" in data


@pytest.mark.anyio
async def test_rejects_non_docx_file():
    transport = ASGITransport(app=app)
    async with AsyncClient(transport=transport, base_url="http://test") as client:
        files = {"file": ("documento.txt", b"Texto de teste", "text/plain")}
        res = await client.post("/api/format", files=files)
        assert res.status_code == 400
        assert "Formato de arquivo inválido" in res.json()["detail"]


@pytest.mark.anyio
async def test_format_and_download_flow(sample_docx):
    transport = ASGITransport(app=app)
    async with AsyncClient(transport=transport, base_url="http://test") as client:
        docx_bytes = sample_docx.read_bytes()
        files = {"file": ("meu_tcc.docx", docx_bytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")}
        data = {
            "font_name": "Times New Roman",
            "font_size": 12,
            "work_type": "tcc",
            "insert_auto_toc": False,
        }
        res = await client.post("/api/format", files=files, data=data)
        assert res.status_code == 200
        res_json = res.json()
        assert "job_id" in res_json
        assert "report" in res_json
        assert len(res_json["report"]["changes_summary"]) > 0

        # Testa download do arquivo formatado
        job_id = res_json["job_id"]
        dl_res = await client.get(f"/api/download/{job_id}/docx")
        assert dl_res.status_code == 200
        assert len(dl_res.content) > 1000

