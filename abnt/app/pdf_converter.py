"""
Integração opcional com LibreOffice em modo headless para conversão de .docx em .pdf.
"""

import os
import shutil
import subprocess
from pathlib import Path
from typing import Optional


COMMON_WINDOWS_PATHS = [
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
]


def get_soffice_executable() -> Optional[str]:
    for cmd in ("soffice", "libreoffice"):
        found = shutil.which(cmd)
        if found:
            return found

    if os.name == "nt":
        for candidate in COMMON_WINDOWS_PATHS:
            if os.path.isfile(candidate):
                return candidate

    return None


def is_libreoffice_available() -> bool:
    return get_soffice_executable() is not None


def convert_docx_to_pdf(docx_path: str, output_dir: str, timeout_seconds: int = 45) -> Optional[str]:
    soffice_bin = get_soffice_executable()
    if not soffice_bin:
        return None

    docx_file = Path(docx_path)
    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    try:
        subprocess.run(
            [
                soffice_bin,
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                str(out_dir),
                str(docx_file),
            ],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=timeout_seconds,
        )
    except (subprocess.SubprocessError, OSError):
        return None

    expected_pdf = out_dir / f"{docx_file.stem}.pdf"
    if expected_pdf.is_file():
        return str(expected_pdf)
    return None

