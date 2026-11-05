"""
Testes unitários e de integração do Formatador ABNT.
Garante a REGRA DE OURO: o texto extraído antes e depois da formatação deve ser 100% IDÊNTICO.
"""

import os
from pathlib import Path
import pytest
from docx import Document

from app.formatter import format_docx_abnt, extract_complete_document_fingerprint
from app.abnt_rules import ABNT_RULES
from tests.generate_fixtures import create_sample_documents


@pytest.fixture(scope="session")
def sample_docs(tmp_path_factory):
    """Gera fixtures de documentos de teste."""
    fixtures_dir = Path(__file__).resolve().parent / "fixtures"
    bad_path, good_path = create_sample_documents(fixtures_dir)
    return bad_path, good_path


def test_sample_baguncado_preserves_text_100_percent(sample_docs, tmp_path):
    """Valida que o documento desformatado tem seu texto 100% inalterado."""
    bad_path, _ = sample_docs
    output_path = tmp_path / "formatted_baguncado.docx"

    doc_before = Document(str(bad_path))
    fingerprint_before = extract_complete_document_fingerprint(doc_before)

    report = format_docx_abnt(
        input_path=str(bad_path),
        output_path=str(output_path),
        font_name="Times New Roman",
        font_size_pt=12,
        work_type="tcc",
    )

    doc_after = Document(str(output_path))
    fingerprint_after = extract_complete_document_fingerprint(doc_after)

    # REGRA DE OURO: Comparação estrita de parágrafos, tabelas e notas
    assert fingerprint_before["paragraphs"] == fingerprint_after["paragraphs"], "O texto dos parágrafos foi alterado!"
    assert fingerprint_before["tables"] == fingerprint_after["tables"], "O texto das tabelas foi alterado!"
    assert fingerprint_before["footnotes"] == fingerprint_after["footnotes"], "O texto das notas foi alterado!"

    # Validações de conformidade ABNT no documento de saída
    for sec in doc_after.sections:
        assert round(sec.top_margin.cm, 1) == 3.0
        assert round(sec.left_margin.cm, 1) == 3.0
        assert round(sec.bottom_margin.cm, 1) == 2.0
        assert round(sec.right_margin.cm, 1) == 2.0

    # Verifica se os contadores do relatório foram preenchidos
    assert report.sections_margins_adjusted > 0
    assert report.paragraphs_total > 0
    assert report.headings_formatted["primaria_1"] >= 1


def test_sample_correto_preserves_text_and_runs_smoothly(sample_docs, tmp_path):
    """Valida que um documento já correto não sofre alterações destrutivas."""
    _, good_path = sample_docs
    output_path = tmp_path / "formatted_correto.docx"

    doc_before = Document(str(good_path))
    fingerprint_before = extract_complete_document_fingerprint(doc_before)

    report = format_docx_abnt(
        input_path=str(good_path),
        output_path=str(output_path),
        font_name="Arial",
        font_size_pt=12,
        work_type="tcc",
    )

    doc_after = Document(str(output_path))
    fingerprint_after = extract_complete_document_fingerprint(doc_after)

    assert fingerprint_before["paragraphs"] == fingerprint_after["paragraphs"]
    assert fingerprint_before["tables"] == fingerprint_after["tables"]
    assert report.font_applied == "Arial"

