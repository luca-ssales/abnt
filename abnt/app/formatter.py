"""
Motor principal de formatação ABNT (NBR 14724:2024, NBR 6024:2012, NBR 10520:2023, NBR 6023:2025, NBR 6028:2021)
e Manual Institucional de Normalização Acadêmica (Edição 2026).

REGRA DE OURO E MANUAL ACADÊMICO:
- Separação estrita entre CAPA e FOLHA DE ROSTO em páginas independentes.
- Distribuição vertical harmônica e elegante na página:
  * CAPA:
    - Instituição/Curso no topo (tamanho 14, centralizado, sem negrito)
    - Autor com espaçamento vertical planejado (tamanho 14, centralizado)
    - Título e Subtítulo centralizados no centro da mancha gráfica (tamanho 14, Título em negrito)
    - Local e Ano fixados na margem inferior da folha (tamanho 14, centralizado, sem negrito)
    - Finalização com quebra de página (<w:br w:type="page"/>)
  * FOLHA DE ROSTO:
    - Autor no topo (tamanho 14, negrito)
    - RA se presente (posicionado junto ao autor)
    - Título e Subtítulo no terço médio (tamanho 14, Título em negrito)
    - Natureza do trabalho e Orientador do meio da mancha até a margem direita (recuo 7.5 cm, tamanho 10, simples)
    - Local e Ano na base da página (tamanho 14, sem negrito), seguidos de quebra de página
  * SEÇÕES E ELEMENTOS:
    - Seções Primárias (1, 2...) e sem numeração (RESUMO, SUMÁRIO, REFERÊNCIAS): tamanho 14, negrito, maiúsculas
    - Seções Secundárias (1.1): tamanho 12, negrito
    - Seções Terciárias/Quaternárias (1.1.1): tamanho 12
    - Citações diretas longas (> 3 linhas): recuo 4 cm, tamanho 10, simples, 18 pt antes e depois
    - Ilustrações e Tabelas: título no topo (10 pt), fonte na base (10 pt), centralizadas
    - Paginação: algarismos arábicos a 2 cm da borda superior no canto direito, exibida a partir da Introdução
- Preserva integralmente 100% de todo o conteúdo original do documento enviado.
"""

import re
from typing import Tuple, List, Optional
from docx import Document
from docx.document import Document as DocxDocument
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

from app.abnt_rules import ABNT_RULES
from app.detector import (
    ParagraphRole,
    classify_paragraph,
    _normalize_title,
    TOC_LINE_RE,
    has_manual_page_break,
    NATUREZA_KEYWORDS_RE,
)
from app.report import FormattingReport
from app.xml_helpers import (
    iter_all_Paragraph_runs_including_hyperlinks,
    apply_font_to_r_element,
    add_page_number_to_header,
    clear_header_page_numbers,
    insert_section_break_before_paragraph,
    format_footnotes_xml,
    attach_toc_field_to_paragraph,
)


YEAR_RE = re.compile(r"^\s*20\d{2}\s*$")
RA_RE = re.compile(r"^\s*(?:RA|R\.A\.|Matr[íi]cula)\s*:", re.IGNORECASE)
ORIENTADOR_RE = re.compile(r"^\s*(?:orientador|orientadora|coorientador|coorientadora|professor|professora)\s*:", re.IGNORECASE)


def extract_complete_document_fingerprint(doc: DocxDocument) -> dict:
    paragraphs_text = [p.text for p in doc.paragraphs]

    tables_text = []
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                tables_text.append("\n".join(p.text for p in cell.paragraphs))

    footnotes_text = []
    for part in doc.part.package.parts:
        if "footnotes.xml" in str(part.partname).lower():
            el = getattr(part, "element", None)
            if el is not None:
                texts = el.xpath(".//w:t/text()")
                footnotes_text.append("".join(texts))

    return {
        "paragraphs": paragraphs_text,
        "tables": tables_text,
        "footnotes": footnotes_text,
    }


def _collect_original_typography(paragraph: Paragraph, report: FormattingReport) -> None:
    if paragraph.style and paragraph.style.font:
        if paragraph.style.font.name:
            report.original_fonts.add(paragraph.style.font.name)
        if paragraph.style.font.size:
            report.original_font_sizes.add(round(paragraph.style.font.size.pt, 1))

    for run in paragraph.runs:
        if run.font.name:
            report.original_fonts.add(run.font.name)
        if run.font.size:
            report.original_font_sizes.add(round(run.font.size.pt, 1))


def _format_paragraph_runs(
    paragraph: Paragraph,
    font_name: str,
    font_size_pt: float,
    bold: bool = None,
    italic: bool = None,
    uppercase: bool = None,
) -> None:
    for r_el in iter_all_Paragraph_runs_including_hyperlinks(paragraph):
        apply_font_to_r_element(
            r_el,
            font_name=font_name,
            font_size_pt=font_size_pt,
            bold=bold,
            italic=italic,
            uppercase=uppercase,
            force_black=True,
        )


def _apply_paragraph_layout(
    paragraph: Paragraph,
    alignment: WD_ALIGN_PARAGRAPH,
    line_spacing: float,
    space_before_pt: float,
    space_after_pt: float,
    left_indent_cm: float = 0.0,
    first_line_indent_cm: float = 0.0,
    page_break_before: bool = False,
    keep_with_next: bool = False,
    report: FormattingReport = None,
) -> None:
    pf = paragraph.paragraph_format

    if report is not None:
        if paragraph.alignment != alignment:
            report.paragraphs_alignment_adjusted += 1
        if pf.line_spacing != line_spacing or pf.space_before != Pt(space_before_pt) or pf.space_after != Pt(space_after_pt):
            report.paragraphs_spacing_adjusted += 1
        target_first_indent = Cm(first_line_indent_cm) if first_line_indent_cm > 0 else Cm(0)
        if first_line_indent_cm > 0 and pf.first_line_indent != target_first_indent:
            report.paragraphs_indent_adjusted += 1

    paragraph.alignment = alignment
    pf.line_spacing = line_spacing
    pf.space_before = Pt(space_before_pt)
    pf.space_after = Pt(space_after_pt)
    pf.left_indent = Cm(left_indent_cm)
    pf.right_indent = Cm(0)
    pf.first_line_indent = Cm(first_line_indent_cm)
    pf.page_break_before = page_break_before
    pf.keep_with_next = keep_with_next


def _configure_base_styles(doc: DocxDocument, font_name: str, font_size_pt: int) -> None:
    try:
        normal_style = doc.styles["Normal"]
        normal_style.font.name = font_name
        normal_style.font.size = Pt(font_size_pt)
        normal_style.font.color.rgb = RGBColor(0, 0, 0)
    except KeyError:
        pass


def _configure_sections_and_margins(doc: DocxDocument, report: FormattingReport) -> None:
    for section in doc.sections:
        section.page_width = Cm(ABNT_RULES.page_width_cm)
        section.page_height = Cm(ABNT_RULES.page_height_cm)
        section.top_margin = Cm(ABNT_RULES.margin_top_cm)
        section.left_margin = Cm(ABNT_RULES.margin_left_cm)
        section.bottom_margin = Cm(ABNT_RULES.margin_bottom_cm)
        section.right_margin = Cm(ABNT_RULES.margin_right_cm)
        section.header_distance = Cm(ABNT_RULES.header_distance_cm)
        report.sections_margins_adjusted += 1


def _format_tables(doc: DocxDocument, font_name: str, report: FormattingReport) -> None:
    for table in doc.tables:
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    _collect_original_typography(p, report)
                    p.paragraph_format.line_spacing = ABNT_RULES.table_line_spacing
                    p.paragraph_format.space_before = Pt(2)
                    p.paragraph_format.space_after = Pt(2)
                    p.paragraph_format.first_line_indent = Cm(0)
                    _format_paragraph_runs(
                        p,
                        font_name=font_name,
                        font_size_pt=ABNT_RULES.table_font_size_pt,
                    )
        report.tables_formatted += 1


def _separate_capa_and_folha_de_rosto(doc: DocxDocument, font_name: str, report: FormattingReport) -> bool:
    """
    Separa rigorosamente CAPA e FOLHA DE ROSTO conforme o Manual Institucional (2026, p. 9, 10, 42 e 44).
    Garante:
    - CAPA:
      * Instituição: tamanho 14, sem negrito, centralizado no topo.
      * Autor: tamanho 14, sem negrito, centralizado com espaçamento vertical de 120 pt.
      * Título e Subtítulo: tamanho 14, Título em negrito e Subtítulo normal, centralizado com espaçamento de 160 pt.
      * Local e Ano: tamanho 14, sem negrito, centralizados na margem inferior.
      * Finaliza com quebra de página explícita.
    - FOLHA DE ROSTO:
      * Autor: tamanho 14, negrito, centralizado no topo (+ RA logo abaixo).
      * Título: tamanho 14, negrito, centralizado com espaçamento de 50 pt.
      * Natureza do trabalho e Orientador: tamanho 10, simples, recuo de 7.5 cm à direita.
      * Local e Ano: tamanho 14, sem negrito, na margem inferior com quebra de página.
    """
    resumo_idx = -1
    for i, p in enumerate(doc.paragraphs):
        if _normalize_title(p.text) in ("RESUMO", "ABSTRACT"):
            resumo_idx = i
            break

    if resumo_idx <= 0:
        return False

    pre_pars = [p for i, p in enumerate(doc.paragraphs) if i < resumo_idx]

    has_natureza = any(NATUREZA_KEYWORDS_RE.search(p.text) for p in pre_pars)
    if not has_natureza:
        return False

    page_break_count = sum(1 for p in pre_pars if has_manual_page_break(p))
    if page_break_count >= 2:
        return False

    inst_pars: List[Paragraph] = []
    author_pars: List[Paragraph] = []
    ra_pars: List[Paragraph] = []
    title_pars: List[Paragraph] = []
    nat_pars: List[Paragraph] = []
    ori_pars: List[Paragraph] = []
    local_pars: List[Paragraph] = []
    year_pars: List[Paragraph] = []
    empty_with_br_pars: List[Paragraph] = []

    for p in pre_pars:
        t = p.text.strip()
        if not t:
            if has_manual_page_break(p):
                empty_with_br_pars.append(p)
            continue

        if RA_RE.match(t):
            ra_pars.append(p)
        elif ORIENTADOR_RE.match(t):
            ori_pars.append(p)
        elif NATUREZA_KEYWORDS_RE.search(t):
            if re.search(r"(?:orientador|professor)\b", t, re.IGNORECASE):
                ori_pars.append(p)
            else:
                nat_pars.append(p)
        elif YEAR_RE.match(t):
            year_pars.append(p)
        elif any(k in t.upper() for k in ["UNIVERSIDADE", "FACULDADE", "CENTRO UNIVERSITÁRIO", "INSTITUTO", "CURSO DE", "DEPARTAMENTO"]):
            inst_pars.append(p)
        elif len(author_pars) == 0 and len(title_pars) == 0:
            author_pars.append(p)
        elif len(year_pars) == 0 and ("–" in t or "-" in t or t.upper() in ["SÃO PAULO", "GUARULHOS", "RIO DE JANEIRO", "BRASÍLIA", "BELO HORIZONTE", "CURITIBA", "PORTO ALEGRE", "SALVADOR", "RECIFE", "FORTALEZA"]) and len(t) < 40:
            local_pars.append(p)
        else:
            title_pars.append(p)

    if not nat_pars:
        return False

    resumo_p = doc.paragraphs[resumo_idx]

    def clone_paragraph_element(src_p: Paragraph):
        return parse_xml(src_p._p.xml)

    # 1. Configura a CAPA (Manual p. 9 e p. 42):
    # Instituição (Arial 14, sem negrito, centralizado)
    for p in inst_pars:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.line_spacing = ABNT_RULES.body_line_spacing
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(6)
        _format_paragraph_runs(p, font_name=font_name, font_size_pt=14, bold=False)

    # Autor da Capa (Arial 14, sem negrito, centralizado com espaço vertical de 120 pt até o título)
    for idx_a, p in enumerate(author_pars):
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.line_spacing = ABNT_RULES.body_line_spacing
        p.paragraph_format.space_before = Pt(24 if inst_pars else 0)
        p.paragraph_format.space_after = Pt(120 if idx_a == len(author_pars) - 1 else 6)
        _format_paragraph_runs(p, font_name=font_name, font_size_pt=14, bold=False)

    # Título da Capa (Arial 14, Título em negrito e subtítulo sem negrito, centralizado no meio da folha com 160 pt após)
    for idx_t, p in enumerate(title_pars):
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.line_spacing = ABNT_RULES.body_line_spacing
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(160 if idx_t == len(title_pars) - 1 else 6)
        is_subtitle = idx_t > 0
        _format_paragraph_runs(p, font_name=font_name, font_size_pt=14, bold=not is_subtitle)

    # Local da Capa (Arial 14, sem negrito, centralizado na margem inferior)
    for p in local_pars:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.line_spacing = ABNT_RULES.body_line_spacing
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(6)
        _format_paragraph_runs(p, font_name=font_name, font_size_pt=14, bold=False)

    # Ano da Capa (Arial 14, sem negrito, centralizado)
    for p in year_pars:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.line_spacing = ABNT_RULES.body_line_spacing
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        _format_paragraph_runs(p, font_name=font_name, font_size_pt=14, bold=False)

    # Insere quebra de página no final da Capa
    if year_pars:
        last_year_p = year_pars[-1]
        if not has_manual_page_break(last_year_p):
            r_br = last_year_p.add_run()
            br_el = OxmlElement("w:br")
            br_el.set(qn("w:type"), "page")
            r_br._r.append(br_el)

    # 2. Monta a FOLHA DE ROSTO (Manual p. 10 e p. 44):
    fr_elements = []

    # Autor na Folha de Rosto (Arial 14, negrito, centralizado no topo)
    for p in author_pars:
        fr_author_el = clone_paragraph_element(p)
        p_temp = Paragraph(fr_author_el, doc)
        p_temp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_temp.paragraph_format.space_before = Pt(0)
        p_temp.paragraph_format.space_after = Pt(6 if ra_pars else 110)
        _format_paragraph_runs(p_temp, font_name=font_name, font_size_pt=14, bold=True)
        fr_elements.append(fr_author_el)

    # RA se houver (posicionado junto ao autor)
    for p in ra_pars:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(100)
        _format_paragraph_runs(p, font_name=font_name, font_size_pt=12, bold=False)
        fr_elements.append(p._p)

    # Título na Folha de Rosto (Arial 14, Título negrito, centralizado com 50 pt após)
    for idx_t, p in enumerate(title_pars):
        fr_title_el = clone_paragraph_element(p)
        p_temp = Paragraph(fr_title_el, doc)
        p_temp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_temp.paragraph_format.space_before = Pt(0)
        p_temp.paragraph_format.space_after = Pt(50 if idx_t == len(title_pars) - 1 else 6)
        is_subtitle = idx_t > 0
        _format_paragraph_runs(p_temp, font_name=font_name, font_size_pt=14, bold=not is_subtitle)
        fr_elements.append(fr_title_el)

    # Natureza do trabalho (Manual p. 10: Arial 10, simples, recuo 7.5 cm à direita, justificado)
    for p in nat_pars:
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.left_indent = Cm(ABNT_RULES.natureza_left_indent_cm)
        p.paragraph_format.line_spacing = 1.0
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(12)
        _format_paragraph_runs(p, font_name=font_name, font_size_pt=10, bold=False)
        fr_elements.append(p._p)

    # Orientador/Professor (Manual p. 10: Arial 10, simples, recuo 7.5 cm à direita com 90 pt após)
    for p in ori_pars:
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.left_indent = Cm(ABNT_RULES.natureza_left_indent_cm)
        p.paragraph_format.line_spacing = 1.0
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(90)
        _format_paragraph_runs(p, font_name=font_name, font_size_pt=10, bold=False)
        fr_elements.append(p._p)

    # Local e Ano na Folha de Rosto (Arial 14, sem negrito, centralizado)
    for p in local_pars:
        fr_loc_el = clone_paragraph_element(p)
        p_temp = Paragraph(fr_loc_el, doc)
        p_temp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_temp.paragraph_format.space_before = Pt(0)
        p_temp.paragraph_format.space_after = Pt(6)
        _format_paragraph_runs(p_temp, font_name=font_name, font_size_pt=14, bold=False)
        fr_elements.append(fr_loc_el)

    for p in year_pars:
        fr_yr_el = clone_paragraph_element(p)
        p_temp = Paragraph(fr_yr_el, doc)
        p_temp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_temp.paragraph_format.space_before = Pt(0)
        p_temp.paragraph_format.space_after = Pt(0)
        _format_paragraph_runs(p_temp, font_name=font_name, font_size_pt=14, bold=False)
        fr_elements.append(fr_yr_el)

    # Remove os parágrafos originais da Capa que agora residem na Folha de Rosto
    for p in nat_pars + ori_pars + ra_pars:
        parent = p._p.getparent()
        if parent is not None:
            parent.remove(p._p)

    # Insere os elementos da Folha de Rosto imediatamente antes do Resumo
    for el in fr_elements:
        resumo_p._p.addprevious(el)

    # Remove parágrafos vazios duplicados e assegura que a Folha de Rosto termine com quebra de página
    for p_empty in empty_with_br_pars:
        parent = p_empty._p.getparent()
        if parent is not None:
            parent.remove(p_empty._p)

    if fr_elements:
        last_fr_el = fr_elements[-1]
        p_last_fr = Paragraph(last_fr_el, doc)
        if not has_manual_page_break(p_last_fr):
            r_br_fr = p_last_fr.add_run()
            br_el_fr = OxmlElement("w:br")
            br_el_fr.set(qn("w:type"), "page")
            r_br_fr._r.append(br_el_fr)

    report.changes_summary.append(
        "Capa e Folha de Rosto geradas em páginas independentes conforme Manual 2026 / NBR 14724:2024: Capa em tamanho 14 com distribuição vertical calculada; Folha de Rosto com autor no topo e nota de apresentação recuada a 7,5 cm em tamanho 10."
    )
    return True


def format_docx_abnt(
    input_path: str,
    output_path: str,
    font_name: str = "Arial",
    font_size_pt: int = 12,
    work_type: str = "tcc",
    insert_auto_toc: bool = False,
) -> FormattingReport:
    if font_name not in ABNT_RULES.allowed_fonts:
        font_name = ABNT_RULES.default_font

    work_cfg = ABNT_RULES.work_types.get(work_type, ABNT_RULES.work_types["tcc"])

    doc = Document(input_path)

    report = FormattingReport(
        font_applied=font_name,
        font_size_applied=font_size_pt,
        work_type=work_type,
    )

    _configure_base_styles(doc, font_name, font_size_pt)

    # 1. Separação de Capa e Folha de Rosto conforme Manual Institucional 2026 / NBR 14724:2024
    if work_cfg.primary_heading_page_break:
        _separate_capa_and_folha_de_rosto(doc, font_name, report)

    current_section_context = "PRE_TEXTUAL"
    has_passed_textual_start = False
    first_textual_paragraph = None
    first_textual_index = -1
    sumario_heading_paragraph = None

    paragraphs = doc.paragraphs
    report.paragraphs_total = len(paragraphs)

    # 2. Pré-varredura: localiza com exatidão a Introdução real no corpo textual
    for idx, p in enumerate(paragraphs):
        t_raw = p.text.strip()
        if not t_raw:
            continue
        norm_u = _normalize_title(t_raw)
        if TOC_LINE_RE.match(t_raw):
            continue
        if norm_u in ABNT_RULES.textual_start_markers or (norm_u.startswith("1 ") and "..." not in t_raw):
            first_textual_paragraph = p
            first_textual_index = idx
            break

    # 3. Iteração e Formatação dos Parágrafos
    for idx, p in enumerate(paragraphs):
        _collect_original_typography(p, report)
        norm_upper = _normalize_title(p.text)
        next_p = paragraphs[idx + 1] if idx + 1 < len(paragraphs) else None

        if first_textual_index != -1 and idx >= first_textual_index:
            has_passed_textual_start = True
            if current_section_context == "PRE_TEXTUAL" or current_section_context == "SUMÁRIO":
                current_section_context = "TEXTUAL"

        role, warning_msg = classify_paragraph(
            p,
            current_section_context=current_section_context,
            has_passed_textual_start=has_passed_textual_start,
            next_paragraph=next_p,
        )

        if warning_msg:
            report.add_warning("Classificação Ambígua (Revisão Manual)", p.text, warning_msg)

        if role == ParagraphRole.PAGE_BREAK_EMPTY:
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.0
            _format_paragraph_runs(p, font_name=font_name, font_size_pt=font_size_pt)
            continue

        if role == ParagraphRole.EMPTY:
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = ABNT_RULES.body_line_spacing
            _format_paragraph_runs(p, font_name=font_name, font_size_pt=font_size_pt)
            continue

        # Títulos sem indicativo numérico (Manual p. 11, 12, 13, 17: tamanho 14, negrito, centralizado)
        if role == ParagraphRole.UNNUMBERED_HEADING:
            current_section_context = norm_upper
            if norm_upper == "SUMÁRIO":
                sumario_heading_paragraph = p

            prev_p = paragraphs[idx - 1] if idx > 0 else None
            prev_has_page_break = prev_p is not None and has_manual_page_break(prev_p)
            should_break = work_cfg.primary_heading_page_break and idx > 0 and not prev_has_page_break

            _apply_paragraph_layout(
                p,
                alignment=WD_ALIGN_PARAGRAPH.CENTER,
                line_spacing=ABNT_RULES.heading_line_spacing,
                space_before_pt=0.0 if should_break or prev_has_page_break else ABNT_RULES.heading_space_before_pt,
                space_after_pt=ABNT_RULES.heading_space_after_pt,
                page_break_before=should_break,
                keep_with_next=True,
                report=report,
            )
            _format_paragraph_runs(
                p,
                font_name=font_name,
                font_size_pt=ABNT_RULES.heading_primary_font_size_pt,  # Tamanho 14
                bold=True,
                italic=False,
                uppercase=True,
            )
            report.headings_formatted["sem_numeracao_centralizado"] += 1
            continue

        if role in (
            ParagraphRole.HEADING_1,
            ParagraphRole.HEADING_2,
            ParagraphRole.HEADING_3,
            ParagraphRole.HEADING_4,
            ParagraphRole.HEADING_5,
        ):
            has_passed_textual_start = True
            current_section_context = "TEXTUAL"

        # Seção Primária (Manual p. 19: tamanho 14, maiúsculas, negrito)
        if role == ParagraphRole.HEADING_1:
            prev_p = paragraphs[idx - 1] if idx > 0 else None
            prev_has_page_break = prev_p is not None and has_manual_page_break(prev_p)
            should_break_page = work_cfg.primary_heading_page_break and idx > 0 and not prev_has_page_break

            _apply_paragraph_layout(
                p,
                alignment=WD_ALIGN_PARAGRAPH.LEFT,
                line_spacing=ABNT_RULES.heading_line_spacing,
                space_before_pt=0.0 if should_break_page or prev_has_page_break else ABNT_RULES.heading_space_before_pt,
                space_after_pt=ABNT_RULES.heading_space_after_pt,
                page_break_before=should_break_page,
                keep_with_next=True,
                report=report,
            )
            _format_paragraph_runs(
                p,
                font_name=font_name,
                font_size_pt=ABNT_RULES.heading_primary_font_size_pt,  # Tamanho 14
                bold=True,
                italic=False,
                uppercase=True,
            )
            report.headings_formatted["primaria_1"] += 1

        elif role == ParagraphRole.HEADING_2:
            # Seção secundária (Manual p. 19: tamanho 12, negrito)
            _apply_paragraph_layout(
                p,
                alignment=WD_ALIGN_PARAGRAPH.LEFT,
                line_spacing=ABNT_RULES.heading_line_spacing,
                space_before_pt=ABNT_RULES.heading_space_before_pt,
                space_after_pt=ABNT_RULES.heading_space_after_pt,
                keep_with_next=True,
                report=report,
            )
            _format_paragraph_runs(
                p,
                font_name=font_name,
                font_size_pt=ABNT_RULES.heading_secondary_font_size_pt,  # Tamanho 12
                bold=True,
                italic=False,
                uppercase=False,
            )
            report.headings_formatted["secundaria_1_1"] += 1

        elif role == ParagraphRole.HEADING_3:
            # Seção terciária (Manual p. 19: tamanho 12)
            _apply_paragraph_layout(
                p,
                alignment=WD_ALIGN_PARAGRAPH.LEFT,
                line_spacing=ABNT_RULES.heading_line_spacing,
                space_before_pt=ABNT_RULES.heading_space_before_pt,
                space_after_pt=ABNT_RULES.heading_space_after_pt,
                keep_with_next=True,
                report=report,
            )
            _format_paragraph_runs(
                p,
                font_name=font_name,
                font_size_pt=ABNT_RULES.heading_secondary_font_size_pt,  # Tamanho 12
                bold=True,
                italic=False,
                uppercase=False,
            )
            report.headings_formatted["terciaria_1_1_1"] += 1

        elif role in (ParagraphRole.HEADING_4, ParagraphRole.HEADING_5):
            _apply_paragraph_layout(
                p,
                alignment=WD_ALIGN_PARAGRAPH.LEFT,
                line_spacing=ABNT_RULES.heading_line_spacing,
                space_before_pt=ABNT_RULES.heading_space_before_pt,
                space_after_pt=ABNT_RULES.heading_space_after_pt,
                keep_with_next=True,
                report=report,
            )
            _format_paragraph_runs(
                p,
                font_name=font_name,
                font_size_pt=ABNT_RULES.heading_secondary_font_size_pt,
                bold=True,
                italic=(role == ParagraphRole.HEADING_4),
                uppercase=False,
            )
            report.headings_formatted["quaternaria_quinaria"] += 1

        elif role == ParagraphRole.PRE_TEXTUAL_CENTERED:
            current_after = p.paragraph_format.space_after.pt if p.paragraph_format.space_after else 0.0
            target_after = current_after if current_after > 18.0 else 0.0
            current_before = p.paragraph_format.space_before.pt if p.paragraph_format.space_before else 0.0

            _apply_paragraph_layout(
                p,
                alignment=WD_ALIGN_PARAGRAPH.CENTER,
                line_spacing=ABNT_RULES.body_line_spacing,
                space_before_pt=current_before,
                space_after_pt=target_after,
                report=report,
            )
            report.pre_textual_preserved += 1

        elif role == ParagraphRole.NATUREZA_TRABALHO:
            # Natureza do trabalho (Manual p. 10 e p. 44: tamanho 10, simples, recuo 7.5 cm)
            _apply_paragraph_layout(
                p,
                alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                line_spacing=ABNT_RULES.natureza_line_spacing,
                space_before_pt=0.0,
                space_after_pt=12.0,
                left_indent_cm=ABNT_RULES.natureza_left_indent_cm,
                first_line_indent_cm=0.0,
                report=report,
            )
            _format_paragraph_runs(p, font_name=font_name, font_size_pt=ABNT_RULES.natureza_font_size_pt)
            report.pre_textual_preserved += 1

        elif role == ParagraphRole.ABSTRACT_BODY:
            # Resumo (Manual p. 12: parágrafo único, justificado, espaçamento simples)
            _apply_paragraph_layout(
                p,
                alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                line_spacing=work_cfg.abstract_line_spacing,
                space_before_pt=0.0,
                space_after_pt=12.0,
                left_indent_cm=0.0,
                first_line_indent_cm=0.0,
                report=report,
            )
            _format_paragraph_runs(p, font_name=font_name, font_size_pt=font_size_pt)
            report.paragraphs_body_formatted += 1

        elif role == ParagraphRole.KEYWORDS:
            # Palavras-chave (Manual p. 12: logo abaixo do texto, espaçamento simples)
            _apply_paragraph_layout(
                p,
                alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                line_spacing=1.0,
                space_before_pt=6.0,
                space_after_pt=12.0,
                left_indent_cm=0.0,
                first_line_indent_cm=0.0,
                report=report,
            )
            _format_paragraph_runs(p, font_name=font_name, font_size_pt=font_size_pt)

        elif role == ParagraphRole.LONG_QUOTE:
            # Citações longas > 3 linhas (Manual p. 19 e p. 27: recuo 4 cm, tamanho 10, simples, 18 pt antes e depois)
            _apply_paragraph_layout(
                p,
                alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                line_spacing=ABNT_RULES.long_quote_line_spacing,
                space_before_pt=ABNT_RULES.long_quote_space_before_pt,
                space_after_pt=ABNT_RULES.long_quote_space_after_pt,
                left_indent_cm=ABNT_RULES.long_quote_left_indent_cm,
                first_line_indent_cm=ABNT_RULES.long_quote_first_line_indent_cm,
                report=report,
            )
            _format_paragraph_runs(
                p,
                font_name=font_name,
                font_size_pt=ABNT_RULES.long_quote_font_size_pt,
            )
            report.long_quotes_formatted += 1

        elif role == ParagraphRole.CAPTION:
            # Ilustrações e Tabelas (Manual p. 20, 21 e 22: tamanho 10, simples, centralizado)
            _apply_paragraph_layout(
                p,
                alignment=WD_ALIGN_PARAGRAPH.CENTER,
                line_spacing=ABNT_RULES.caption_line_spacing,
                space_before_pt=ABNT_RULES.caption_space_before_pt,
                space_after_pt=ABNT_RULES.caption_space_after_pt,
                left_indent_cm=0.0,
                first_line_indent_cm=0.0,
                report=report,
            )
            _format_paragraph_runs(
                p,
                font_name=font_name,
                font_size_pt=ABNT_RULES.caption_font_size_pt,
            )
            report.captions_formatted += 1

        elif role == ParagraphRole.REFERENCE_ITEM:
            # Referências (Manual p. 17 e p. 29: alinhadas à esquerda, espaçamento simples, 12 pt after)
            _apply_paragraph_layout(
                p,
                alignment=WD_ALIGN_PARAGRAPH.LEFT,
                line_spacing=ABNT_RULES.reference_line_spacing,
                space_before_pt=ABNT_RULES.reference_space_before_pt,
                space_after_pt=ABNT_RULES.reference_space_after_pt,
                left_indent_cm=0.0,
                first_line_indent_cm=0.0,
                report=report,
            )
            _format_paragraph_runs(
                p,
                font_name=font_name,
                font_size_pt=font_size_pt,
            )
            report.references_formatted += 1

        elif role == ParagraphRole.TOC_ENTRY:
            # Sumário (Manual p. 13 e 14: alinhado à esquerda, sem recuo de 1ª linha)
            _apply_paragraph_layout(
                p,
                alignment=WD_ALIGN_PARAGRAPH.LEFT,
                line_spacing=ABNT_RULES.body_line_spacing,
                space_before_pt=0.0,
                space_after_pt=0.0,
                left_indent_cm=0.0,
                first_line_indent_cm=0.0,
                report=report,
            )
            _format_paragraph_runs(p, font_name=font_name, font_size_pt=font_size_pt)

        elif role == ParagraphRole.LIST_ITEM:
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.line_spacing = ABNT_RULES.body_line_spacing
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            _format_paragraph_runs(p, font_name=font_name, font_size_pt=font_size_pt)
            report.list_items_formatted += 1

        else:
            # Texto Corrido Geral (Manual p. 19: tamanho 12, espaçamento 1,5, recuo 1,25 cm, justificado)
            _apply_paragraph_layout(
                p,
                alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                line_spacing=ABNT_RULES.body_line_spacing,
                space_before_pt=ABNT_RULES.body_space_before_pt,
                space_after_pt=ABNT_RULES.body_space_after_pt,
                left_indent_cm=0.0,
                first_line_indent_cm=ABNT_RULES.body_first_line_indent_cm,
                report=report,
            )
            _format_paragraph_runs(p, font_name=font_name, font_size_pt=font_size_pt)
            report.paragraphs_body_formatted += 1

    # Inserção opcional de TOC dinâmico do Word
    if insert_auto_toc and sumario_heading_paragraph is not None:
        attach_toc_field_to_paragraph(sumario_heading_paragraph, font_name, font_size_pt)
        report.toc_inserted = True

    # 4. Paginação (Manual seção 3.6 / NBR 14724):
    # Contagem desde a folha de rosto, exibição a 2 cm da borda superior no canto direito a partir da Introdução
    if first_textual_paragraph is not None and work_cfg.primary_heading_page_break:
        split_ok = insert_section_break_before_paragraph(
            first_textual_paragraph,
            font_name=font_name,
            page_number_size_pt=ABNT_RULES.small_font_size_pt,
        )
        if split_ok:
            report.pagination_split_at_intro = True

    _configure_sections_and_margins(doc, report)

    if len(doc.sections) >= 2 and report.pagination_split_at_intro:
        clear_header_page_numbers(doc.sections[0])
        for sec in doc.sections[1:]:
            add_page_number_to_header(sec, font_name=font_name, font_size_pt=ABNT_RULES.small_font_size_pt)
        report.pagination_configured = True
    else:
        for sec in doc.sections:
            add_page_number_to_header(sec, font_name=font_name, font_size_pt=ABNT_RULES.small_font_size_pt)
        report.pagination_configured = True

    # 5. Tabelas e Notas de Rodapé
    _format_tables(doc, font_name=font_name, report=report)
    report.footnotes_formatted = format_footnotes_xml(doc, font_name=font_name)

    doc.save(output_path)
    return report
