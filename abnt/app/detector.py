"""
Módulo de detecção e classificação estrutural de parágrafos segundo a ABNT
(NBR 14724:2024, NBR 6024:2012, NBR 10520:2023, NBR 6028:2021, NBR 6023:2025).
Importante: Este módulo apenas lê e classifica cada parágrafo, preservando 100% do texto.
"""

import re
from enum import Enum
from typing import Optional, Tuple
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn

from app.abnt_rules import ABNT_RULES


class ParagraphRole(str, Enum):
    EMPTY = "empty"
    PAGE_BREAK_EMPTY = "page_break_empty"          # Parágrafo em branco que contém quebra de página manual
    PRE_TEXTUAL_CENTERED = "pre_textual_centered"  # Elementos da Capa / Folha de rosto
    NATUREZA_TRABALHO = "natureza_trabalho"        # Natureza do trabalho (do meio da mancha até a margem direita)
    UNNUMBERED_HEADING = "unnumbered_heading"      # RESUMO, ABSTRACT, SUMÁRIO, REFERÊNCIAS...
    HEADING_1 = "heading_1"                        # 1 INTRODUÇÃO
    HEADING_2 = "heading_2"                        # 1.1 OBJETIVOS
    HEADING_3 = "heading_3"                        # 1.1.1 Objetivo específico
    HEADING_4 = "heading_4"                        # 1.1.1.1 ...
    HEADING_5 = "heading_5"                        # 1.1.1.1.1 ...
    ABSTRACT_BODY = "abstract_body"                # Texto do Resumo / Abstract (parágrafo único, espaçamento simples)
    KEYWORDS = "keywords"                          # Palavras-chave / Keywords
    LONG_QUOTE = "long_quote"                      # Citação direta longa (> 3 linhas)
    CAPTION = "caption"                            # Identificação/Título (acima) ou Fonte/Nota (abaixo) de figuras/tabelas
    REFERENCE_ITEM = "reference_item"              # Item de referência bibliográfica (NBR 6023:2025)
    TOC_ENTRY = "toc_entry"                        # Entrada existente do sumário manual/automático
    LIST_ITEM = "list_item"                        # Item de lista com marcador ou numeração automática
    BODY_TEXT = "body_text"                        # Texto corrido padrão


# Expressão regular para detectar seções numeradas (NBR 6024:2012):
# "1 INTRODUÇÃO", "1.1 Metodologia", "2.1.1 Aspectos".
# Exige um espaço entre o número e o texto (sem ponto ou hífen após o último algarismo).
NUMBERED_HEADING_RE = re.compile(
    r"^\s*(?P<num>\d+(?:\.\d+){0,4})\.?\s+(?P<title>[A-ZÁÉÍÓÚÂÊÔÃÕÇ][^\n]{1,115})$"
)

# Legendas e fontes de ilustrações e tabelas (NBR 14724:2024, 5.8 e 5.9)
CAPTION_RE = re.compile(
    r"^\s*(?:Figura|Tabela|Quadro|Gráfico|Ilustração|Fotografia|Organograma|Fluxograma|Mapa|Diagrama)\s+\d+\s*[\-–—:.]|^\s*Fonte\s*:",
    re.IGNORECASE,
)

# Linhas de sumário manual com pontilhados ou espaçamento e número de página ao final
TOC_LINE_RE = re.compile(r"^.{3,120}(?:\.{2,}|\s{3,})\s*\d+\s*$")

# Padrão de palavras-chave (NBR 6028:2021)
KEYWORDS_RE = re.compile(r"^\s*(?:Palavras[\- ]chave|Keywords|Palabras clave|Mots[\- ]clés)\s*:", re.IGNORECASE)

# Padrão textual de Natureza do Trabalho na Folha de Rosto
NATUREZA_KEYWORDS_RE = re.compile(
    r"(?:trabalho acad[eê]mico|requisito parcial|disserta[cç][aã]o|tese|monografia|projeto de gradua[cç][aã]o|orientador|professor)\b",
    re.IGNORECASE,
)


def _normalize_title(text: str) -> str:
    cleaned = re.sub(r"\s+", " ", text.strip()).upper()
    cleaned = cleaned.rstrip(".:;-–—")
    return cleaned


def has_manual_page_break(paragraph: Paragraph) -> bool:
    """Verifica se o parágrafo contém quebra de página manual (<w:br w:type='page'/>)."""
    return any(
        el.get(qn("w:type")) == "page"
        for el in paragraph._p.iter(qn("w:br"))
    )


def is_list_paragraph(paragraph: Paragraph) -> bool:
    pPr = paragraph._p.pPr
    if pPr is not None and pPr.find(qn("w:numPr")) is not None:
        return True
    style_name = (paragraph.style.name or "").lower() if paragraph.style else ""
    if any(k in style_name for k in ("list", "lista", "bullet", "marcador")):
        return True
    stripped = paragraph.text.lstrip()
    if stripped.startswith(("• ", "- ", "– ", "* ")):
        return True
    return False


def get_effective_left_indent_cm(paragraph: Paragraph) -> float:
    indent = paragraph.paragraph_format.left_indent
    if indent is None and paragraph.style and paragraph.style.paragraph_format:
        indent = paragraph.style.paragraph_format.left_indent
    if indent is not None:
        return float(indent.cm)
    return 0.0


def is_all_italic(paragraph: Paragraph) -> bool:
    non_empty_runs = [r for r in paragraph.runs if r.text and r.text.strip()]
    if not non_empty_runs:
        return False
    return all(bool(r.italic) for r in non_empty_runs)


def is_all_bold(paragraph: Paragraph) -> bool:
    non_empty_runs = [r for r in paragraph.runs if r.text and r.text.strip()]
    if not non_empty_runs:
        return False
    return all(bool(r.bold) for r in non_empty_runs)


def classify_paragraph(
    paragraph: Paragraph,
    current_section_context: str,
    has_passed_textual_start: bool,
    next_paragraph: Optional[Paragraph] = None,
) -> Tuple[ParagraphRole, Optional[str]]:
    """
    Classifica um parágrafo do documento retornando (ParagraphRole, warning_message_se_ambiguo).
    Preserva rigorosamente a integridade do texto sem modificá-lo.
    """
    raw_text = paragraph.text
    text = raw_text.strip()
    if not text:
        if has_manual_page_break(paragraph):
            return ParagraphRole.PAGE_BREAK_EMPTY, None
        return ParagraphRole.EMPTY, None

    style_name = (paragraph.style.name or "").lower() if paragraph.style else ""
    norm_upper = _normalize_title(text)

    # 1. Títulos sem indicativo numérico (NBR 14724:2024 / NBR 6024:2012)
    # Apenas se for uma linha curta isolada de título e NÃO uma entrada com número de página do sumário!
    if norm_upper in ABNT_RULES.unnumbered_centered_titles and not TOC_LINE_RE.match(text):
        return ParagraphRole.UNNUMBERED_HEADING, None

    # 2. Entradas do Sumário (NBR 6027:2012)
    if current_section_context == "SUMÁRIO" or style_name.startswith(("toc", "sumário")):
        # Se for linha pontilhada ou se o contexto for sumário e não for o próprio título da introdução real
        if TOC_LINE_RE.match(text) or (current_section_context == "SUMÁRIO" and not norm_upper in ABNT_RULES.textual_start_markers):
            return ParagraphRole.TOC_ENTRY, None

    # 3. Referências Bibliográficas (NBR 6023:2025)
    if current_section_context in ("REFERÊNCIAS", "REFERÊNCIAS BIBLIOGRÁFICAS"):
        if len(text) < 120 and norm_upper in ABNT_RULES.unnumbered_centered_titles and not TOC_LINE_RE.match(text):
            return ParagraphRole.UNNUMBERED_HEADING, None
        return ParagraphRole.REFERENCE_ITEM, None

    # 4. Palavras-chave / Keywords (NBR 6028:2021)
    if KEYWORDS_RE.match(text):
        return ParagraphRole.KEYWORDS, None

    # 5. Estilos de Título Nativos do Word (Heading 1..5)
    for level, role in [
        ("1", ParagraphRole.HEADING_1),
        ("2", ParagraphRole.HEADING_2),
        ("3", ParagraphRole.HEADING_3),
        ("4", ParagraphRole.HEADING_4),
        ("5", ParagraphRole.HEADING_5),
    ]:
        if style_name in (f"heading {level}", f"título {level}", f"titulo {level}"):
            return role, None

    # 6. Legendas e Fontes de Figuras, Tabelas, Quadros e Gráficos
    if "caption" in style_name or "legenda" in style_name or CAPTION_RE.match(text):
        if len(text) <= 350:
            return ParagraphRole.CAPTION, None

    # 7. Detecção de Títulos Numerados por Heurística (NBR 6024:2012)
    # Exige que não seja entrada pontilhada de sumário e que não termine em pontuação de parágrafo comum
    if len(text) <= 120 and not text.endswith((".", ";", ",", "!", "?")) and not TOC_LINE_RE.match(text):
        match = NUMBERED_HEADING_RE.match(text)
        if match:
            num_str = match.group("num")
            title_part = match.group("title").strip()
            parts = num_str.split(".")
            level = len(parts)

            first_word = title_part.split()[0].lower() if title_part.split() else ""
            false_positive_words = {
                "de", "do", "da", "dos", "das", "em", "no", "na", "para", "por",
                "a", "o", "e", "ou", "anos", "dias", "meses", "horas", "km", "mg", "kg", "ml"
            }

            is_emphasized = is_all_bold(paragraph) or text.isupper() or level >= 2
            if first_word not in false_positive_words and is_emphasized:
                role_map = {
                    1: ParagraphRole.HEADING_1,
                    2: ParagraphRole.HEADING_2,
                    3: ParagraphRole.HEADING_3,
                    4: ParagraphRole.HEADING_4,
                    5: ParagraphRole.HEADING_5,
                }
                return role_map.get(min(level, 5), ParagraphRole.HEADING_5), None
            elif first_word not in false_positive_words and not is_emphasized:
                return (
                    ParagraphRole.BODY_TEXT,
                    "Linha iniciada por número sem destaque visual; mantida como texto corrido para evitar falso título.",
                )

    # 8. Elementos Pré-Textuais (Capa, Folha de Rosto, Natureza do Trabalho)
    if not has_passed_textual_start:
        left_indent_cm = get_effective_left_indent_cm(paragraph)
        # Natureza do trabalho detectada por indentação prévia (> 5cm) ou por termos acadêmicos típicos
        is_natureza_text = bool(NATUREZA_KEYWORDS_RE.search(text))
        if left_indent_cm >= 5.0 or is_natureza_text:
            return ParagraphRole.NATUREZA_TRABALHO, None

        if current_section_context in ("RESUMO", "ABSTRACT", "RESUMEN", "RÉSUMÉ"):
            return ParagraphRole.ABSTRACT_BODY, None

        if current_section_context == "PRE_TEXTUAL":
            return ParagraphRole.PRE_TEXTUAL_CENTERED, None

    # Se estiver dentro de RESUMO / ABSTRACT
    if current_section_context in ("RESUMO", "ABSTRACT", "RESUMEN", "RÉSUMÉ"):
        return ParagraphRole.ABSTRACT_BODY, None

    # 9. Listas com marcadores ou numeração
    if is_list_paragraph(paragraph):
        return ParagraphRole.LIST_ITEM, None

    # 10. Citações Diretas Longas (> 3 linhas — NBR 10520:2023)
    left_indent_cm = get_effective_left_indent_cm(paragraph)
    is_quote_style = any(q in style_name for q in ("quote", "citação", "citacao", "block text", "recuo"))

    if is_quote_style or left_indent_cm >= ABNT_RULES.long_quote_existing_indent_threshold_cm:
        return ParagraphRole.LONG_QUOTE, None

    # Heurística avançada para citação longa desformatada:
    # Parágrafo isolado que começa e termina com aspas ou todo em itálico,
    # com mais de 3 linhas (~24 palavras ou ~150 caracteres), seguido de fonte/ano
    has_quotes = (
        (text.startswith(('"', '“', "'")) and text.endswith(('"', '”', "'")))
        or (text.startswith(('"', '“')) and bool(re.search(r"\(\s*[A-ZÁÉÍÓÚÂÊÔÃÕÇ][^)]*\d{4}[^)]*\)\s*\.?\s*$", text)))
    )
    is_next_fonte = False
    if next_paragraph and next_paragraph.text.strip():
        is_next_fonte = bool(CAPTION_RE.match(next_paragraph.text.strip()))

    words_count = len(text.split())
    is_long_enough = (words_count >= ABNT_RULES.long_quote_min_words or len(text) >= ABNT_RULES.long_quote_min_chars)

    if (has_quotes or is_all_italic(paragraph)) and is_long_enough:
        if is_next_fonte or has_quotes:
            return ParagraphRole.LONG_QUOTE, None

    # 11. Padrão: Texto Corrido ABNT
    return ParagraphRole.BODY_TEXT, None
