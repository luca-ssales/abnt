"""
Manipulações diretas do XML OpenXML (via lxml / python-docx oxml) para recursos
que a API de alto nível do python-docx não cobre diretamente:
- Numeração de páginas (campo PAGE) no cabeçalho superior direito
- Formatação das notas de rodapé (footnotes.xml) preservando texto e referências
- Preservação e formatação visual de hiperlinks (<w:hyperlink>) dentro dos parágrafos
- Inserção opcional de campo de Sumário Automático (TOC) do Word
"""

from docx.document import Document as DocxDocument
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

from app.abnt_rules import ABNT_RULES


def iter_all_Paragraph_runs_including_hyperlinks(paragraph):
    return paragraph._p.xpath(".//w:r")


def apply_font_to_r_element(
    r_element,
    font_name: str,
    font_size_pt: float,
    bold: bool = None,
    italic: bool = None,
    uppercase: bool = None,
    force_black: bool = True,
) -> None:
    rPr = r_element.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        r_element.insert(0, rPr)

    # Fonte
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:ascii"), font_name)
    rFonts.set(qn("w:hAnsi"), font_name)
    rFonts.set(qn("w:cs"), font_name)
    rFonts.set(qn("w:eastAsia"), font_name)
    for theme_attr in ("w:asciiTheme", "w:hAnsiTheme", "w:cstheme", "w:eastAsiaTheme"):
        if qn(theme_attr) in rFonts.attrib:
            del rFonts.attrib[qn(theme_attr)]

    # Tamanho
    half_pts = str(int(round(font_size_pt * 2)))
    for tag_name in ("w:sz", "w:szCs"):
        sz_el = rPr.find(qn(tag_name))
        if sz_el is None:
            sz_el = OxmlElement(tag_name)
            rPr.append(sz_el)
        sz_el.set(qn("w:val"), half_pts)

    # Cor preta
    if force_black:
        color_el = rPr.find(qn("w:color"))
        if color_el is None:
            color_el = OxmlElement("w:color")
            rPr.append(color_el)
        color_el.set(qn("w:val"), "000000")
        for theme_color_attr in ("w:themeColor", "w:themeTint", "w:themeShade"):
            if qn(theme_color_attr) in color_el.attrib:
                del color_el.attrib[qn(theme_color_attr)]

    # Negrito
    if bold is not None:
        for b_tag in ("w:b", "w:bCs"):
            existing_b = rPr.find(qn(b_tag))
            if bold:
                if existing_b is None:
                    existing_b = OxmlElement(b_tag)
                    rPr.append(existing_b)
                existing_b.set(qn("w:val"), "1")
            else:
                if existing_b is not None:
                    rPr.remove(existing_b)

    # Itálico
    if italic is not None:
        for i_tag in ("w:i", "w:iCs"):
            existing_i = rPr.find(qn(i_tag))
            if italic:
                if existing_i is None:
                    existing_i = OxmlElement(i_tag)
                    rPr.append(existing_i)
                existing_i.set(qn("w:val"), "1")
            else:
                if existing_i is not None:
                    rPr.remove(existing_i)

    # Caixa Alta visual sem tocar nos caracteres de <w:t>
    if uppercase is not None:
        existing_caps = rPr.find(qn("w:caps"))
        if uppercase:
            if existing_caps is None:
                caps_el = OxmlElement("w:caps")
                caps_el.set(qn("w:val"), "1")
                rPr.append(caps_el)
            else:
                existing_caps.set(qn("w:val"), "1")
        else:
            if existing_caps is not None:
                rPr.remove(existing_caps)


def add_page_number_to_header(section, font_name: str, font_size_pt: int = 10) -> None:
    header = section.header
    header.is_linked_to_previous = False

    if header.paragraphs:
        p = header.paragraphs[0]
    else:
        p = header.add_paragraph()

    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.0

    existing_instr = p._p.xpath(".//w:instrText[contains(text(), 'PAGE')]")
    if existing_instr:
        for r_el in iter_all_Paragraph_runs_including_hyperlinks(p):
            apply_font_to_r_element(r_el, font_name, font_size_pt, bold=False, italic=False)
        return

    run = p.add_run()
    fldChar_begin = OxmlElement("w:fldChar")
    fldChar_begin.set(qn("w:fldCharType"), "begin")

    instrText = OxmlElement("w:instrText")
    instrText.set(qn("xml:space"), "preserve")
    instrText.text = " PAGE "

    fldChar_separate = OxmlElement("w:fldChar")
    fldChar_separate.set(qn("w:fldCharType"), "separate")

    fldChar_end = OxmlElement("w:fldChar")
    fldChar_end.set(qn("w:fldCharType"), "end")

    run._r.append(fldChar_begin)
    run._r.append(instrText)
    run._r.append(fldChar_separate)
    run._r.append(fldChar_end)

    apply_font_to_r_element(run._r, font_name, font_size_pt, bold=False, italic=False)


def clear_header_page_numbers(section) -> None:
    header = section.header
    header.is_linked_to_previous = False
    for p in header.paragraphs:
        for r in list(p._p.findall(qn("w:r"))):
            if r.xpath(".//w:instrText[contains(text(), 'PAGE')]"):
                p._p.remove(r)


def insert_section_break_before_paragraph(
    paragraph,
    font_name: str,
    page_number_size_pt: int = 10,
) -> bool:
    prev_p = paragraph._p.getprevious()
    while prev_p is not None and prev_p.tag != qn("w:p"):
        prev_p = prev_p.getprevious()

    if prev_p is None:
        return False

    prev_pPr = prev_p.find(qn("w:pPr"))
    if prev_pPr is None:
        prev_pPr = OxmlElement("w:pPr")
        prev_p.insert(0, prev_pPr)

    if prev_pPr.find(qn("w:sectPr")) is not None:
        return False

    sectPr = OxmlElement("w:sectPr")

    pgSz = OxmlElement("w:pgSz")
    pgSz.set(qn("w:w"), str(int(round(ABNT_RULES.page_width_cm * 567))))
    pgSz.set(qn("w:h"), str(int(round(ABNT_RULES.page_height_cm * 567))))
    sectPr.append(pgSz)

    pgMar = OxmlElement("w:pgMar")
    pgMar.set(qn("w:top"), str(int(round(ABNT_RULES.margin_top_cm * 567))))
    pgMar.set(qn("w:bottom"), str(int(round(ABNT_RULES.margin_bottom_cm * 567))))
    pgMar.set(qn("w:left"), str(int(round(ABNT_RULES.margin_left_cm * 567))))
    pgMar.set(qn("w:right"), str(int(round(ABNT_RULES.margin_right_cm * 567))))
    pgMar.set(qn("w:header"), str(int(round(ABNT_RULES.header_distance_cm * 567))))
    pgMar.set(qn("w:footer"), str(int(round(1.25 * 567))))
    pgMar.set(qn("w:gutter"), "0")
    sectPr.append(pgMar)

    sectType = OxmlElement("w:type")
    sectType.set(qn("w:val"), "nextPage")
    sectPr.append(sectType)

    prev_pPr.append(sectPr)
    return True


def format_footnotes_xml(doc: DocxDocument, font_name: str) -> int:
    formatted_count = 0
    package = doc.part.package

    for part in package.parts:
        partname_str = str(part.partname).lower()
        if "footnotes.xml" in partname_str or "endnotes.xml" in partname_str:
            element = getattr(part, "element", None)
            if element is None:
                continue

            notes = element.xpath(".//w:footnote | .//w:endnote")
            for note in notes:
                note_id = note.get(qn("w:id"))
                note_type = note.get(qn("w:type"))
                if note_type in ("separator", "continuationSeparator", "continuationNotice"):
                    continue
                if note_id is not None:
                    try:
                        if int(note_id) <= 0:
                            continue
                    except ValueError:
                        pass

                paragraphs = note.findall(qn("w:p"))
                has_text = False
                for p in paragraphs:
                    pPr = p.find(qn("w:pPr"))
                    if pPr is None:
                        pPr = OxmlElement("w:pPr")
                        p.insert(0, pPr)

                    jc = pPr.find(qn("w:jc"))
                    if jc is None:
                        jc = OxmlElement("w:jc")
                        pPr.append(jc)
                    jc.set(qn("w:val"), "both")

                    spacing = pPr.find(qn("w:spacing"))
                    if spacing is None:
                        spacing = OxmlElement("w:spacing")
                        pPr.append(spacing)
                    spacing.set(qn("w:before"), "0")
                    spacing.set(qn("w:after"), "0")
                    spacing.set(qn("w:line"), "240")
                    spacing.set(qn("w:lineRule"), "auto")

                    ind = pPr.find(qn("w:ind"))
                    if ind is None:
                        ind = OxmlElement("w:ind")
                        pPr.append(ind)
                    ind.set(qn("w:left"), "0")
                    ind.set(qn("w:firstLine"), "0")

                    for r in p.xpath(".//w:r"):
                        apply_font_to_r_element(
                            r,
                            font_name=font_name,
                            font_size_pt=ABNT_RULES.footnote_font_size_pt,
                            force_black=True,
                        )
                        if r.findall(qn("w:t")):
                            has_text = True

                if has_text:
                    formatted_count += 1

    return formatted_count


def attach_toc_field_to_paragraph(paragraph, font_name: str, font_size_pt: int = 12) -> None:
    run = paragraph.add_run()
    fldChar_begin = OxmlElement("w:fldChar")
    fldChar_begin.set(qn("w:fldCharType"), "begin")

    instrText = OxmlElement("w:instrText")
    instrText.set(qn("xml:space"), "preserve")
    instrText.text = ' TOC \\o "1-5" \\h \\z \\u '

    fldChar_separate = OxmlElement("w:fldChar")
    fldChar_separate.set(qn("w:fldCharType"), "separate")

    fldChar_end = OxmlElement("w:fldChar")
    fldChar_end.set(qn("w:fldCharType"), "end")

    run._r.append(fldChar_begin)
    run._r.append(instrText)
    run._r.append(fldChar_separate)
    run._r.append(fldChar_end)

    apply_font_to_r_element(run._r, font_name, font_size_pt, bold=False, italic=False)

