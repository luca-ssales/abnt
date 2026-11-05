"""
Gera dois documentos .docx para a suíte de testes:
1. sample_baguncado.docx: documento com fontes misturadas, margens erradas, espaçamento 1.0,
   tabelas, hiperlinks e citação longa.
2. sample_correto.docx: documento já dentro dos padrões ABNT.
"""

from pathlib import Path
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def add_hyperlink(paragraph, url, text, color="0000FF", underline=True):
    """Adiciona um hiperlink nativo do Word no parágrafo."""
    part = paragraph.part
    r_id = part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)

    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)

    new_run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")

    if color:
        c = OxmlElement("w:color")
        c.set(qn("w:val"), color)
        rPr.append(c)

    if underline:
        u = OxmlElement("w:u")
        u.set(qn("w:val"), "single")
        rPr.append(u)

    new_run.append(rPr)
    text_node = OxmlElement("w:t")
    text_node.text = text
    new_run.append(text_node)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)


def create_sample_documents(fixtures_dir: Path):
    fixtures_dir.mkdir(parents=True, exist_ok=True)

    # -------------------------------------------------------------
    # 1. DOCUMENTO BAGUNÇADO
    # -------------------------------------------------------------
    doc_bad = Document()

    # Margens propositalmente erradas (1 polegada = 2.54cm em tudo)
    for sec in doc_bad.sections:
        sec.top_margin = Inches(1)
        sec.bottom_margin = Inches(1)
        sec.left_margin = Inches(1)
        sec.right_margin = Inches(1)

    p_capa = doc_bad.add_paragraph("UNIVERSIDADE EXEMPLO")
    p_capa.alignment = WD_ALIGN_PARAGRAPH.CENTER

    p_titulo = doc_bad.add_paragraph("TÍTULO DO TRABALHO ACADÊMICO")
    p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER

    p_sec1 = doc_bad.add_paragraph("1 INTRODUÇÃO")
    # Sem negrito, fonte Comic Sans ou Calibri
    for r in p_sec1.runs:
        r.font.name = "Calibri"
        r.font.size = Pt(11)

    p_body1 = doc_bad.add_paragraph(
        "Este é o primeiro parágrafo de introdução do trabalho. Ele foi escrito originalmente em fonte Calibri tamanho 11, com espaçamento simples e sem nenhum recuo na primeira linha, violando as regras da ABNT NBR 14724. "
        "A integridade deste parágrafo deve ser rigorosamente mantida após a formatação pelo sistema."
    )
    p_body1.paragraph_format.line_spacing = 1.0
    p_body1.paragraph_format.first_line_indent = Inches(0)
    for r in p_body1.runs:
        r.font.name = "Calibri"
        r.font.size = Pt(11)

    # Hiperlink no segundo parágrafo
    p_body2 = doc_bad.add_paragraph("Mais detalhes sobre as normas podem ser consultados no portal oficial: ")
    add_hyperlink(p_body2, "https://www.abnt.org.br", "Portal ABNT")
    p_body2.add_run(". O texto continua normalmente após o link mantendo toda a integridade.")

    # Citação direta longa (> 3 linhas) com recuo pré-existente
    p_quote = doc_bad.add_paragraph(
        "A ciência moderna desenvolveu-se a partir da conjunção de métodos empíricos rigorosos e da formulação matemática dos fenômenos naturais. "
        "Tal transformação exigiu que os pesquisadores adotassem uma postura sistemática de verificação e falseabilidade perante qualquer hipótese levantada, "
        "garantindo assim que o conhecimento produzido pudesse ser replicado independentemente das crenças pessoais de quem o formulou (SILVA, 2021, p. 45)."
    )
    p_quote.paragraph_format.left_indent = Cm(3.0)
    for r in p_quote.runs:
        r.font.name = "Arial"
        r.font.size = Pt(9)

    # Seção secundária
    p_sec2 = doc_bad.add_paragraph("1.1 Objetivos da Pesquisa")
    p_sec2.runs[0].font.name = "Calibri"

    p_body3 = doc_bad.add_paragraph("O objetivo principal desta pesquisa é demonstrar a viabilidade de formatação automatizada.")

    # Legenda e Tabela
    p_leg = doc_bad.add_paragraph("Tabela 1 - Comparativo de desempenho")
    table = doc_bad.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "Critério"
    table.cell(0, 1).text = "Resultado"
    table.cell(1, 0).text = "Preservação de Conteúdo"
    table.cell(1, 1).text = "100%"

    p_fonte = doc_bad.add_paragraph("Fonte: Elaborado pelos autores (2025).")

    # Referências
    p_ref_title = doc_bad.add_paragraph("REFERÊNCIAS")
    p_ref1 = doc_bad.add_paragraph("ASSOCIAÇÃO BRASILEIRA DE NORMAS TÉCNICAS. NBR 14724: informação e documentação: trabalhos acadêmicos: apresentação. Rio de Janeiro: ABNT, 2011.")
    p_ref2 = doc_bad.add_paragraph("SILVA, João. Metodologia científica aplicada à computação. São Paulo: Editora Acadêmica, 2021.")

    bad_path = fixtures_dir / "sample_baguncado.docx"
    doc_bad.save(str(bad_path))

    # -------------------------------------------------------------
    # 2. DOCUMENTO CORRETO
    # -------------------------------------------------------------
    doc_good = Document()

    for sec in doc_good.sections:
        sec.top_margin = Cm(3.0)
        sec.left_margin = Cm(3.0)
        sec.bottom_margin = Cm(2.0)
        sec.right_margin = Cm(2.0)

    p_good_intro = doc_good.add_paragraph("1 INTRODUÇÃO")
    p_good_intro.runs[0].bold = True
    p_good_intro.runs[0].font.name = "Times New Roman"
    p_good_intro.runs[0].font.size = Pt(12)

    p_good_body = doc_good.add_paragraph(
        "Este é um documento que já se encontra nas conformidades da ABNT. "
        "Ele utiliza fonte Times New Roman tamanho 12, espaçamento entre linhas de 1,5 e recuo de 1,25 cm na primeira linha do parágrafo."
    )
    p_good_body.paragraph_format.line_spacing = 1.5
    p_good_body.paragraph_format.first_line_indent = Cm(1.25)
    p_good_body.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    for r in p_good_body.runs:
        r.font.name = "Times New Roman"
        r.font.size = Pt(12)

    good_path = fixtures_dir / "sample_correto.docx"
    doc_good.save(str(good_path))

    return bad_path, good_path


if __name__ == "__main__":
    fixtures = Path(__file__).resolve().parent / "fixtures"
    create_sample_documents(fixtures)
    print("Fixtures de teste geradas com sucesso!")

