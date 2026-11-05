"""
Centralização das regras da ABNT (NBR 14724:2024, NBR 6024:2012, NBR 6028:2021, NBR 10520:2023, NBR 6023:2025, NBR 6027:2012)
e Manual Institucional de Normalização Acadêmica (2026).
Todos os valores de margem, tamanho de fonte, recuo e espaçamento estão concentrados aqui.
"""

from dataclasses import dataclass, field
from typing import Dict, List


@dataclass
class WorkTypeOverrides:
    """Configurações específicas por tipo de trabalho acadêmico."""
    name: str
    description: str
    primary_heading_page_break: bool = True
    abstract_line_spacing: float = 1.0


@dataclass
class ABNTConfig:
    """Configuração mestre das normas ABNT vigentes e Manual Acadêmico."""

    # --- 1. Página e Margens (NBR 14724:2024 / Manual seção 3.1 e 3.2) ---
    # Papel A4: 21 cm x 29,7 cm
    page_width_cm: float = 21.0
    page_height_cm: float = 29.7
    # Anverso: Superior 3 cm, Esquerda 3 cm, Inferior 2 cm, Direita 2 cm
    margin_top_cm: float = 3.0
    margin_left_cm: float = 3.0
    margin_bottom_cm: float = 2.0
    margin_right_cm: float = 2.0
    header_distance_cm: float = 2.0  # Numeração no canto superior direito a 2 cm da borda superior

    # Verso: Superior 3 cm, Direita 3 cm, Inferior 2 cm, Esquerda 2 cm
    margin_verso_left_cm: float = 2.0
    margin_verso_right_cm: float = 3.0

    # --- 2. Tipografia e Tamanhos de Fonte (Manual 2026 seções 2.2.1, 2.2.3, 3.4 e 3.7) ---
    allowed_fonts: List[str] = field(default_factory=lambda: ["Arial", "Times New Roman"])
    default_font: str = "Arial"  # Padrão institucional do Manual 2026

    # Texto corrido geral
    body_font_size_pt: int = 12

    # Tamanho de destaque (Capa, Folha de Rosto e Títulos Primários conforme Manual 2026)
    # - Capa: Instituição, Autor, Título, Local e Ano em tamanho 14 (Manual p. 9 e p. 42)
    # - Folha de Rosto: Autor e Título em tamanho 14 (Manual p. 10 e p. 44)
    # - Títulos Primários: Seções 1, 2, 3... e sem numeração (RESUMO, SUMÁRIO, REFERÊNCIAS) em tamanho 14 (Manual p. 19)
    heading_primary_font_size_pt: int = 14

    # Subseções secundárias, terciárias e quaternárias (Manual p. 19): tamanho 12
    heading_secondary_font_size_pt: int = 12

    # Elementos menores (Manual seções 3.4, 3.9, 3.10, 3.11): tamanho 10
    # Citações diretas longas (> 3 linhas), notas de rodapé, paginação, natureza do trabalho, orientador, legendas e fontes
    small_font_size_pt: int = 10

    # Resumo acadêmico (Manual p. 12): tamanho 11 ou 12 (padrão 12 institucional/ABNT)
    abstract_font_size_pt: int = 12

    font_color_rgb: tuple = (0, 0, 0)  # Texto predominantemente na cor preta

    # --- 3. Texto Corrido / Parágrafos (NBR 14724:2024 / Manual seção 3.3 e 3.5) ---
    body_line_spacing: float = 1.5
    body_first_line_indent_cm: float = 1.25  # Padrão ABNT / tradicional
    body_space_before_pt: float = 0.0
    body_space_after_pt: float = 0.0

    # --- 4. Citações Diretas Longas > 3 linhas (NBR 10520:2023 / Manual seção 3.4 e 5.2) ---
    # Recuo de 4 cm da margem esquerda, fonte tamanho 10, espaçamento simples, sem aspas, justificado
    # Separadas do texto por um espaço de 1,5 linha (18 pt) antes e depois
    long_quote_left_indent_cm: float = 4.0
    long_quote_first_line_indent_cm: float = 0.0
    long_quote_line_spacing: float = 1.0
    long_quote_font_size_pt: int = 10
    long_quote_space_before_pt: float = 18.0
    long_quote_space_after_pt: float = 18.0
    long_quote_min_chars: int = 150
    long_quote_min_words: int = 24
    long_quote_existing_indent_threshold_cm: float = 2.0

    # --- 5. Títulos e Seções (NBR 6024:2012 / Manual seção 3.5 e 3.7) ---
    # Separados do texto que os sucede por 1 espaço de 1,5 linha (18 pt)
    heading_line_spacing: float = 1.5
    heading_space_before_pt: float = 18.0
    heading_space_after_pt: float = 18.0

    # Títulos sem indicativo numérico (Manual p. 11, 12, 13, 17):
    # Centralizados, em maiúsculas, negrito, tamanho 14
    unnumbered_centered_titles: List[str] = field(default_factory=lambda: [
        "ERRATA",
        "FOLHA DE APROVAÇÃO",
        "DEDICATÓRIA",
        "AGRADECIMENTOS",
        "EPÍGRAFE",
        "RESUMO",
        "ABSTRACT",
        "RESUMEN",
        "RÉSUMÉ",
        "LISTA DE ILUSTRAÇÕES",
        "LISTA DE FIGURAS",
        "LISTA DE QUADROS",
        "LISTA DE TABELAS",
        "LISTA DE ABREVIATURAS E SIGLAS",
        "LISTA DE ABREVIATURAS",
        "LISTA DE SIGLAS",
        "LISTA DE SÍMBOLOS",
        "SUMÁRIO",
        "REFERÊNCIAS",
        "REFERÊNCIAS BIBLIOGRÁFICAS",
        "GLOSSÁRIO",
        "APÊNDICE",
        "APÊNDICES",
        "ANEXO",
        "ANEXOS",
        "ÍNDICE",
    ])

    textual_start_markers: List[str] = field(default_factory=lambda: [
        "INTRODUÇÃO",
        "1 INTRODUÇÃO",
        "1. INTRODUÇÃO",
        "1 - INTRODUÇÃO",
    ])

    # --- 6. Elementos Menores (Legendas, Fontes, Tabelas, Natureza, Referências) ---
    # Ilustrações e Tabelas (Manual seções 3.9 e 3.10):
    # Identificação na parte superior (título), Fonte consultada na parte inferior, tamanho 10, espaçamento simples
    caption_font_size_pt: int = 10
    caption_line_spacing: float = 1.0
    caption_space_before_pt: float = 6.0
    caption_space_after_pt: float = 6.0

    footnote_font_size_pt: int = 10
    footnote_line_spacing: float = 1.0

    table_font_size_pt: int = 10
    table_line_spacing: float = 1.0

    # Natureza do trabalho (Manual p. 10 e p. 44):
    # Do meio da mancha gráfica até a margem direita (recuo de 7.5 cm a 8.0 cm), espaçamento simples, tamanho 10
    natureza_left_indent_cm: float = 7.5
    natureza_font_size_pt: int = 10
    natureza_line_spacing: float = 1.0

    # Referências (NBR 6023:2025 / Manual p. 17 e p. 29):
    # Alinhadas à margem esquerda, espaçamento simples, separadas por 1 linha em branco de espaço simples (12 pt after)
    reference_font_size_pt: int = 12
    reference_line_spacing: float = 1.0
    reference_space_before_pt: float = 0.0
    reference_space_after_pt: float = 12.0

    # --- 7. Upload ---
    max_upload_size_mb: int = 20

    # --- 8. Tipos de Trabalho Acadêmico ---
    work_types: Dict[str, WorkTypeOverrides] = field(default_factory=lambda: {
        "tcc": WorkTypeOverrides(
            name="TCC / Monografia / Dissertação (Manual 2026 / NBR 14724)",
            description="Capa e folha de rosto independentes; seções primárias em nova página; numeração a partir da Introdução.",
            primary_heading_page_break=True,
            abstract_line_spacing=1.0,
        ),
        "artigo": WorkTypeOverrides(
            name="Artigo Científico (Manual seção 4 / NBR 6022)",
            description="Seções primárias contínuas no fluxo textual; estrutura compacta de artigo.",
            primary_heading_page_break=False,
            abstract_line_spacing=1.0,
        ),
        "relatorio": WorkTypeOverrides(
            name="Relatório Técnico / Acadêmico (NBR 10719)",
            description="Seções primárias iniciam em nova página; estrutura técnica padronizada.",
            primary_heading_page_break=True,
            abstract_line_spacing=1.0,
        ),
    })


ABNT_RULES = ABNTConfig()
