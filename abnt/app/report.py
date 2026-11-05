"""
Estruturas de dados para o relatório "Antes e Depois" e avisos de revisão manual.
"""

from dataclasses import dataclass, field, asdict
from typing import Dict, List, Set


@dataclass
class FormattingReport:
    """Armazena todas as estatísticas e alterações realizadas durante a formatação ABNT."""

    font_applied: str = "Times New Roman"
    font_size_applied: int = 12
    work_type: str = "tcc"

    original_fonts: Set[str] = field(default_factory=set)
    original_font_sizes: Set[float] = field(default_factory=set)

    sections_margins_adjusted: int = 0
    paragraphs_total: int = 0
    paragraphs_body_formatted: int = 0
    paragraphs_spacing_adjusted: int = 0
    paragraphs_indent_adjusted: int = 0
    paragraphs_alignment_adjusted: int = 0

    headings_formatted: Dict[str, int] = field(default_factory=lambda: {
        "primaria_1": 0,
        "secundaria_1_1": 0,
        "terciaria_1_1_1": 0,
        "quaternaria_quinaria": 0,
        "sem_numeracao_centralizado": 0,
    })

    long_quotes_formatted: int = 0
    references_formatted: int = 0
    captions_formatted: int = 0
    tables_formatted: int = 0
    footnotes_formatted: int = 0
    list_items_formatted: int = 0
    pre_textual_preserved: int = 0
    toc_inserted: bool = False
    pagination_configured: bool = False
    pagination_split_at_intro: bool = False

    changes_summary: List[str] = field(default_factory=list)
    warnings: List[Dict[str, str]] = field(default_factory=list)

    def add_warning(self, category: str, excerpt: str, reason: str) -> None:
        clean_excerpt = excerpt.strip()
        if len(clean_excerpt) > 110:
            clean_excerpt = clean_excerpt[:107] + "..."
        self.warnings.append({
            "category": category,
            "excerpt": clean_excerpt,
            "reason": reason,
        })

    def build_summary(self) -> None:
        self.changes_summary.clear()

        if self.sections_margins_adjusted > 0:
            self.changes_summary.append(
                f"Papel ajustado para A4 (21 x 29,7 cm) e margens padronizadas (Superior: 3 cm, Esquerda: 3 cm, Inferior: 2 cm, Direita: 2 cm) em {self.sections_margins_adjusted} seção(ões)."
            )

        orig_fonts_list = sorted(f for f in self.original_fonts if f)
        orig_sizes_list = sorted(s for s in self.original_font_sizes if s)
        if orig_fonts_list or orig_sizes_list:
            fonts_str = ", ".join(orig_fonts_list) if orig_fonts_list else "Padrão do tema"
            sizes_str = ", ".join(f"{s:g} pt" for s in orig_sizes_list) if orig_sizes_list else "variado"
            self.changes_summary.append(
                f"Fonte padronizada de [{fonts_str} ({sizes_str})] para {self.font_applied} {self.font_size_applied} pt (texto corrido) e 10 pt (elementos menores), cor preta."
            )
        else:
            self.changes_summary.append(
                f"Fonte padronizada para {self.font_applied} {self.font_size_applied} pt (texto corrido) e 10 pt (elementos menores), cor preta."
            )

        if self.paragraphs_spacing_adjusted > 0:
            self.changes_summary.append(
                f"{self.paragraphs_spacing_adjusted} parágrafo(s) tiveram o espaçamento entre linhas e/ou espaçamento antes/depois corrigidos."
            )

        if self.paragraphs_indent_adjusted > 0:
            self.changes_summary.append(
                f"{self.paragraphs_indent_adjusted} parágrafo(s) de texto corrido receberam recuo de primeira linha de 1,25 cm."
            )

        if self.paragraphs_alignment_adjusted > 0:
            self.changes_summary.append(
                f"{self.paragraphs_alignment_adjusted} parágrafo(s) tiveram o alinhamento ajustado conforme a função (justificado no texto, centralizado/esquerda em títulos e referências)."
            )

        total_headings = sum(self.headings_formatted.values())
        if total_headings > 0:
            details = []
            if self.headings_formatted["primaria_1"]:
                details.append(f"{self.headings_formatted['primaria_1']} primária(s)")
            if self.headings_formatted["secundaria_1_1"]:
                details.append(f"{self.headings_formatted['secundaria_1_1']} secundária(s)")
            if self.headings_formatted["terciaria_1_1_1"]:
                details.append(f"{self.headings_formatted['terciaria_1_1_1']} terciária(s)")
            if self.headings_formatted["quaternaria_quinaria"]:
                details.append(f"{self.headings_formatted['quaternaria_quinaria']} quaternária(s)/quinária(s)")
            if self.headings_formatted["sem_numeracao_centralizado"]:
                details.append(f"{self.headings_formatted['sem_numeracao_centralizado']} sem indicativo numérico")
            self.changes_summary.append(
                f"{total_headings} título(s) de seção formatados pela NBR 6024 ({', '.join(details)})."
            )

        if self.long_quotes_formatted > 0:
            self.changes_summary.append(
                f"{self.long_quotes_formatted} citação(ões) direta(s) longa(s) formatada(s) com recuo esquerdo de 4 cm, fonte 10 pt e espaçamento simples (NBR 10520)."
            )

        if self.references_formatted > 0:
            self.changes_summary.append(
                f"{self.references_formatted} item(ns) na seção de Referências ajustado(s) com alinhamento à esquerda, espaçamento simples e 1 linha em branco entre referências (NBR 6023)."
            )

        if self.captions_formatted > 0:
            self.changes_summary.append(
                f"{self.captions_formatted} legenda(s)/fonte(s) de ilustrações e tabelas formatada(s) em tamanho 10 pt, espaçamento simples."
            )

        if self.tables_formatted > 0:
            self.changes_summary.append(
                f"{self.tables_formatted} tabela(s) tiveram a tipografia interna ajustada para {self.font_applied} 10 pt e espaçamento simples, preservando bordas e células."
            )

        if self.footnotes_formatted > 0:
            self.changes_summary.append(
                f"{self.footnotes_formatted} nota(s) de rodapé ajustada(s) para {self.font_applied} 10 pt, espaçamento simples e alinhamento justificado."
            )

        if self.pagination_configured:
            if self.pagination_split_at_intro:
                self.changes_summary.append(
                    "Paginação configurada no canto superior direito (algarismos arábicos): oculta nas páginas pré-textuais e exibida a partir da Introdução (contando desde o início)."
                )
            else:
                self.changes_summary.append(
                    "Paginação configurada no canto superior direito em algarismos arábicos."
                )

        if self.toc_inserted:
            self.changes_summary.append(
                "Campo de Sumário Automático (TOC do Word) configurado para atualização dinâmica."
            )

    def to_dict(self) -> dict:
        self.build_summary()
        data = asdict(self)
        data["original_fonts"] = sorted(f for f in self.original_fonts if f)
        data["original_font_sizes"] = sorted(s for s in self.original_font_sizes if s)
        return data

