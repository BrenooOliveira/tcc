#!/usr/bin/env python3
"""
monografia_converter.py
=======================
Conversor de Monografia Markdown → DOCX com conformidade ABNT (NBR 14724 / NBR 10520:2023 / NBR 6023).

Uso:
    python monografia_converter.py monografia.md
    python monografia_converter.py monografia.md --output saida.docx --font Times

Saídas:
    • <output>.docx               — documento final formatado
    • relatorio_pendencias.md     — pendências bibliográficas e de formatação
    • terminal                    — resumo das alterações aplicadas

Arquitetura (OOP):
    CONFIG              → parâmetros centralizados (margens, fontes, tamanhos)
    ABNTStyleManager    → registra e retorna estilos python-docx
    MarkdownParser      → transforma .md em lista de Block dataclasses
    AcademicValidator   → valida regras ABNT, popula lista de pendências
    DocxGenerator       → consome blocks e estilos, constrói o .docx
    ReportGenerator     → escreve relatorio_pendencias.md e log de terminal
"""

import re
import sys
import argparse
import os
import shutil
import subprocess
import tempfile
from copy import deepcopy
from dataclasses import dataclass
from pathlib import Path
from typing import Optional
from datetime import datetime

# --------------------------------------------------------------------------- #
# python-docx imports
# --------------------------------------------------------------------------- #
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from figure_registry import FigureRegistry
from omml_math import OmmlError, append_equation

# =========================================================================== #
# CONFIGURAÇÃO CENTRAL — altere aqui sem tocar no resto do código             #
# =========================================================================== #
CONFIG: dict = {
    # --- Margens ABNT NBR 14724 ---
    "margin_top_cm":    3.0,   # superior
    "margin_left_cm":   3.0,   # esquerda
    "margin_bottom_cm": 2.0,   # inferior
    "margin_right_cm":  2.0,   # direita

    # --- Fonte ---
    "font_name":        "Arial",   # "Arial" ou "Times New Roman"
    "font_size_pt":     12,        # corpo do texto
    "font_size_quote_pt": 10,      # citações longas e notas de rodapé
    "font_size_ref_pt": 12,        # referências (mesmo que corpo, mas single-space)

    # --- Espaçamento ---
    "line_spacing_body": 1.5,      # corpo do texto (múltiplo)
    "line_spacing_quote": 1.0,     # citações longas e referências (simples)

    # --- Parágrafos ---
    "para_indent_cm":   1.25,      # recuo de primeira linha (corpo)
    "quote_indent_cm":  4.0,       # recuo esquerdo de citação longa ABNT

    # --- Headings ---
    "heading_bold": True,
    "heading_uppercase_levels": {1, 2},  # H1 e H2 em caixa-alta
    "heading_font_size": {
        1: 12,   # H1 — mesmo tamanho, caixa-alta + negrito
        2: 12,   # H2 — idem
        3: 12,   # H3 — negrito, sem caixa-alta
        4: 12,   # H4 — itálico
    },

    # --- Citações ABNT (sistema autor-data) ---
    # Padrão: (SOBRENOME, ano) ou SOBRENOME (ano)
    "citation_pattern_inline": r'\(([A-ZÁÉÍÓÚÃÕÂÊÔÇÀÜ][A-ZÁÉÍÓÚÃÕÂÊÔÇÀÜa-záéíóúãõâêôçàü\s]+(?:et al\.)?),\s*(\d{4}(?:,\s*p\.\s*[\d\-]+)?)\)',
    "citation_pattern_inline2": r'\b([A-ZÁÉÍÓÚÃÕÂÊÔÇÀÜ]{2,}(?:\s+[A-ZÁÉÍÓÚÃÕÂÊÔÇÀÜ]{2,})?(?:\s+et al\.)?)\s+\((\d{4}(?:,\s*p\.\s*[\d\-]+)?)\)',
}


# =========================================================================== #
# DATACLASSES — representação intermediária dos blocos do markdown            #
# =========================================================================== #
@dataclass
class Block:
    """Base para todos os blocos de conteúdo."""
    pass


@dataclass
class HeadingBlock(Block):
    level: int       # 1–6
    text: str
    raw: str = ""


@dataclass
class ParagraphBlock(Block):
    text: str
    raw: str = ""
    first_line_indent: bool = True


@dataclass
class QuoteBlock(Block):
    """Citação longa ABNT (> texto no markdown)."""
    text: str
    raw: str = ""


@dataclass
class ListItemBlock(Block):
    text: str
    ordered: bool = False
    number: int = 1
    raw: str = ""


@dataclass
class ImageBlock(Block):
    alt: str
    url: str
    raw: str = ""


@dataclass
class TableBlock(Block):
    headers: list[str]
    rows: list[list[str]]
    raw: str = ""
    caption: Optional[str] = None


@dataclass
class MathBlock(Block):
    expr: str
    display: bool = True
    raw: str = ""


@dataclass
class HorizontalRuleBlock(Block):
    pass


@dataclass
class ReferenceEntryBlock(Block):
    """Um item de referência bibliográfica."""
    text: str
    raw: str = ""


@dataclass
class Pendencia:
    """Registro de uma pendência bibliográfica ou de formatação."""
    tipo: str            # ex.: "DADO FALTANTE", "et al.", "URL ausente"
    descricao: str
    trecho: str          # trecho do texto que gerou a pendência
    linha: Optional[int] = None


# =========================================================================== #
# ABNT STYLE MANAGER                                                          #
# =========================================================================== #
class ABNTStyleManager:
    """
    Registra e configura todos os estilos ABNT no documento python-docx.
    Cada método retorna o nome do estilo para uso em DocxGenerator.
    """

    def __init__(self, doc: Document, config: dict):
        self.doc = doc
        self.cfg = config
        self._setup_styles()

    # ------------------------------------------------------------------ #
    def _setup_styles(self):
        """Configura estilos no documento. Chamado uma vez no __init__."""
        self._configure_normal()
        self._configure_headings()
        self._configure_quote()
        self._configure_reference()
        self._configure_figure_caption()
        self._configure_image_placeholder()

    # ------------------------------------------------------------------ #
    def _configure_normal(self):
        """Estilo 'Normal' — corpo do texto ABNT."""
        style = self.doc.styles['Normal']
        font = style.font
        font.name = self.cfg["font_name"]
        font.size = Pt(self.cfg["font_size_pt"])

        pf = style.paragraph_format
        pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        pf.first_line_indent = Cm(self.cfg["para_indent_cm"])
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        pf.line_spacing = self.cfg["line_spacing_body"]

    # ------------------------------------------------------------------ #
    def _configure_headings(self):
        """Estilos Heading 1–4 conforme ABNT NBR 14724."""
        for level in range(1, 5):
            style_name = f'Heading {level}'
            try:
                style = self.doc.styles[style_name]
            except KeyError:
                style = self.doc.styles.add_style(style_name, 1)  # WD_STYLE_TYPE.PARAGRAPH

            font = style.font
            font.name = self.cfg["font_name"]
            font.size = Pt(self.cfg["heading_font_size"].get(level, 12))
            font.bold = True if level <= 2 else (True if level == 3 else False)
            font.italic = (level == 4)
            font.color.rgb = RGBColor(0, 0, 0)

            pf = style.paragraph_format
            pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
            pf.first_line_indent = Cm(0)
            pf.space_before = Pt(12)
            pf.space_after = Pt(6)
            pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
            pf.line_spacing = self.cfg["line_spacing_body"]
            pf.keep_with_next = True

    # ------------------------------------------------------------------ #
    def _configure_quote(self):
        """Estilo 'ABNT Quote' — citações longas (> 3 linhas)."""
        style_name = 'ABNT Quote'
        if style_name not in [s.name for s in self.doc.styles]:
            style = self.doc.styles.add_style(style_name, 1)
        else:
            style = self.doc.styles[style_name]

        font = style.font
        font.name = self.cfg["font_name"]
        font.size = Pt(self.cfg["font_size_quote_pt"])
        font.bold = False
        font.italic = False

        pf = style.paragraph_format
        pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        pf.left_indent = Cm(self.cfg["quote_indent_cm"])
        pf.first_line_indent = Cm(0)
        pf.space_before = Pt(6)
        pf.space_after = Pt(6)
        pf.line_spacing_rule = WD_LINE_SPACING.SINGLE

    # ------------------------------------------------------------------ #
    def _configure_reference(self):
        """Estilo 'ABNT Reference' — entradas bibliográficas."""
        style_name = 'ABNT Reference'
        if style_name not in [s.name for s in self.doc.styles]:
            style = self.doc.styles.add_style(style_name, 1)
        else:
            style = self.doc.styles[style_name]

        font = style.font
        font.name = self.cfg["font_name"]
        font.size = Pt(self.cfg["font_size_ref_pt"])

        pf = style.paragraph_format
        pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        pf.first_line_indent = Cm(0)
        pf.left_indent = Cm(0)
        pf.space_before = Pt(0)
        pf.space_after = Pt(6)
        pf.line_spacing_rule = WD_LINE_SPACING.SINGLE

    # ------------------------------------------------------------------ #
    def _configure_figure_caption(self):
        """Estilo para legendas de figuras e tabelas."""
        style_name = 'ABNT Caption'
        if style_name not in [s.name for s in self.doc.styles]:
            style = self.doc.styles.add_style(style_name, 1)
        else:
            style = self.doc.styles[style_name]

        font = style.font
        font.name = self.cfg["font_name"]
        font.size = Pt(10)

        pf = style.paragraph_format
        pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pf.first_line_indent = Cm(0)
        pf.space_before = Pt(3)
        pf.space_after = Pt(6)
        pf.line_spacing_rule = WD_LINE_SPACING.SINGLE

    # ------------------------------------------------------------------ #
    def _configure_image_placeholder(self):
        """Estilo para placeholders de imagem."""
        style_name = 'ABNT Image Placeholder'
        if style_name not in [s.name for s in self.doc.styles]:
            style = self.doc.styles.add_style(style_name, 1)
        else:
            style = self.doc.styles[style_name]

        font = style.font
        font.name = self.cfg["font_name"]
        font.size = Pt(10)
        font.italic = True
        font.color.rgb = RGBColor(0x80, 0x80, 0x80)

        pf = style.paragraph_format
        pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pf.space_before = Pt(6)
        pf.space_after = Pt(3)


# =========================================================================== #
# MARKDOWN PARSER                                                              #
# =========================================================================== #
class MarkdownParser:
    """
    Faz o parsing do arquivo .md e retorna uma lista de Block dataclasses.

    Blocos reconhecidos:
        HeadingBlock    — # Título
        ParagraphBlock  — texto comum
        QuoteBlock      — > citação longa
        ListItemBlock   — - item ou 1. item
        ImageBlock      — ![alt](url)
        TableBlock      — tabelas markdown pipe-style
        MathBlock       — equações em bloco delimitadas por $$
        $...$           — fórmula inline, preservada dentro do parágrafo
        HorizontalRuleBlock — ---
        ReferenceEntryBlock — parágrafos dentro da seção de referências
    """

    # Regex de bloco
    _RE_HEADING   = re.compile(r'^(#{1,6})\s+(.+)$')
    _RE_QUOTE     = re.compile(r'^>\s?(.*)$')
    _RE_ULIST     = re.compile(r'^[-*]\s+(.+)$')
    _RE_OLIST     = re.compile(r'^(\d+)\.\s+(.+)$')
    _RE_IMAGE     = re.compile(r'^!\[([^\]]*)\]\(([^)]+)\)\s*$')
    _RE_HRULE     = re.compile(r'^---+\s*$')
    _RE_EMPTY     = re.compile(r'^\s*$')
    _RE_TABLE_ROW = re.compile(r'^\s*\|.*\|\s*$')
    _RE_TABLE_SEP = re.compile(r'^\s*\|?\s*:?-{3,}:?\s*(\|\s*:?-{3,}:?\s*)+\|?\s*$')
    _RE_BLOCK_MATH = re.compile(r'^\s*\$\$(.*)\$\$\s*$')

    def __init__(self, md_path: str):
        self.md_path = Path(md_path)
        self.blocks: list[Block] = []
        self._in_references = False
        self._pending_lines: list[str] = []

    # ------------------------------------------------------------------ #
    def parse(self) -> list[Block]:
        """Ponto de entrada. Retorna lista de Blocks."""
        text = self.md_path.read_text(encoding='utf-8')
        lines = text.splitlines()

        i = 0
        while i < len(lines):
            line = lines[i]

            # --- Heading ---
            m = self._RE_HEADING.match(line)
            if m:
                self._flush_pending()
                level = len(m.group(1))
                heading_text = m.group(2).strip()
                blk = HeadingBlock(level=level, text=heading_text, raw=line)
                self.blocks.append(blk)

                # Detectar início da seção de referências
                if re.search(r'REFERÊNCIAS', heading_text, re.IGNORECASE):
                    self._in_references = True
                else:
                    self._in_references = False
                i += 1
                continue

            # --- Regra horizontal ---
            if self._RE_HRULE.match(line):
                self._flush_pending()
                self.blocks.append(HorizontalRuleBlock())
                i += 1
                continue

            # --- Tabela pipe‑style ---
            if self._RE_TABLE_ROW.match(line):
                # coleta linhas da tabela (inclui linha separadora)
                tbl_lines = [line]
                i += 1
                while i < len(lines) and self._RE_TABLE_ROW.match(lines[i]):
                    tbl_lines.append(lines[i])
                    i += 1
                # primeira linha = cabeçalhos, segunda = separador, resto = dados
                header = [c.strip() for c in tbl_lines[0].strip('|').split('|')]
                rows = []
                for row_line in tbl_lines[2:]:
                    rows.append([c.strip() for c in row_line.strip('|').split('|')])
                self._flush_pending()
                self.blocks.append(TableBlock(headers=header, rows=rows, raw="\n".join(tbl_lines)))
                continue

            # --- Equação em bloco ($$ ... $$) ---
            m = self._RE_BLOCK_MATH.match(line)
            if m:
                self._flush_pending()
                expr = m.group(1).strip()
                self.blocks.append(MathBlock(expr=expr, display=True, raw=line))
                i += 1
                continue
            if line.strip().startswith('$$'):
                self._flush_pending()
                i += 1
                math_lines = []
                while i < len(lines) and not lines[i].strip().endswith('$$'):
                    math_lines.append(lines[i])
                    i += 1
                if i < len(lines):
                    closing = lines[i].strip()
                    if closing != '$$':
                        math_lines.append(closing[:-2].rstrip())
                    i += 1
                expr = ' '.join(x.strip() for x in math_lines).strip()
                if expr:
                    self.blocks.append(MathBlock(expr=expr, display=True, raw='\n'.join(math_lines)))
                continue

            # --- Tabela markdown ---
            if (
                i + 1 < len(lines)
                and self._RE_TABLE_ROW.match(line)
                and self._RE_TABLE_SEP.match(lines[i + 1])
            ):
                self._flush_pending()
                table_lines = [line]
                i += 2  # pula cabeçalho e separador
                while i < len(lines) and self._RE_TABLE_ROW.match(lines[i]):
                    table_lines.append(lines[i])
                    i += 1

                headers = self._split_table_row(table_lines[0])
                rows = [self._split_table_row(row) for row in table_lines[1:]]
                if headers:
                    self.blocks.append(TableBlock(headers=headers, rows=rows, raw='\n'.join(table_lines)))
                continue

            # --- Imagem ---
            m = self._RE_IMAGE.match(line)
            if m:
                self._flush_pending()
                self.blocks.append(ImageBlock(alt=m.group(1), url=m.group(2), raw=line))
                i += 1
                continue

            # --- Citação longa (blockquote) ---
            if self._RE_QUOTE.match(line):
                self._flush_pending()
                quote_lines = []
                while i < len(lines) and self._RE_QUOTE.match(lines[i]):
                    quote_lines.append(self._RE_QUOTE.match(lines[i]).group(1))
                    i += 1
                text = ' '.join(quote_lines).strip()
                self.blocks.append(QuoteBlock(text=text, raw='\n'.join(quote_lines)))
                continue

            # --- Lista não-ordenada ---
            m = self._RE_ULIST.match(line)
            if m:
                self._flush_pending()
                self.blocks.append(ListItemBlock(text=m.group(1), ordered=False, raw=line))
                i += 1
                continue

            # --- Lista ordenada ---
            m = self._RE_OLIST.match(line)
            if m:
                self._flush_pending()
                self.blocks.append(ListItemBlock(
                    text=m.group(2), ordered=True, number=int(m.group(1)), raw=line
                ))
                i += 1
                continue

            # --- Linha vazia — flush de parágrafo acumulado ---
            if self._RE_EMPTY.match(line):
                self._flush_pending()
                i += 1
                continue

            # --- Linha de texto — acumular parágrafo ---
            self._pending_lines.append(line)
            i += 1

        self._flush_pending()
        self._assign_table_captions()
        return self.blocks

    # ------------------------------------------------------------------ #
    def _flush_pending(self):
        """Consolida linhas acumuladas num ParagraphBlock ou ReferenceEntryBlock."""
        if not self._pending_lines:
            return

        text = ' '.join(self._pending_lines).strip()
        if text:
            if self._in_references:
                self.blocks.append(ReferenceEntryBlock(text=text, raw=text))
            else:
                self.blocks.append(ParagraphBlock(text=text, raw=text))
        self._pending_lines = []

    # ------------------------------------------------------------------ #
    def _assign_table_captions(self):
        """Dá um título curto às tabelas sem legenda explícita, para a lista de tabelas."""
        for blk in self.blocks:
            if not isinstance(blk, TableBlock) or blk.caption:
                continue
            headers = [re.sub(r'\*+', '', cell).strip() for cell in blk.headers]
            headers = [cell for cell in headers if cell]
            blk.caption = ' / '.join(headers) if headers else 'Tabela sem título no markdown.'

    @staticmethod
    def _split_table_row(line: str) -> list[str]:
        stripped = line.strip().strip('|')
        return [cell.strip() for cell in stripped.split('|')]


# =========================================================================== #
# MATH RENDERER + INLINE FORMATTER                                             #
# =========================================================================== #
class InlineFormatter:
    """
    Recebe uma string com markdown inline e adiciona runs formatados
    a um parágrafo python-docx.

    Suporta:
        **texto**   → negrito
        *texto*     → itálico
        `código`    → monoespaçado
        [DADO FALTANTE] → placeholder em destaque
        $...$       → fórmula inline
    """

    # Tokens: (tipo, regex)
    _PATTERNS = [
        ('bold_italic', re.compile(r'\*\*\*(.+?)\*\*\*')),
        ('bold',        re.compile(r'\*\*(.+?)\*\*')),
        ('italic',      re.compile(r'(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)')),
        ('code',        re.compile(r'`(.+?)`')),
        ('placeholder', re.compile(r'\[DADO FALTANTE[^\]]*\]')),
        ('link',        re.compile(r'\[([^\]]+)\]\([^)]+\)')),
        ('math',        re.compile(r'(?<!\\)\$([^\n$]+?)(?<!\\)\$')),
    ]

    def __init__(self, font_name: str, font_size_pt: int, on_math_failure=None):
        self.font_name = font_name
        self.font_size_pt = font_size_pt
        self.on_math_failure = on_math_failure

    # ------------------------------------------------------------------ #
    def apply(self, paragraph, text: str, font_size_pt: Optional[int] = None):
        """Adiciona runs formatados ao parágrafo python-docx."""
        fs = font_size_pt or self.font_size_pt
        remaining = text

        # Construir lista de segmentos: (texto, tipo)
        segments = self._tokenize(remaining)

        for seg_text, seg_type in segments:
            if seg_type == 'math':
                try:
                    append_equation(paragraph, seg_text, display=False)
                except OmmlError as exc:
                    run = paragraph.add_run(f'${seg_text}$')
                    run.font.name = 'Courier New'
                    run.font.size = Pt(fs)
                    if self.on_math_failure:
                        self.on_math_failure(seg_text, str(exc))
                continue

            run = paragraph.add_run(seg_text)
            run.font.name = self.font_name
            run.font.size = Pt(fs)

            if seg_type == 'bold':
                run.bold = True
            elif seg_type == 'italic':
                run.italic = True
            elif seg_type == 'bold_italic':
                run.bold = True
                run.italic = True
            elif seg_type == 'code':
                run.font.name = 'Courier New'
            elif seg_type == 'placeholder':
                run.bold = True
                run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
            # 'link' e 'text' → sem formatação especial

    # ------------------------------------------------------------------ #
    def _tokenize(self, text: str) -> list[tuple[str, str]]:
        """Quebra o texto em segmentos com seus tipos."""
        if not text:
            return []

        # Encontrar primeira ocorrência de qualquer padrão
        earliest_match = None
        earliest_start = len(text) + 1
        earliest_type = None

        for tipo, pat in self._PATTERNS:
            m = pat.search(text)
            if m and m.start() < earliest_start:
                earliest_match = m
                earliest_start = m.start()
                earliest_type = tipo

        if earliest_match is None:
            return [(text, 'text')]

        result = []
        # Texto antes do match
        if earliest_start > 0:
            result.append((text[:earliest_start], 'text'))

        # O match
        m = earliest_match
        if earliest_type == 'link':
            # Exibir texto do link (grupo 1)
            inner = m.group(1)
        elif earliest_type == 'placeholder':
            inner = m.group(0)  # manter o placeholder completo
        else:
            inner = m.group(1)

        result.append((inner, earliest_type))

        # Recursão no restante
        rest = text[m.end():]
        result.extend(self._tokenize(rest))
        return result


# =========================================================================== #
# ACADEMIC VALIDATOR                                                           #
# =========================================================================== #
class AcademicValidator:
    """
    Valida regras ABNT sobre os blocos e coleta pendências.

    Regras implementadas:
        V1 — Sistema de chamada deve ser exclusivamente autor-data (sem [1], [2])
        V2 — "et al." deve aparecer em itálico no texto gerado
        V3 — Referências sem URL/DOI → pendência
        V4 — Referências sem autor → pendência + placeholder
        V5 — Referências sem ano → pendência + placeholder
        V6 — Imagens com URL externa → pendência (deve ser hospedada no GitHub)
        V7 — Placeholder PENDENTE no texto → pendência
    """

    _RE_NUMERIC_CITATION = re.compile(r'\[\d+\]')
    _RE_ET_AL = re.compile(r'et al\.')
    _RE_AUTHOR_DATE = re.compile(r'\([A-ZÁÉÍÓÚÃÕÂÊÔÇÀÜ][^)]{2,},\s*\d{4}', re.UNICODE)
    _RE_URL = re.compile(r'https?://\S+')
    _RE_DOI = re.compile(r'10\.\d{4,}/\S+')
    _RE_PLACEHOLDER = re.compile(r'PENDENTE|⚠️|TODO|FIXME|\[XXX\]', re.IGNORECASE)
    _RE_NO_AUTHOR = re.compile(r'^\[?(sem autor|anônimo|s\.?n\.?)\]?', re.IGNORECASE)

    def __init__(self, source_md_path: Optional[str] = None, figure_registry_path: Optional[str] = None):
        self.pendencias: list[Pendencia] = []
        self._line_num = 0
        self._source_md_path = Path(source_md_path).resolve() if source_md_path else None
        self._figure_registry_path = Path(figure_registry_path).resolve() if figure_registry_path else None
        self.registry = FigureRegistry(self._figure_registry_path)

    # ------------------------------------------------------------------ #
    def validate_all(self, blocks: list[Block]) -> list[Block]:
        """
        Valida e transforma os blocos:
        - Injeta [DADO FALTANTE] onde necessário
        - Registra pendências
        Retorna os blocos modificados.
        """
        validated = []
        for i, blk in enumerate(blocks):
            self._line_num = i
            blk = self._validate_block(blk)
            validated.append(blk)
        self._validate_registry_files()
        return validated

    # ------------------------------------------------------------------ #
    def _validate_block(self, blk: Block) -> Block:
        """Despacha validação por tipo de bloco."""
        if isinstance(blk, (ParagraphBlock, QuoteBlock)):
            blk = self._validate_text_block(blk)
        elif isinstance(blk, ReferenceEntryBlock):
            blk = self._validate_reference(blk)
        elif isinstance(blk, ImageBlock):
            blk = self._validate_image(blk)
        elif isinstance(blk, TableBlock):
            blk = self._validate_table(blk)
        elif isinstance(blk, MathBlock):
            blk = self._validate_math(blk)
        elif isinstance(blk, HeadingBlock):
            blk = self._validate_heading(blk)
        return blk

    # ------------------------------------------------------------------ #
    def _validate_text_block(self, blk):
        """V1: detectar citações numéricas; V7: detectar placeholders."""
        text = blk.text

        # V1 — citações numéricas proibidas
        if self._RE_NUMERIC_CITATION.search(text):
            self.pendencias.append(Pendencia(
                tipo='CITAÇÃO NUMÉRICA (proibida)',
                descricao='Sistema numérico [1] detectado. ABNT exige sistema autor-data.',
                trecho=text[:120],
                linha=self._line_num,
            ))

        for expr in re.findall(r'(?<!\\)\$([^$\n]*?)(?<!\\)\$', text):
            if not expr.strip():
                self.pendencias.append(Pendencia(
                    tipo='FÓRMULA VAZIA',
                    descricao='Fórmula inline delimitada por $ sem expressão.',
                    trecho=text[:120],
                    linha=self._line_num,
                ))

        # V7 — placeholders explícitos
        if self._RE_PLACEHOLDER.search(text):
            self.pendencias.append(Pendencia(
                tipo='PLACEHOLDER NO TEXTO',
                descricao='Trecho com marcador de conteúdo pendente encontrado no corpo do texto.',
                trecho=text[:120],
                linha=self._line_num,
            ))

        return blk

    # ------------------------------------------------------------------ #
    def _validate_reference(self, blk: ReferenceEntryBlock) -> ReferenceEntryBlock:
        """
        V3 — URL/DOI ausente
        V4 — Autor ausente
        V5 — Ano ausente
        """
        text = blk.text
        modificado = False

        # V4 — Autor
        # Referências ABNT começam com SOBRENOME, Nome ou Entidade
        # Se o texto começa com artigo ou preposição, provavelmente falta autor
        first_word = text.split()[0] if text.split() else ''
        if self._RE_NO_AUTHOR.match(text):
            self.pendencias.append(Pendencia(
                tipo='AUTOR AUSENTE',
                descricao='Entrada de referência sem identificação de autor.',
                trecho=text[:120],
                linha=self._line_num,
            ))
            text = '[DADO FALTANTE: AUTOR] ' + text
            modificado = True

        # V5 — Ano (deve haver um grupo de 4 dígitos > 1900)
        if not re.search(r'\b(19|20)\d{2}\b', text):
            self.pendencias.append(Pendencia(
                tipo='ANO AUSENTE',
                descricao='Referência sem ano de publicação identificável.',
                trecho=text[:120],
                linha=self._line_num,
            ))
            # Adicionar placeholder próximo ao final, antes do "Acesso em" se houver
            text = text.rstrip('.') + ' [DADO FALTANTE: ANO].'
            modificado = True

        # V3 — URL ou DOI
        has_url = bool(self._RE_URL.search(text))
        has_doi = bool(self._RE_DOI.search(text))
        if not has_url and not has_doi:
            self.pendencias.append(Pendencia(
                tipo='URL/DOI AUSENTE',
                descricao='Referência sem URL verificável ou DOI. '
                          'Verificar se existe versão digital acessível (e-book, repositório).',
                trecho=text[:120],
                linha=self._line_num,
            ))
            text = text.rstrip('.') + ' [DADO FALTANTE: URL ou DOI].'
            modificado = True

        # V3b — "Acesso em:" presente mas sem URL
        if re.search(r'Acesso em', text, re.IGNORECASE) and not has_url:
            self.pendencias.append(Pendencia(
                tipo='URL AUSENTE COM "ACESSO EM"',
                descricao='"Acesso em:" presente mas sem URL na referência.',
                trecho=text[:120],
                linha=self._line_num,
            ))

        if modificado:
            blk = ReferenceEntryBlock(text=text, raw=blk.raw)
        return blk

    # ------------------------------------------------------------------ #
    def _validate_image(self, blk: ImageBlock) -> ImageBlock:
        """V6 — imagens com URL externa."""
        if blk.url.startswith('http') and 'github' not in blk.url.lower():
            self.pendencias.append(Pendencia(
                tipo='IMAGEM URL EXTERNA',
                descricao='Imagem hospedada fora do GitHub do projeto. '
                          'Recomendado hospedar em raw.githubusercontent.com para garantir disponibilidade.',
                trecho=f'![{blk.alt}]({blk.url[:80]})',
                linha=self._line_num,
            ))
        if 'PENDENTE' in blk.url.upper():
            self.pendencias.append(Pendencia(
                tipo='IMAGEM PENDENTE',
                descricao='URL de imagem marcada como PENDENTE. Substituir pela URL real antes da entrega.',
                trecho=f'![{blk.alt}]({blk.url[:80]})',
                linha=self._line_num,
            ))
        if self._source_md_path and not blk.url.startswith('http'):
            img_path = (self._source_md_path.parent / blk.url).resolve()
            if not img_path.exists():
                self.pendencias.append(Pendencia(
                    tipo='ASSET DE IMAGEM AUSENTE',
                    descricao='Asset local de imagem não encontrado no caminho referenciado pelo markdown.',
                    trecho=f'![{blk.alt}]({blk.url[:80]})',
                    linha=self._line_num,
                ))
        if self.registry.loaded and not blk.url.startswith('http') and not self.registry.contains(blk.url):
            self.pendencias.append(Pendencia(
                tipo='IMAGEM FORA DO REGISTRO',
                descricao='Imagem referenciada na monografia não foi encontrada no figure_registry.',
                trecho=f'![{blk.alt}]({blk.url[:80]})',
                linha=self._line_num,
            ))
        elif not self.registry.loaded and not blk.url.startswith('http'):
            self.pendencias.append(Pendencia(
                tipo='REGISTRO DE FIGURAS AUSENTE',
                descricao='figure_registry.md não foi encontrado para validar a imagem.',
                trecho=f'![{blk.alt}]({blk.url[:80]})',
                linha=self._line_num,
            ))
        return blk

    # ------------------------------------------------------------------ #
    def _validate_table(self, blk: TableBlock) -> TableBlock:
        if not blk.headers or not blk.rows:
            self.pendencias.append(Pendencia(
                tipo='TABELA MALFORMADA',
                descricao='Tabela markdown sem cabeçalho ou sem linhas de dados.',
                trecho=(blk.raw or '')[:120],
                linha=self._line_num,
            ))
        return blk

    # ------------------------------------------------------------------ #
    def _validate_math(self, blk: MathBlock) -> MathBlock:
        if not blk.expr.strip():
            self.pendencias.append(Pendencia(
                tipo='FÓRMULA VAZIA',
                descricao='Bloco matemático delimitado por $$ sem expressão.',
                trecho=blk.raw[:120],
                linha=self._line_num,
            ))
        return blk

    # ------------------------------------------------------------------ #
    def _validate_heading(self, blk: HeadingBlock) -> HeadingBlock:
        """Verifica se o heading referencia conteúdo ainda não escrito."""
        if self._RE_PLACEHOLDER.search(blk.text):
            self.pendencias.append(Pendencia(
                tipo='HEADING COM PLACEHOLDER',
                descricao='Título de seção contém marcador de conteúdo pendente.',
                trecho=blk.text,
                linha=self._line_num,
            ))
        return blk

    # ------------------------------------------------------------------ #
    def _validate_registry_files(self):
        """Confere se cada asset local indexado no registro existe em disco."""
        if not self.registry.loaded or not self._source_md_path:
            return
        markdown_dir = self._source_md_path.parent
        missing = []
        for arquivo in sorted(self.registry.assets):
            if self.registry.resolve(arquivo, markdown_dir) is None:
                missing.append(arquivo)
        for arquivo in missing:
            self.pendencias.append(Pendencia(
                tipo='ASSET DO REGISTRO AUSENTE',
                descricao='Arquivo citado no figure_registry não foi encontrado no disco.',
                trecho=arquivo,
                linha=None,
            ))


# =========================================================================== #
# DOCX GENERATOR                                                               #
# =========================================================================== #
class DocxGenerator:
    """
    Constrói o .docx a partir da lista de Blocks e do ABNTStyleManager.
    """

    def __init__(self, config: dict, source_md_path: Optional[str] = None):
        self.cfg = config
        self.doc = Document()
        self.style_mgr = ABNTStyleManager(self.doc, config)
        # Renderização de equações agora usa pandoc → OMML (latex_to_omml); MathRenderer removido.
        self.inline = InlineFormatter(config["font_name"], config["font_size_pt"])
        self.source_md_path = Path(source_md_path).resolve() if source_md_path else None
        self.figure_counter = 0
        self.table_counter = 0
        self.runtime_pendencias: list[Pendencia] = []
        self._set_page_margins()

    # ------------------------------------------------------------------ #
    def _set_page_margins(self):
        """Aplica margens ABNT em todas as seções."""
        for section in self.doc.sections:
            section.top_margin    = Cm(self.cfg["margin_top_cm"])
            section.left_margin   = Cm(self.cfg["margin_left_cm"])
            section.bottom_margin = Cm(self.cfg["margin_bottom_cm"])
            section.right_margin  = Cm(self.cfg["margin_right_cm"])

    # ------------------------------------------------------------------ #
    def build(self, blocks: list[Block]) -> Document:
        """Ponto de entrada. Percorre blocks e delega a métodos específicos."""
        for blk in blocks:
            self._render_block(blk)
        return self.doc

    # ------------------------------------------------------------------ #
    def _render_block(self, blk: Block):
        if isinstance(blk, HeadingBlock):
            self._render_heading(blk)
        elif isinstance(blk, ParagraphBlock):
            self._render_paragraph(blk)
        elif isinstance(blk, QuoteBlock):
            self._render_quote(blk)
        elif isinstance(blk, ListItemBlock):
            self._render_list_item(blk)
        elif isinstance(blk, ImageBlock):
            self._render_image(blk)
        elif isinstance(blk, TableBlock):
            self._render_table(blk)
        elif isinstance(blk, MathBlock):
            self._render_math(blk)
        elif isinstance(blk, HorizontalRuleBlock):
            self._render_hrule()
        elif isinstance(blk, ReferenceEntryBlock):
            self._render_reference(blk)

    # ------------------------------------------------------------------ #
    def _render_heading(self, blk: HeadingBlock):
        """Insere heading com caixa-alta condicional e negrito."""
        text = blk.text
        if blk.level in self.cfg["heading_uppercase_levels"]:
            text = text.upper()

        try:
            para = self.doc.add_heading(level=blk.level)
        except Exception:
            para = self.doc.add_paragraph()

        # Limpar runs existentes (add_heading pode adicionar run vazio)
        for run in para.runs:
            run.text = ''

        run = para.add_run(text)
        run.font.name = self.cfg["font_name"]
        run.font.size = Pt(self.cfg["heading_font_size"].get(blk.level, 12))
        run.font.bold = blk.level <= 3
        run.font.italic = (blk.level == 4)
        run.font.color.rgb = RGBColor(0, 0, 0)

        para.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        para.paragraph_format.first_line_indent = Cm(0)

    # ------------------------------------------------------------------ #
    def _render_paragraph(self, blk: ParagraphBlock):
        """Insere parágrafo de corpo com recuo de primeira linha."""
        para = self.doc.add_paragraph()
        para.style = self.doc.styles['Normal']

        # Detectar parágrafo sem recuo (ex: primeiros de seção, listas disfarçadas)
        if not blk.first_line_indent:
            para.paragraph_format.first_line_indent = Cm(0)

        # Processar "et al." em itálico antes de adicionar runs
        text = self._mark_et_al(blk.text)
        self.inline.apply(para, text)

    # ------------------------------------------------------------------ #
    def _render_quote(self, blk: QuoteBlock):
        """Citação longa ABNT: 4cm recuo, 10pt, espaçamento simples."""
        para = self.doc.add_paragraph()
        para.style = self.doc.styles['ABNT Quote']
        text = self._mark_et_al(blk.text)
        self.inline.apply(para, text, font_size_pt=self.cfg["font_size_quote_pt"])

    # ------------------------------------------------------------------ #
    def _render_list_item(self, blk: ListItemBlock):
        """Insere item de lista. Usa List Bullet ou List Number se disponível."""
        style_name = 'List Bullet' if not blk.ordered else 'List Number'
        try:
            para = self.doc.add_paragraph(style=style_name)
        except Exception:
            para = self.doc.add_paragraph()
            prefix = f'{blk.number}. ' if blk.ordered else '• '
            blk = ListItemBlock(text=prefix + blk.text, ordered=blk.ordered)

        text = self._mark_et_al(blk.text)
        self.inline.apply(para, text)

        # Ajustar fonte
        for run in para.runs:
            run.font.name = self.cfg["font_name"]
            run.font.size = Pt(self.cfg["font_size_pt"])

    # ------------------------------------------------------------------ #
    def _render_image(self, blk: ImageBlock):
        image_path = self._resolve_image_path(blk.url)
        rendered_path = None
        if image_path is not None and image_path.exists():
            ext = image_path.suffix.lower()
            if ext in {'.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp'}:
                rendered_path = image_path
            elif ext == '.svg':
                rendered_path = self._convert_svg_to_png(image_path)
            elif ext == '.mermaid':
                rendered_path = self._render_mermaid_to_png(image_path)
        elif blk.url.startswith('http'):
            self.runtime_pendencias.append(Pendencia(
                tipo='IMAGEM EXTERNA NÃO BAIXADA',
                descricao='Imagem remota não foi baixada automaticamente para inserção no DOCX.',
                trecho=blk.url,
                linha=None,
            ))

        if rendered_path and rendered_path.exists():
            para_img = self.doc.add_paragraph()
            para_img.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para_img.add_run()
            run.add_picture(str(rendered_path), width=Cm(15))
        else:
            para_img = self.doc.add_paragraph()
            para_img.style = self.doc.styles['ABNT Image Placeholder']
            run = para_img.add_run(f'[ FIGURA NÃO INSERIDA: {blk.alt or "sem descrição"} ]')
            run.font.name = self.cfg["font_name"]
            run.font.size = Pt(10)
            run.font.italic = True
            run.font.color.rgb = RGBColor(0x80, 0x80, 0x80)
            para_img.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if blk.url:
                para_url = self.doc.add_paragraph()
                para_url.style = self.doc.styles['ABNT Image Placeholder']
                r = para_url.add_run(f'Referência do asset: {blk.url}')
                r.font.size = Pt(8)
                r.font.name = self.cfg["font_name"]
                para_url.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER

        self._render_caption('Figura', blk.alt or 'Sem descrição.')

    # ------------------------------------------------------------------ #
    def _render_table(self, blk: TableBlock):
        self.table_counter += 1
        title = blk.caption or 'Conteúdo tabular extraído do markdown.'
        self._render_caption('Tabela', title)

        col_count = max(len(blk.headers), max((len(r) for r in blk.rows), default=0))
        if col_count == 0:
            return
        table = self.doc.add_table(rows=1, cols=col_count)
        table.style = 'Table Grid'
        table.autofit = False

        # Largura útil da página ABNT (21cm - 3cm esq - 2cm dir = 16cm)
        total_width_cm = 16.0
        total_twips = int(total_width_cm * 567)
        col_twips = total_twips // col_count
        col_width_cm = total_width_cm / col_count

        # Configura largura total no elemento w:tblPr
        tblPr = table._tbl.tblPr
        tblW = tblPr.find(qn('w:tblW'))
        if tblW is None:
            tblW = OxmlElement('w:tblW')
            tblPr.append(tblW)
        tblW.set(qn('w:w'), str(total_twips))
        tblW.set(qn('w:type'), 'dxa')

        hdr_cells = table.rows[0].cells
        for idx in range(col_count):
            val = blk.headers[idx] if idx < len(blk.headers) else ''
            hdr_cells[idx].width = Cm(col_width_cm)
            para = hdr_cells[idx].paragraphs[0]
            para.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            para.paragraph_format.first_line_indent = Cm(0)
            para.paragraph_format.space_before = Pt(2)
            para.paragraph_format.space_after = Pt(2)
            run = para.add_run(val)
            run.font.name = self.cfg["font_name"]
            run.font.size = Pt(9)
            run.bold = True

            # Define largura no w:tcPr
            tcPr = hdr_cells[idx]._tc.get_or_add_tcPr()
            tcW = tcPr.find(qn('w:tcW'))
            if tcW is None:
                tcW = OxmlElement('w:tcW')
                tcPr.append(tcW)
            tcW.set(qn('w:w'), str(col_twips))
            tcW.set(qn('w:type'), 'dxa')

        for row in blk.rows:
            row_cells = table.add_row().cells
            for idx in range(col_count):
                val = row[idx] if idx < len(row) else ''
                row_cells[idx].width = Cm(col_width_cm)
                para = row_cells[idx].paragraphs[0]
                para.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
                para.paragraph_format.first_line_indent = Cm(0)
                para.paragraph_format.space_before = Pt(1)
                para.paragraph_format.space_after = Pt(1)
                run = para.add_run(val)
                run.font.name = self.cfg["font_name"]
                run.font.size = Pt(9)

                tcPr = row_cells[idx]._tc.get_or_add_tcPr()
                tcW = tcPr.find(qn('w:tcW'))
                if tcW is None:
                    tcW = OxmlElement('w:tcW')
                    tcPr.append(tcW)
                tcW.set(qn('w:w'), str(col_twips))
                tcW.set(qn('w:type'), 'dxa')

    # ------------------------------------------------------------------ #
    def _render_math(self, blk: MathBlock):
        """Renderiza equação LaTeX → OMML usando latex2mathml (omml_math.py).

        Usa ``append_equation(parágrafo, latex_str)`` que converte via
        latex2mathml, sem depender do pandoc. Se falhar, insere texto
        monoespaçado e registra pendência.
        """
        from omml_math import append_equation, OmmlError

        latex = blk.expr.strip()

        # Cria o parágrafo centrado sem recuo
        para = self.doc.add_paragraph()
        para.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        para.paragraph_format.first_line_indent = Cm(0)
        para.paragraph_format.space_before = Pt(6)
        para.paragraph_format.space_after = Pt(6)

        try:
            append_equation(para, latex, display=blk.display)
            return
        except OmmlError:
            pass  # cai no fallback abaixo
        except Exception:
            pass

        # Fallback: texto monoespaçado com a expressão LaTeX
        self.runtime_pendencias.append(Pendencia(
            tipo='EQUAÇÃO SEM CONVERSÃO',
            descricao='Falha ao converter LaTeX → OMML (latex2mathml); expressão mantida como texto.',
            trecho=blk.expr[:120],
            linha=None,
        ))
        run = para.add_run(f"$${latex}$$" if blk.display else f"${latex}$")
        run.font.name = 'Courier New'
        run.font.size = Pt(10)

    # ------------------------------------------------------------------ #
    def _render_hrule(self):
        """Linha horizontal → quebra temática (parágrafo vazio com borda inferior)."""
        para = self.doc.add_paragraph()
        para.paragraph_format.space_before = Pt(6)
        para.paragraph_format.space_after = Pt(6)
        # Borda inferior via XML
        pPr = para._p.get_or_add_pPr()
        pBdr = OxmlElement('w:pBdr')
        bottom = OxmlElement('w:bottom')
        bottom.set(qn('w:val'), 'single')
        bottom.set(qn('w:sz'), '6')
        bottom.set(qn('w:space'), '1')
        bottom.set(qn('w:color'), 'auto')
        pBdr.append(bottom)
        pPr.append(pBdr)

    # ------------------------------------------------------------------ #
    def _render_reference(self, blk: ReferenceEntryBlock):
        """Entrada de referência bibliográfica: 10pt, espaçamento simples, sem recuo."""
        para = self.doc.add_paragraph()
        para.style = self.doc.styles['ABNT Reference']

        # Processar et al. e inline markup
        text = self._mark_et_al(blk.text)
        self.inline.apply(para, text, font_size_pt=self.cfg["font_size_ref_pt"])

    # ------------------------------------------------------------------ #
    @staticmethod
    def _mark_et_al(text: str) -> str:
        """
        Garante que "et al." seja renderizado em itálico.
        O InlineFormatter vai tratar *et al.* como itálico.
        """
        # Substituir "et al." por "*et al.*" se ainda não estiver em itálico
        return re.sub(
            r'(?<!\*)\bet al\.\b(?!\*)',
            '*et al.*',
            text
        )

    # ------------------------------------------------------------------ #
    def _render_caption(self, label: str, text: str):
        if label == 'Figura':
            self.figure_counter += 1
            number = self.figure_counter
        else:
            number = self.table_counter

        para = self.doc.add_paragraph()
        para.style = self.doc.styles['ABNT Caption']
        run = para.add_run(f'{label} {number} — {text}')
        run.font.name = self.cfg["font_name"]
        run.font.size = Pt(10)
        para.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # ------------------------------------------------------------------ #
    def _resolve_image_path(self, url: str) -> Optional[Path]:
        if not url or url.startswith('http'):
            return None
        candidate = Path(url)
        if candidate.is_absolute():
            return candidate
        if self.source_md_path:
            return (self.source_md_path.parent / candidate).resolve()
        return candidate.resolve()

    # ------------------------------------------------------------------ #
    def _convert_svg_to_png(self, svg_path: Path) -> Optional[Path]:
        """Converte SVG para PNG usando Inkscape CLI se disponível, ou cairosvg como fallback."""
        # Se já existe um PNG com sufixo manual (_png.png ou .rendered.png)
        manual_png = svg_path.with_name(svg_path.stem + '_png.png')
        if manual_png.exists():
            return manual_png

        out = svg_path.with_suffix('.svg.rendered.png')
        if out.exists():
            return out

        inkscape = shutil.which('inkscape')
        if inkscape:
            cmd = [inkscape, '--export-type=png', f'--export-filename={out}', str(svg_path)]
            try:
                subprocess.run(cmd, check=True, capture_output=True, timeout=30)
                if out.exists():
                    return out
            except Exception as exc:
                self.runtime_pendencias.append(Pendencia(
                    tipo='SVG SEM CONVERSÃO',
                    descricao=f'Falha no Inkscape ao converter SVG: {exc}',
                    trecho=str(svg_path),
                    linha=None,
                ))

        try:
            import cairosvg
            cairosvg.svg2png(url=str(svg_path), write_to=str(out))
            if out.exists():
                return out
        except Exception:
            pass

        if not out.exists():
            self.runtime_pendencias.append(Pendencia(
                tipo='SVG SEM CONVERSÃO',
                descricao='Não foi possível converter SVG (Inkscape/cairosvg falhou ou ausente).',
                trecho=str(svg_path),
                linha=None,
            ))
            return None
        return out

    # ------------------------------------------------------------------ #
    def _render_mermaid_to_png(self, mermaid_path: Path) -> Optional[Path]:
        # Se o usuário já gerou ou tirou print manualmente (_png.png)
        manual_png = mermaid_path.with_name(mermaid_path.stem + '_png.png')
        if manual_png.exists():
            return manual_png

        out = mermaid_path.with_suffix('.mermaid.rendered.png')
        if out.exists():
            return out

        mmdc = shutil.which('mmdc')
        if not mmdc:
            self.runtime_pendencias.append(Pendencia(
                tipo='MERMAID SEM RENDERIZAÇÃO',
                descricao='Comando mmdc não encontrado para renderizar arquivo .mermaid.',
                trecho=str(mermaid_path),
                linha=None,
            ))
            return None

        cmd = [mmdc, '-i', str(mermaid_path), '-o', str(out), '-b', 'white']
        try:
            subprocess.run(cmd, check=True, capture_output=True, text=True, timeout=60)
            if out.exists():
                return out
        except subprocess.CalledProcessError as exc:
            err = (exc.stderr.strip() or exc.stdout.strip() or str(exc))[:200]
            self.runtime_pendencias.append(Pendencia(
                tipo='MERMAID SEM RENDERIZAÇÃO',
                descricao=f'Falha no mmdc (Chromium/Puppeteer ausente ou indisponível): {err}',
                trecho=str(mermaid_path),
                linha=None,
            ))
            return None
        except Exception as exc:
            self.runtime_pendencias.append(Pendencia(
                tipo='MERMAID SEM RENDERIZAÇÃO',
                descricao=f'Erro ao executar mmdc: {exc}',
                trecho=str(mermaid_path),
                linha=None,
            ))
            return None
        return out if out.exists() else None

    # ------------------------------------------------------------------ #
    def save(self, output_path: str):
        self.doc.save(output_path)


# =========================================================================== #
# REPORT GENERATOR                                                             #
# =========================================================================== #
class ReportGenerator:
    """
    Gera relatorio_pendencias.md e imprime resumo no terminal.
    """

    def __init__(self, pendencias: list[Pendencia], source_file: str, output_file: str):
        self.pendencias = pendencias
        self.source = source_file
        self.output = output_file
        self.ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    # ------------------------------------------------------------------ #
    def write_report(self, report_path: str = 'relatorio_pendencias.md'):
        lines = [
            '# Relatório de Pendências — Monografia ABNT',
            f'',
            f'**Gerado em:** {self.ts}  ',
            f'**Fonte:** `{self.source}`  ',
            f'**Saída:** `{self.output}`  ',
            f'**Total de pendências:** {len(self.pendencias)}',
            '',
            '---',
            '',
        ]

        if not self.pendencias:
            lines.append('✅ Nenhuma pendência encontrada.')
        else:
            # Agrupar por tipo
            by_type: dict[str, list[Pendencia]] = {}
            for p in self.pendencias:
                by_type.setdefault(p.tipo, []).append(p)

            for tipo, items in sorted(by_type.items()):
                lines.append(f'## {tipo} ({len(items)} ocorrência{"s" if len(items) > 1 else ""})')
                lines.append('')
                for idx, p in enumerate(items, 1):
                    lines.append(f'### {idx}. {p.descricao}')
                    if p.linha is not None:
                        lines.append(f'- **Bloco #{p.linha}**')
                    lines.append(f'- **Trecho:** `{p.trecho}`')
                    lines.append('')

        lines.append('---')
        lines.append('*Relatório gerado automaticamente por monografia_converter.py*')

        Path(report_path).write_text('\n'.join(lines), encoding='utf-8')
        print(f'📋 Relatório salvo em: {report_path}')

    # ------------------------------------------------------------------ #
    def print_summary(self, blocks: list[Block]):
        """Imprime resumo no terminal."""
        counts = {}
        for blk in blocks:
            t = type(blk).__name__
            counts[t] = counts.get(t, 0) + 1

        print('\n' + '='*60)
        print('  MONOGRAFIA CONVERTER — ABNT NBR 14724')
        print('='*60)
        print(f'  Fonte:  {self.source}')
        print(f'  Saída:  {self.output}')
        print(f'  Data:   {self.ts}')
        print('-'*60)
        print('  BLOCOS PROCESSADOS:')
        for blk_type, count in sorted(counts.items()):
            label = {
                'HeadingBlock': 'Títulos',
                'ParagraphBlock': 'Parágrafos',
                'QuoteBlock': 'Citações longas',
                'ListItemBlock': 'Itens de lista',
                'ImageBlock': 'Imagens',
                'TableBlock': 'Tabelas markdown',
                'MathBlock': 'Fórmulas matemáticas',
                'HorizontalRuleBlock': 'Separadores',
                'ReferenceEntryBlock': 'Referências bibliográficas',
            }.get(blk_type, blk_type)
            print(f'    {label:<30} {count:>4}')
        print('-'*60)

        if self.pendencias:
            tipos_count: dict[str, int] = {}
            for p in self.pendencias:
                tipos_count[p.tipo] = tipos_count.get(p.tipo, 0) + 1

            print(f'  ⚠️  PENDÊNCIAS ABNT: {len(self.pendencias)}')
            for tipo, cnt in sorted(tipos_count.items()):
                print(f'    • {tipo}: {cnt}')
        else:
            print('  ✅  Nenhuma pendência ABNT encontrada.')

        print('='*60 + '\n')


# =========================================================================== #
# MAIN ORCHESTRATOR                                                            #
# =========================================================================== #
def main():
    parser = argparse.ArgumentParser(
        description='Converte monografia.md → DOCX com conformidade ABNT.'
    )
    parser.add_argument('input', help='Arquivo .md de entrada')
    parser.add_argument('--output', '-o', default=None,
                        help='Arquivo .docx de saída (padrão: mesmo nome do input)')
    parser.add_argument('--report', '-r', default='relatorio_pendencias.md',
                        help='Arquivo de relatório de pendências')
    parser.add_argument('--font', choices=['Arial', 'Times New Roman'],
                        default='Arial', help='Fonte principal (padrão: Arial)')
    args = parser.parse_args()

    # Validar entrada
    input_path = Path(args.input)
    if not input_path.exists():
        print(f'❌ Arquivo não encontrado: {input_path}')
        sys.exit(1)

    output_path = args.output or str(input_path.with_suffix('.docx'))

    # Aplicar configuração de fonte
    cfg = deepcopy(CONFIG)
    cfg["font_name"] = args.font

    print(f'🔄 Iniciando conversão: {input_path} → {output_path}')

    # 1. Parse
    print('  [1/4] Parseando markdown...')
    md_parser = MarkdownParser(str(input_path))
    blocks = md_parser.parse()
    print(f'        {len(blocks)} blocos extraídos.')

    # 2. Validação ABNT
    print('  [2/4] Validando regras ABNT...')
    registry_path = input_path.parent / 'figure_registry.md'
    validator = AcademicValidator(
        source_md_path=str(input_path),
        figure_registry_path=str(registry_path) if registry_path.exists() else None,
    )
    blocks = validator.validate_all(blocks)
    print(f'        {len(validator.pendencias)} pendências encontradas.')

    # 3. Gerar DOCX
    print('  [3/4] Gerando DOCX...')
    generator = DocxGenerator(cfg, source_md_path=str(input_path))
    generator.build(blocks)
    generator.save(output_path)
    print(f'        ✅ Salvo: {output_path}')

    # 4. Relatório
    print('  [4/4] Gerando relatório...')
    pendencias = validator.pendencias + generator.runtime_pendencias
    reporter = ReportGenerator(
        pendencias=pendencias,
        source_file=str(input_path),
        output_file=output_path,
    )
    reporter.write_report(args.report)
    reporter.print_summary(blocks)


if __name__ == '__main__':
    main()
