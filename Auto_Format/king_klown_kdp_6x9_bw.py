#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
king_klown_kdp_6x9_bw.py

Convertit / standardise un .docx exporté de Google Docs ou Word vers un
format intérieur Amazon KDP 6 x 9 pouces, noir et blanc, typographie réduite.

Entrée :
    .docx

Sorties :
    .docx formaté KDP 6x9 noir et blanc
    .pdf optionnel via LibreOffice si installé

Dépendance :
    pip install python-docx

Lancer le GUI :
    python king_klown_kdp_6x9_bw.py

CLI :
    python king_klown_kdp_6x9_bw.py "manuel.docx" --pdf
    python king_klown_kdp_6x9_bw.py "./docs" --recursive --page-count 205 --pdf

Notes KDP :
- Format intérieur par défaut : 6 x 9 pouces, sans bleed.
- Marges miroir pour livre relié.
- Noir et blanc : aucune couleur vive, tableaux et encadrés en gris.
- Pour un intérieur avec bleed, le script offre --bleed, mais pour un manuel
  texte/tableaux, --no-bleed reste recommandé.
"""

from __future__ import annotations

import argparse
import queue
import re
import shutil
import subprocess
import sys
import threading
import time
import tkinter as tk
from dataclasses import dataclass, field
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Callable, Iterable, Optional
from zipfile import ZIP_DEFLATED, ZipFile

try:
    from docx import Document
    from docx.document import Document as DocxDocument
    from docx.enum.section import WD_SECTION_START
    from docx.enum.style import WD_STYLE_TYPE
    from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Inches, Pt, RGBColor
except ImportError as exc:
    raise SystemExit(
        "Dépendance manquante : python-docx.\n"
        "Installe-la avec : pip install python-docx"
    ) from exc


# ---------------------------------------------------------------------------
# Charte KDP 6x9 noir et blanc
# ---------------------------------------------------------------------------

BLACK = "000000"
DARK = "1A1A1A"
GRAY = "666666"
MID_GRAY = "A9A9A9"
LIGHT_GRAY = "EFEFEF"
VERY_LIGHT_GRAY = "F7F7F7"
BORDER = "CFCFCF"
WHITE = "FFFFFF"


@dataclass
class KDPConfig:
    # KDP 6x9. Sans bleed : 6.0 x 9.0. Avec bleed : 6.125 x 9.25.
    bleed: bool = False
    page_width_in: float = 6.0
    page_height_in: float = 9.0

    # Marges miroir. Les valeurs sont volontairement un peu au-dessus des
    # minimums KDP pour éviter les rejets et améliorer la lecture.
    page_count: int = 205
    top_margin_in: float = 0.55
    bottom_margin_in: float = 0.55
    outside_margin_in: float = 0.50
    inside_margin_in: Optional[float] = None

    header_distance_in: float = 0.28
    footer_distance_in: float = 0.32

    body_font: str = "EB Garamond"
    heading_font: str = "EB Garamond"
    mono_font: str = "Consolas"

    # Typo plus petite pour 6x9.
    body_size_pt: float = 9.6
    line_spacing: float = 1.0
    body_alignment: str = "left"  # left, smart-justify, justify

    add_toc: bool = False
    add_header_footer: bool = True
    header_layout: str = "chapter"  # chapter, brand-chapter, minimal, none
    footer_text: str = "Univers-Cité King Klown"

    chapter_page_breaks: bool = True
    clear_google_run_formatting: bool = True
    normalize_heading_levels: bool = True
    standardize_tables: bool = True
    add_callout_boxes: bool = True
    export_pdf: bool = False


@dataclass
class ProcessStats:
    source: Path
    output: Path
    headings: int = 0
    callouts: int = 0
    formula_blocks: int = 0
    code_lines: int = 0
    tables: int = 0
    blank_paragraphs_removed: int = 0
    warnings: list[str] = field(default_factory=list)
    dry_run: bool = False


# ---------------------------------------------------------------------------
# OOXML helpers
# ---------------------------------------------------------------------------

def _hex(rgb: str) -> str:
    return rgb.replace("#", "").upper()


def get_or_add(parent, child_tag: str):
    child = parent.find(qn(child_tag))
    if child is None:
        child = OxmlElement(child_tag)
        parent.append(child)
    return child


def enable_mirror_margins(doc: DocxDocument) -> None:
    settings = doc.settings._element
    if settings.find(qn("w:mirrorMargins")) is None:
        settings.append(OxmlElement("w:mirrorMargins"))


def set_cell_shading(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = get_or_add(tc_pr, "w:shd")
    shd.set(qn("w:fill"), _hex(fill))


def set_cell_borders(cell, color: str = BORDER, size: str = "6") -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = get_or_add(tc_pr, "w:tcBorders")

    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        element = borders.find(qn(f"w:{edge}"))
        if element is None:
            element = OxmlElement(f"w:{edge}")
            borders.append(element)
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), size)
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), _hex(color))


def set_cell_margins(cell, top: int = 70, start: int = 80, bottom: int = 70, end: int = 80) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = get_or_add(tc_pr, "w:tcMar")
    values = {"top": top, "start": start, "bottom": bottom, "end": end}

    for side, value in values.items():
        element = tc_mar.find(qn(f"w:{side}"))
        if element is None:
            element = OxmlElement(f"w:{side}")
            tc_mar.append(element)
        element.set(qn("w:w"), str(value))
        element.set(qn("w:type"), "dxa")


def set_paragraph_shading(paragraph, fill: str) -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    shd = get_or_add(p_pr, "w:shd")
    shd.set(qn("w:fill"), _hex(fill))


def set_paragraph_border(paragraph, side: str = "left", color: str = MID_GRAY, size: str = "10") -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    borders = get_or_add(p_pr, "w:pBdr")

    border = borders.find(qn(f"w:{side}"))
    if border is None:
        border = OxmlElement(f"w:{side}")
        borders.append(border)

    border.set(qn("w:val"), "single")
    border.set(qn("w:sz"), size)
    border.set(qn("w:space"), "5")
    border.set(qn("w:color"), _hex(color))


def remove_paragraph(paragraph) -> None:
    element = paragraph._element
    parent = element.getparent()
    if parent is not None:
        parent.remove(element)


def add_field(paragraph, instr: str, placeholder: str = "") -> None:
    run = paragraph.add_run()

    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    run._r.append(begin)

    instr_text = OxmlElement("w:instrText")
    instr_text.set(qn("xml:space"), "preserve")
    instr_text.text = instr
    run._r.append(instr_text)

    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    run._r.append(separate)

    if placeholder:
        t = OxmlElement("w:t")
        t.text = placeholder
        run._r.append(t)

    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.append(end)


def enable_update_fields_on_open(docx_path: Path) -> None:
    tmp_path = docx_path.with_suffix(".tmp.docx")
    with ZipFile(docx_path, "r") as zin, ZipFile(tmp_path, "w", ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/settings.xml":
                xml = data.decode("utf-8")
                if "w:updateFields" not in xml:
                    xml = xml.replace("</w:settings>", '<w:updateFields w:val="true"/></w:settings>')
                    data = xml.encode("utf-8")
            zout.writestr(item, data)
    tmp_path.replace(docx_path)


def clear_run_formatting_keep_emphasis(run) -> None:
    r_pr = run._r.rPr
    if r_pr is None:
        return

    for tag in [
        "w:rFonts", "w:sz", "w:szCs", "w:color", "w:highlight", "w:shd",
        "w:spacing", "w:kern", "w:position", "w:lang",
    ]:
        for child in list(r_pr.findall(qn(tag))):
            r_pr.remove(child)


# ---------------------------------------------------------------------------
# KDP margins
# ---------------------------------------------------------------------------

def recommended_inside_margin_in(page_count: int) -> float:
    """KDP gutter guideline plus small safety.

    Official minimums:
    24-150: 0.375"
    151-300: 0.5"
    301-500: 0.625"
    501-700: 0.75"
    701-828: 0.875"

    The script adds a small safety margin for long manuals.
    """
    if page_count <= 150:
        return 0.45
    if page_count <= 300:
        return 0.56
    if page_count <= 500:
        return 0.68
    if page_count <= 700:
        return 0.81
    return 0.94


def apply_bleed_size(cfg: KDPConfig) -> tuple[float, float]:
    if cfg.bleed:
        # KDP interior bleed: +0.125 width, +0.25 height for 6x9.
        return 6.125, 9.25
    return cfg.page_width_in, cfg.page_height_in


# ---------------------------------------------------------------------------
# Styles
# ---------------------------------------------------------------------------

def style_lookup(doc: DocxDocument, *names: str):
    lower = {s.name.lower(): s for s in doc.styles if getattr(s, "name", None)}
    for name in names:
        try:
            return doc.styles[name]
        except KeyError:
            pass
        found = lower.get(name.lower())
        if found is not None:
            return found
    return None


def ensure_paragraph_style(doc: DocxDocument, name: str, base: Optional[str] = None):
    style = style_lookup(doc, name)
    if style is not None:
        return style

    style = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
    if base:
        base_style = style_lookup(doc, base)
        if base_style is not None:
            style.base_style = base_style
    return style


def set_font(style, name: str, size_pt: Optional[float] = None,
             bold: Optional[bool] = None, italic: Optional[bool] = None,
             color: Optional[str] = None) -> None:
    style.font.name = name
    if style._element.rPr is not None and style._element.rPr.rFonts is not None:
        style._element.rPr.rFonts.set(qn("w:eastAsia"), name)
        style._element.rPr.rFonts.set(qn("w:cs"), name)
    if size_pt is not None:
        style.font.size = Pt(size_pt)
    if bold is not None:
        style.font.bold = bold
    if italic is not None:
        style.font.italic = italic
    if color is not None:
        style.font.color.rgb = RGBColor.from_string(_hex(color))


def alignment_from_body_setting(body_alignment: str):
    if body_alignment in {"justify", "smart-justify"}:
        return WD_ALIGN_PARAGRAPH.JUSTIFY
    return WD_ALIGN_PARAGRAPH.LEFT


def configure_styles(doc: DocxDocument, cfg: KDPConfig) -> None:
    normal = style_lookup(doc, "Normal", "normal") or ensure_paragraph_style(doc, "Normal")
    set_font(normal, cfg.body_font, cfg.body_size_pt, color=DARK)
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(3.4)
    normal.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    normal.paragraph_format.line_spacing = cfg.line_spacing
    normal.paragraph_format.alignment = alignment_from_body_setting(cfg.body_alignment)

    h1 = style_lookup(doc, "Heading 1", "Titre 1") or ensure_paragraph_style(doc, "Heading 1", "Normal")
    set_font(h1, cfg.heading_font, 13.8, bold=True, color=BLACK)
    h1.paragraph_format.space_before = Pt(20)
    h1.paragraph_format.space_after = Pt(7)
    h1.paragraph_format.keep_with_next = True
    h1.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT

    h2 = style_lookup(doc, "Heading 2", "Titre 2") or ensure_paragraph_style(doc, "Heading 2", "Normal")
    set_font(h2, cfg.heading_font, 11.2, bold=True, color=BLACK)
    h2.paragraph_format.space_before = Pt(12)
    h2.paragraph_format.space_after = Pt(3)
    h2.paragraph_format.keep_with_next = True
    h2.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT

    h3 = style_lookup(doc, "Heading 3", "Titre 3") or ensure_paragraph_style(doc, "Heading 3", "Normal")
    set_font(h3, cfg.heading_font, 10.1, bold=True, color=BLACK)
    h3.paragraph_format.space_before = Pt(9)
    h3.paragraph_format.space_after = Pt(2.5)
    h3.paragraph_format.keep_with_next = True
    h3.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT

    h4 = style_lookup(doc, "Heading 4", "Titre 4") or ensure_paragraph_style(doc, "Heading 4", "Normal")
    set_font(h4, cfg.heading_font, 9.4, bold=True, color=GRAY)
    h4.paragraph_format.space_before = Pt(7)
    h4.paragraph_format.space_after = Pt(2)
    h4.paragraph_format.keep_with_next = True
    h4.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT

    callout = ensure_paragraph_style(doc, "KDP - Encadre", "Normal")
    set_font(callout, cfg.body_font, cfg.body_size_pt, color=DARK)
    callout.paragraph_format.left_indent = Cm(0.28)
    callout.paragraph_format.right_indent = Cm(0.18)
    callout.paragraph_format.space_before = Pt(5)
    callout.paragraph_format.space_after = Pt(5)
    callout.paragraph_format.line_spacing = 1.0
    callout.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT

    formula = ensure_paragraph_style(doc, "KDP - Formule", "Normal")
    set_font(formula, cfg.body_font, cfg.body_size_pt, color=DARK)
    formula.paragraph_format.left_indent = Cm(0.28)
    formula.paragraph_format.space_before = Pt(3)
    formula.paragraph_format.space_after = Pt(3)
    formula.paragraph_format.line_spacing = 1.0
    formula.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT

    transformation = ensure_paragraph_style(doc, "KDP - Transformation", "Normal")
    set_font(transformation, cfg.body_font, cfg.body_size_pt, italic=True, color=GRAY)
    transformation.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    transformation.paragraph_format.space_before = Pt(4)
    transformation.paragraph_format.space_after = Pt(4)
    transformation.paragraph_format.line_spacing = 1.0

    code = ensure_paragraph_style(doc, "KDP - Code", "Normal")
    set_font(code, cfg.mono_font, 8.2, color=DARK)
    code.paragraph_format.left_indent = Cm(0.25)
    code.paragraph_format.space_before = Pt(3)
    code.paragraph_format.space_after = Pt(3)
    code.paragraph_format.line_spacing = 1.0
    code.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT


# ---------------------------------------------------------------------------
# Détection de structure
# ---------------------------------------------------------------------------

H4_RE = re.compile(r"^\s*\d{1,2}\.\d+\.\d+\.\d+\s+.+")
H3_RE = re.compile(r"^\s*\d{1,2}\.\d+\.\d+\s+.+")
H2_RE = re.compile(r"^\s*\d{1,2}\.\d+\s+.+")
H1_RE = re.compile(r"^\s*\d{2}\.\s+.+")
PART_RE = re.compile(r"^\s*Partie\s+\d+\s+[—-]\s+.+", re.IGNORECASE)
CHAPTER_RE = re.compile(r"^\s*(Chapitre|Chapter)\s+\d+\s+[—-]\s+.+", re.IGNORECASE)

CALLOUT_RE = re.compile(
    r"^\s*(Phrase clé|Point de vigilance|Formule courte|Formule officielle|"
    r"Devise|Règle|À retenir|Retenir|Erreur fréquente|Erreurs fréquentes|"
    r"Production attendue|Critères de réussite|Critères|Trace à conserver|"
    r"Livrable|Mini-exercice|Exercice|Objectifs du chapitre|Attention|Vigilance)\b",
    re.IGNORECASE,
)

GENERIC_TITLES = {
    "word document", "document", "untitled document", "sans titre", "document sans titre",
}


def clean_markdown_artifacts(text: str) -> str:
    text = text.strip()
    text = re.sub(r"^#{1,6}\s*", "", text)
    text = re.sub(r"^\*\*(.+)\*\*$", r"\1", text)
    text = re.sub(r"^__(.+)__$", r"\1", text)
    text = text.replace("\\#", "#").replace("\\_", "_").replace("\\*", "*")
    return text.strip()


def replace_paragraph_text(paragraph, new_text: str) -> None:
    if paragraph.runs:
        for i, run in enumerate(paragraph.runs):
            run.text = new_text if i == 0 else ""
    else:
        paragraph.add_run(new_text)


def strip_markdown_heading_marks(paragraph) -> None:
    replace_paragraph_text(paragraph, clean_markdown_artifacts(paragraph.text))


def classify_heading(text: str) -> Optional[str]:
    stripped = text.strip()
    if not stripped:
        return None

    if stripped.startswith("#### "):
        return "Heading 4"
    if stripped.startswith("### "):
        return "Heading 3"
    if stripped.startswith("## "):
        return "Heading 2"
    if stripped.startswith("# "):
        return "Heading 1"

    cleaned = clean_markdown_artifacts(stripped)

    if H4_RE.match(cleaned):
        return "Heading 4"
    if H3_RE.match(cleaned):
        return "Heading 3"
    if H2_RE.match(cleaned):
        return "Heading 2"
    if H1_RE.match(cleaned) or PART_RE.match(cleaned) or CHAPTER_RE.match(cleaned):
        return "Heading 1"
    return None


def is_list_paragraph(paragraph) -> bool:
    p_pr = paragraph._p.pPr
    return p_pr is not None and p_pr.numPr is not None


def has_manual_line_breaks(paragraph) -> bool:
    if "\n" in paragraph.text:
        return True
    return any("\n" in run.text for run in paragraph.runs)


def is_spread_risk_text(text: str) -> bool:
    cleaned = clean_markdown_artifacts(text)
    if not cleaned:
        return False
    lines = [ln.strip() for ln in cleaned.splitlines() if ln.strip()]
    if len(lines) > 1:
        return any(len(ln) <= 120 for ln in lines)
    if len(cleaned) <= 140:
        return True
    words = cleaned.split()
    return len(words) <= 12


def looks_like_formula_block(text: str) -> bool:
    cleaned = clean_markdown_artifacts(text)
    lines = [ln.strip() for ln in cleaned.splitlines() if ln.strip()]

    if not lines:
        return False
    if "→" in cleaned and len(cleaned) <= 220:
        return True
    if len(lines) >= 2 and all(len(ln) <= 90 for ln in lines):
        hits = sum(1 for ln in lines if re.search(r"\bavant\b|\b→\b|:$", ln, flags=re.IGNORECASE))
        return hits >= 1
    if len(cleaned) <= 110 and re.search(r"\bavant\b", cleaned, flags=re.IGNORECASE):
        return True
    return False


def looks_like_code_line(text: str) -> bool:
    cleaned = clean_markdown_artifacts(text)
    return bool(re.match(r"^\s*(```|[a-zA-Z0-9_]+\s*=\s*|[A-Za-z0-9_./-]+\.(py|md|docx|pdf))", cleaned))


def standardize_paragraphs(doc: DocxDocument, cfg: KDPConfig, stats: ProcessStats) -> None:
    first_h1_seen = False
    normal = style_lookup(doc, "Normal", "normal")
    formula_style = style_lookup(doc, "KDP - Formule")
    transformation_style = style_lookup(doc, "KDP - Transformation")
    code_style = style_lookup(doc, "KDP - Code")
    callout_style = style_lookup(doc, "KDP - Encadre")

    for p in doc.paragraphs:
        text = p.text.strip()

        if cfg.clear_google_run_formatting:
            for run in p.runs:
                clear_run_formatting_keep_emphasis(run)

        if cfg.normalize_heading_levels:
            level = classify_heading(text)
            if level:
                strip_markdown_heading_marks(p)
                target = style_lookup(doc, level, f"Titre {level[-1]}")
                if target is not None:
                    p.style = target
                p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT

                if level == "Heading 1":
                    if cfg.chapter_page_breaks and first_h1_seen:
                        p.paragraph_format.page_break_before = True
                    first_h1_seen = True

                stats.headings += 1
                continue

        if is_list_paragraph(p):
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.line_spacing = 1.0
            p.paragraph_format.space_after = Pt(1.6)
            continue

        if text and CALLOUT_RE.match(clean_markdown_artifacts(text)):
            if callout_style is not None:
                p.style = callout_style
            if cfg.add_callout_boxes:
                set_paragraph_shading(p, VERY_LIGHT_GRAY)
                set_paragraph_border(p, "left", MID_GRAY, "10")
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
            stats.callouts += 1
            continue

        if text and looks_like_formula_block(text):
            if "→" in text and transformation_style is not None:
                p.style = transformation_style
                p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif formula_style is not None:
                p.style = formula_style
                p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
            stats.formula_blocks += 1
            continue

        if text and looks_like_code_line(text):
            if code_style is not None:
                p.style = code_style
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
            stats.code_lines += 1
            continue

        if normal is not None and p.style is not None:
            if p.style.name.lower() in {"normal", "normal text", "body text"}:
                p.style = normal

        # Anti-spread : évite les grands blancs dans les exports justifiés.
        if text:
            if cfg.body_alignment == "left":
                p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
            elif cfg.body_alignment == "smart-justify":
                if has_manual_line_breaks(p) or is_spread_risk_text(text):
                    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
                else:
                    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            elif cfg.body_alignment == "justify":
                if has_manual_line_breaks(p) and is_spread_risk_text(text):
                    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
                else:
                    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY


def remove_excess_blank_paragraphs(doc: DocxDocument, stats: ProcessStats, max_blanks: int = 1) -> None:
    blanks = 0
    for p in list(doc.paragraphs):
        if p.text.strip():
            blanks = 0
            continue
        blanks += 1
        if blanks > max_blanks:
            remove_paragraph(p)
            stats.blank_paragraphs_removed += 1


# ---------------------------------------------------------------------------
# Page, header, footer, tables, TOC
# ---------------------------------------------------------------------------

def configure_sections(doc: DocxDocument, cfg: KDPConfig) -> None:
    enable_mirror_margins(doc)

    width, height = apply_bleed_size(cfg)
    inside = cfg.inside_margin_in if cfg.inside_margin_in is not None else recommended_inside_margin_in(cfg.page_count)

    for section in doc.sections:
        section.page_width = Inches(width)
        section.page_height = Inches(height)

        section.top_margin = Inches(cfg.top_margin_in)
        section.bottom_margin = Inches(cfg.bottom_margin_in)
        section.left_margin = Inches(inside)
        section.right_margin = Inches(cfg.outside_margin_in)
        section.header_distance = Inches(cfg.header_distance_in)
        section.footer_distance = Inches(cfg.footer_distance_in)
        section.start_type = WD_SECTION_START.NEW_PAGE


def get_document_title(doc: DocxDocument, fallback: str = "Manuel King Klown") -> str:
    core_title = (doc.core_properties.title or "").strip()
    if core_title and core_title.lower() not in GENERIC_TITLES:
        return core_title[:80]

    for p in doc.paragraphs:
        t = clean_markdown_artifacts(p.text)
        if not t or t.lower() in GENERIC_TITLES:
            continue
        return t[:80].rstrip() + ("…" if len(t) > 80 else "")
    return fallback


def get_heading_1_style_name(doc: DocxDocument) -> str:
    style = style_lookup(doc, "Heading 1", "Titre 1")
    if style is not None:
        return style.name
    return "Heading 1"


def add_current_chapter_field(paragraph, doc: DocxDocument, placeholder: str = "Chapitre courant") -> None:
    style_name = get_heading_1_style_name(doc)
    safe_style_name = style_name.replace('"', r'\"')
    add_field(paragraph, f'STYLEREF "{safe_style_name}" \\* MERGEFORMAT', placeholder)


def format_run(run, cfg: KDPConfig, size: float = 7.6, color: str = GRAY, bold: bool = False) -> None:
    run.font.name = cfg.heading_font
    run.font.size = Pt(size)
    run.font.color.rgb = RGBColor.from_string(_hex(color))
    run.font.bold = bold


def add_header_footer(doc: DocxDocument, cfg: KDPConfig) -> None:
    title = get_document_title(doc)

    for section in doc.sections:
        header = section.header
        p = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
        p.text = ""

        if cfg.header_layout == "none":
            pass
        elif cfg.header_layout == "chapter":
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            add_current_chapter_field(p, doc, title)
            for run in p.runs:
                format_run(run, cfg, size=7.2, color=GRAY)
        elif cfg.header_layout == "brand-chapter":
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            brand = p.add_run("Univers-Cité King Klown - ")
            format_run(brand, cfg, size=7.2, color=BLACK, bold=True)
            add_current_chapter_field(p, doc, title)
            for run in p.runs:
                if run.font.size is None:
                    format_run(run, cfg, size=7.2, color=GRAY)
        elif cfg.header_layout == "minimal":
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run("King Klown")
            format_run(r, cfg, size=7.2, color=GRAY, bold=True)

        footer = section.footer
        fp = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        fp.text = ""
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER

        left = fp.add_run(cfg.footer_text + " - Page ")
        format_run(left, cfg, size=7.2, color=GRAY)
        add_field(fp, "PAGE", "1")
        fp.add_run(" / ")
        add_field(fp, "NUMPAGES", "1")

        for run in fp.runs:
            if run.font.size is None:
                format_run(run, cfg, size=7.2, color=GRAY)


def insert_toc_at_start(doc: DocxDocument) -> None:
    if not doc.paragraphs:
        doc.add_paragraph()

    first = doc.paragraphs[0]

    for p in doc.paragraphs[:10]:
        if p.text.strip().lower() in {"table des matières", "table of contents", "sommaire"}:
            return

    toc_title = first.insert_paragraph_before("Table des matières", style="Heading 1")
    toc_title.paragraph_format.page_break_before = False

    toc_p = first.insert_paragraph_before("")
    add_field(toc_p, r'TOC \o "1-3" \h \z \u', "Cliquez ici puis actualisez la table des matières.")

    sep = first.insert_paragraph_before("")
    sep.add_run().add_break(WD_BREAK.PAGE)


def standardize_tables(doc: DocxDocument, cfg: KDPConfig, stats: ProcessStats) -> None:
    normal = style_lookup(doc, "Normal", "normal")
    stats.tables = len(doc.tables)

    for table in doc.tables:
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = True
        try:
            table.style = "Table Grid"
        except Exception:
            pass

        for row_idx, row in enumerate(table.rows):
            for cell in row.cells:
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                set_cell_borders(cell, color=BORDER, size="5")
                set_cell_margins(cell)

                if row_idx == 0:
                    set_cell_shading(cell, LIGHT_GRAY)
                elif row_idx % 2 == 0:
                    set_cell_shading(cell, VERY_LIGHT_GRAY)
                else:
                    set_cell_shading(cell, WHITE)

                for p in cell.paragraphs:
                    if normal is not None:
                        p.style = normal
                    p.paragraph_format.space_after = Pt(1)
                    p.paragraph_format.line_spacing = 1.0
                    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT

                    for run in p.runs:
                        if cfg.clear_google_run_formatting:
                            clear_run_formatting_keep_emphasis(run)
                        run.font.name = cfg.body_font
                        run.font.size = Pt(8.2)
                        run.font.color.rgb = RGBColor.from_string(BLACK)
                        if row_idx == 0:
                            run.font.bold = True


def mark_header_rows_repeat(doc: DocxDocument) -> None:
    for table in doc.tables:
        if not table.rows:
            continue
        tr_pr = table.rows[0]._tr.get_or_add_trPr()
        tbl_header = tr_pr.find(qn("w:tblHeader"))
        if tbl_header is None:
            tbl_header = OxmlElement("w:tblHeader")
            tbl_header.set(qn("w:val"), "true")
            tr_pr.append(tbl_header)


# ---------------------------------------------------------------------------
# Pipeline
# ---------------------------------------------------------------------------

def backup_file(path: Path) -> Path:
    stamp = time.strftime("%Y%m%d_%H%M%S")
    backup = path.with_suffix(path.suffix + f".bak_{stamp}")
    backup.write_bytes(path.read_bytes())
    return backup


def validate_input_path(input_path: Path) -> None:
    if not input_path.exists():
        raise FileNotFoundError(f"Introuvable : {input_path}")
    if input_path.is_file() and input_path.suffix.lower() != ".docx":
        raise ValueError(f"Le fichier n'est pas un .docx : {input_path}")


def process_docx(input_path: Path, output_path: Path, cfg: KDPConfig,
                 dry_run: bool = False, in_place: bool = False) -> ProcessStats:
    validate_input_path(input_path)

    stats = ProcessStats(source=input_path, output=output_path, dry_run=dry_run)
    doc = Document(str(input_path))

    configure_sections(doc, cfg)
    configure_styles(doc, cfg)
    standardize_paragraphs(doc, cfg, stats)
    remove_excess_blank_paragraphs(doc, stats)

    if cfg.standardize_tables:
        standardize_tables(doc, cfg, stats)
        mark_header_rows_repeat(doc)

    if cfg.add_header_footer:
        add_header_footer(doc, cfg)

    if cfg.add_toc:
        insert_toc_at_start(doc)

    if dry_run:
        return stats

    output_path.parent.mkdir(parents=True, exist_ok=True)

    if in_place:
        backup = backup_file(input_path)
        stats.warnings.append(f"Sauvegarde créée : {backup.name}")

    doc.save(str(output_path))

    if cfg.add_toc or cfg.header_layout in {"chapter", "brand-chapter"}:
        try:
            enable_update_fields_on_open(output_path)
        except Exception as exc:
            stats.warnings.append(f"Champs Word : impossible d'activer updateFields : {exc}")

    if cfg.export_pdf:
        try:
            pdf_path = export_to_pdf(output_path)
            stats.warnings.append(f"PDF exporté : {pdf_path.name}")
        except Exception as exc:
            stats.warnings.append(f"Export PDF ignoré : {exc}")

    return stats


def iter_docx_files(path: Path, recursive: bool) -> Iterable[Path]:
    if path.is_file():
        if path.suffix.lower() == ".docx" and not path.name.startswith("~$"):
            yield path
        return

    pattern = "**/*.docx" if recursive else "*.docx"
    for p in sorted(path.glob(pattern)):
        if p.is_file() and not p.name.startswith("~$"):
            yield p


def output_for(input_file: Path, input_root: Path, output_root: Path, suffix: str, in_place: bool) -> Path:
    if in_place:
        return input_file

    if input_root.is_file():
        rel = Path(input_file.name)
    else:
        rel = input_file.relative_to(input_root)

    out = output_root / rel
    return out.with_name(out.stem + suffix + out.suffix)


def find_libreoffice() -> Optional[str]:
    for candidate in ("soffice", "libreoffice"):
        found = shutil.which(candidate)
        if found:
            return found

    win_paths = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for p in win_paths:
        if Path(p).exists():
            return p

    return None


def export_to_pdf(docx_path: Path) -> Path:
    soffice = find_libreoffice()
    if not soffice:
        raise RuntimeError("LibreOffice introuvable.")

    out_dir = docx_path.parent
    cmd = [
        soffice, "--headless", "--convert-to", "pdf",
        "--outdir", str(out_dir), str(docx_path),
    ]
    completed = subprocess.run(cmd, capture_output=True, text=True, check=False)
    if completed.returncode != 0:
        raise RuntimeError(completed.stderr.strip() or completed.stdout.strip() or "Échec LibreOffice")

    pdf_path = out_dir / (docx_path.stem + ".pdf")
    if not pdf_path.exists():
        raise RuntimeError("PDF attendu non créé.")
    return pdf_path


def run_batch(input_root: Path, output_root: Path, cfg: KDPConfig,
              recursive: bool = False, in_place: bool = False,
              suffix: str = "_KDP_6x9_BW", dry_run: bool = False,
              logger: Optional[Callable[[str], None]] = None) -> tuple[list[ProcessStats], int]:
    log = logger or print
    validate_input_path(input_root)

    files = list(iter_docx_files(input_root, recursive))
    if not files:
        log("Aucun fichier .docx trouvé.")
        return [], 0

    inside = cfg.inside_margin_in if cfg.inside_margin_in is not None else recommended_inside_margin_in(cfg.page_count)
    width, height = apply_bleed_size(cfg)
    log(f"{len(files)} fichier(s) à traiter.")
    log(f"KDP 6x9 BW | page={width:.3f}x{height:.3f} in | inside={inside:.3f} in | outside={cfg.outside_margin_in:.3f} in | body={cfg.body_size_pt:.1f} pt")

    stats_list: list[ProcessStats] = []
    errors = 0

    for input_file in files:
        out = output_for(input_file, input_root, output_root, suffix, in_place)
        try:
            stats = process_docx(input_file, out, cfg, dry_run=dry_run, in_place=in_place)
            stats_list.append(stats)
            prefix = "DRY-RUN" if dry_run else "OK"
            log(
                f"{prefix}  {input_file.name} -> {out}\n"
                f"      titres={stats.headings}, encadrés={stats.callouts}, "
                f"formules={stats.formula_blocks}, code={stats.code_lines}, "
                f"tableaux={stats.tables}, blancs_supprimés={stats.blank_paragraphs_removed}"
            )
            for warning in stats.warnings:
                log(f"      note: {warning}")
        except Exception as exc:
            errors += 1
            log(f"ERREUR  {input_file} : {exc}")

    log(f"Terminé avec {errors} erreur(s)." if errors else "Terminé.")
    return stats_list, errors


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def parse_args(argv: Optional[list[str]] = None):
    parser = argparse.ArgumentParser(description="Formatage Amazon KDP 6x9 noir et blanc pour .docx.")
    parser.add_argument("input", nargs="?", help="Fichier .docx ou dossier. Si absent : ouvre le GUI.")
    parser.add_argument("--output", "-o", default="kdp_6x9_bw", help="Dossier de sortie si --in-place n'est pas utilisé")
    parser.add_argument("--recursive", "-r", action="store_true", help="Traiter aussi les sous-dossiers")
    parser.add_argument("--in-place", action="store_true", help="Écraser les originaux avec backup automatique")
    parser.add_argument("--suffix", default="_KDP_6x9_BW", help="Suffixe ajouté aux fichiers générés")

    parser.add_argument("--page-count", type=int, default=205, help="Page count estimé pour calculer la marge intérieure KDP")
    parser.add_argument("--inside-margin", type=float, default=None, help="Marge intérieure en pouces. Override manuel.")
    parser.add_argument("--outside-margin", type=float, default=0.50, help="Marge extérieure en pouces")
    parser.add_argument("--top-margin", type=float, default=0.55, help="Marge haut en pouces")
    parser.add_argument("--bottom-margin", type=float, default=0.55, help="Marge bas en pouces")
    parser.add_argument("--bleed", action="store_true", help="Format avec bleed : 6.125 x 9.25. À éviter pour texte simple.")

    parser.add_argument("--body-size", type=float, default=9.6, help="Taille du corps en points")
    parser.add_argument("--body-alignment", choices=["left", "smart-justify", "justify"], default="left")
    parser.add_argument("--header-layout", choices=["chapter", "brand-chapter", "minimal", "none"], default="chapter")

    parser.add_argument("--toc", action="store_true", help="Insérer une table des matières")
    parser.add_argument("--no-header-footer", action="store_true", help="Ne pas ajouter header/footer")
    parser.add_argument("--no-chapter-breaks", action="store_true", help="Ne pas forcer les chapitres en nouvelle page")
    parser.add_argument("--keep-google-formatting", action="store_true", help="Conserver le formatage direct Google Docs")
    parser.add_argument("--no-tables", action="store_true", help="Ne pas standardiser les tableaux")
    parser.add_argument("--no-callout-boxes", action="store_true", help="Ne pas encadrer les blocs pédagogiques")
    parser.add_argument("--body-font", default="EB Garamond")
    parser.add_argument("--heading-font", default="EB Garamond")
    parser.add_argument("--footer-text", default="Univers-Cité King Klown")
    parser.add_argument("--dry-run", action="store_true", help="Analyser sans sauvegarder")
    parser.add_argument("--pdf", action="store_true", help="Exporter aussi PDF via LibreOffice")
    return parser.parse_args(argv)


def config_from_args(args) -> KDPConfig:
    cfg = KDPConfig()
    cfg.bleed = args.bleed
    cfg.page_count = args.page_count
    cfg.inside_margin_in = args.inside_margin
    cfg.outside_margin_in = args.outside_margin
    cfg.top_margin_in = args.top_margin
    cfg.bottom_margin_in = args.bottom_margin
    cfg.body_size_pt = args.body_size
    cfg.body_alignment = args.body_alignment
    cfg.header_layout = args.header_layout
    cfg.add_toc = args.toc
    cfg.add_header_footer = not args.no_header_footer
    cfg.chapter_page_breaks = not args.no_chapter_breaks
    cfg.clear_google_run_formatting = not args.keep_google_formatting
    cfg.standardize_tables = not args.no_tables
    cfg.add_callout_boxes = not args.no_callout_boxes
    cfg.body_font = args.body_font
    cfg.heading_font = args.heading_font
    cfg.footer_text = args.footer_text
    cfg.export_pdf = args.pdf
    return cfg


def cli_main(args) -> int:
    if not args.input:
        launch_gui()
        return 0

    input_root = Path(args.input).expanduser().resolve()
    output_root = Path(args.output).expanduser().resolve()
    cfg = config_from_args(args)

    _, errors = run_batch(
        input_root=input_root,
        output_root=output_root,
        cfg=cfg,
        recursive=args.recursive,
        in_place=args.in_place,
        suffix=args.suffix,
        dry_run=args.dry_run,
    )
    return 1 if errors else 0


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

class KDPGUI(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("King Klown KDP 6x9 BW Builder")
        self.geometry("900x660")
        self.minsize(800, 600)

        self.log_queue: queue.Queue[str] = queue.Queue()
        self.worker_thread: Optional[threading.Thread] = None

        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar(value=str(Path.cwd() / "kdp_6x9_bw"))

        self.suffix = tk.StringVar(value="_KDP_6x9_BW")
        self.body_font = tk.StringVar(value="EB Garamond")
        self.heading_font = tk.StringVar(value="EB Garamond")
        self.footer_text = tk.StringVar(value="Univers-Cité King Klown")

        self.page_count = tk.IntVar(value=205)
        self.body_size = tk.DoubleVar(value=9.6)
        self.inside_margin = tk.StringVar(value="auto")
        self.outside_margin = tk.DoubleVar(value=0.50)
        self.top_margin = tk.DoubleVar(value=0.55)
        self.bottom_margin = tk.DoubleVar(value=0.55)

        self.body_alignment = tk.StringVar(value="left")
        self.header_layout = tk.StringVar(value="chapter")

        self.recursive = tk.BooleanVar(value=False)
        self.in_place = tk.BooleanVar(value=False)
        self.bleed = tk.BooleanVar(value=False)
        self.add_toc = tk.BooleanVar(value=True)
        self.header_footer = tk.BooleanVar(value=True)
        self.chapter_breaks = tk.BooleanVar(value=True)
        self.clear_google_formatting = tk.BooleanVar(value=True)
        self.standardize_tables_var = tk.BooleanVar(value=True)
        self.callout_boxes = tk.BooleanVar(value=True)
        self.dry_run = tk.BooleanVar(value=False)
        self.export_pdf_var = tk.BooleanVar(value=True)

        self._build_ui()
        self.after(120, self._pump_log_queue)

    def _build_ui(self) -> None:
        root = ttk.Frame(self, padding=12)
        root.pack(fill="both", expand=True)

        source = ttk.LabelFrame(root, text="Source")
        source.pack(fill="x", pady=(0, 10))
        ttk.Label(source, text="Fichier ou dossier .docx").grid(row=0, column=0, sticky="w", padx=8, pady=8)
        ttk.Entry(source, textvariable=self.input_path).grid(row=0, column=1, sticky="ew", padx=8, pady=8)
        ttk.Button(source, text="Fichier…", command=self._choose_file).grid(row=0, column=2, padx=4, pady=8)
        ttk.Button(source, text="Dossier…", command=self._choose_folder).grid(row=0, column=3, padx=8, pady=8)
        source.columnconfigure(1, weight=1)

        out = ttk.LabelFrame(root, text="Sortie")
        out.pack(fill="x", pady=(0, 10))
        ttk.Label(out, text="Dossier de sortie").grid(row=0, column=0, sticky="w", padx=8, pady=8)
        self.output_entry = ttk.Entry(out, textvariable=self.output_path)
        self.output_entry.grid(row=0, column=1, sticky="ew", padx=8, pady=8)
        ttk.Button(out, text="Choisir…", command=self._choose_output).grid(row=0, column=2, padx=8, pady=8)
        ttk.Label(out, text="Suffixe").grid(row=1, column=0, sticky="w", padx=8, pady=8)
        ttk.Entry(out, textvariable=self.suffix, width=20).grid(row=1, column=1, sticky="w", padx=8, pady=8)
        out.columnconfigure(1, weight=1)

        layout = ttk.LabelFrame(root, text="KDP 6x9 noir et blanc")
        layout.pack(fill="x", pady=(0, 10))

        ttk.Label(layout, text="Pages estimées").grid(row=0, column=0, sticky="w", padx=8, pady=6)
        ttk.Entry(layout, textvariable=self.page_count, width=10).grid(row=0, column=1, sticky="w", padx=8, pady=6)

        ttk.Label(layout, text="Corps pt").grid(row=0, column=2, sticky="w", padx=8, pady=6)
        ttk.Entry(layout, textvariable=self.body_size, width=10).grid(row=0, column=3, sticky="w", padx=8, pady=6)

        ttk.Label(layout, text="Alignement").grid(row=0, column=4, sticky="w", padx=8, pady=6)
        ttk.Combobox(layout, textvariable=self.body_alignment, values=["left", "smart-justify", "justify"], state="readonly", width=14).grid(row=0, column=5, sticky="w", padx=8, pady=6)

        ttk.Label(layout, text="Marge intérieure").grid(row=1, column=0, sticky="w", padx=8, pady=6)
        ttk.Entry(layout, textvariable=self.inside_margin, width=10).grid(row=1, column=1, sticky="w", padx=8, pady=6)

        ttk.Label(layout, text="Extérieure").grid(row=1, column=2, sticky="w", padx=8, pady=6)
        ttk.Entry(layout, textvariable=self.outside_margin, width=10).grid(row=1, column=3, sticky="w", padx=8, pady=6)

        ttk.Label(layout, text="Haut / bas").grid(row=1, column=4, sticky="w", padx=8, pady=6)
        hb = ttk.Frame(layout)
        hb.grid(row=1, column=5, sticky="w", padx=8, pady=6)
        ttk.Entry(hb, textvariable=self.top_margin, width=7).pack(side="left")
        ttk.Label(hb, text=" / ").pack(side="left")
        ttk.Entry(hb, textvariable=self.bottom_margin, width=7).pack(side="left")

        ttk.Label(layout, text="Header").grid(row=2, column=0, sticky="w", padx=8, pady=6)
        ttk.Combobox(layout, textvariable=self.header_layout, values=["chapter", "brand-chapter", "minimal", "none"], state="readonly", width=16).grid(row=2, column=1, sticky="w", padx=8, pady=6)

        ttk.Label(layout, text="Police corps").grid(row=2, column=2, sticky="w", padx=8, pady=6)
        ttk.Entry(layout, textvariable=self.body_font, width=18).grid(row=2, column=3, sticky="w", padx=8, pady=6)

        ttk.Label(layout, text="Police titres").grid(row=2, column=4, sticky="w", padx=8, pady=6)
        ttk.Entry(layout, textvariable=self.heading_font, width=18).grid(row=2, column=5, sticky="w", padx=8, pady=6)

        ttk.Label(layout, text="Pied de page").grid(row=3, column=0, sticky="w", padx=8, pady=6)
        ttk.Entry(layout, textvariable=self.footer_text).grid(row=3, column=1, columnspan=5, sticky="ew", padx=8, pady=6)
        layout.columnconfigure(5, weight=1)

        checks = ttk.LabelFrame(root, text="Traitement")
        checks.pack(fill="x", pady=(0, 10))

        checkbox_data = [
            ("Récursif", self.recursive),
            ("Écraser originaux avec backup", self.in_place),
            ("Bleed 6.125 x 9.25", self.bleed),
            ("Table des matières", self.add_toc),
            ("En-tête / pied de page", self.header_footer),
            ("Sauts de page chapitres", self.chapter_breaks),
            ("Nettoyer formatage Google Docs", self.clear_google_formatting),
            ("Standardiser tableaux", self.standardize_tables_var),
            ("Encadrés gris", self.callout_boxes),
            ("Dry-run seulement", self.dry_run),
            ("Exporter PDF via LibreOffice", self.export_pdf_var),
        ]

        for idx, (label, var) in enumerate(checkbox_data):
            r = idx // 2
            c = (idx % 2) * 2
            ttk.Checkbutton(checks, text=label, variable=var, command=self._toggle_in_place).grid(
                row=r, column=c, columnspan=2, sticky="w", padx=8, pady=4
            )

        actions = ttk.Frame(root)
        actions.pack(fill="x", pady=(0, 10))
        self.run_button = ttk.Button(actions, text="Créer version KDP 6x9 BW", command=self._start_processing)
        self.run_button.pack(side="left")
        ttk.Button(actions, text="Effacer le log", command=self._clear_log).pack(side="left", padx=8)
        self.progress = ttk.Progressbar(actions, mode="indeterminate")
        self.progress.pack(side="right", fill="x", expand=True, padx=(8, 0))

        log_frame = ttk.LabelFrame(root, text="Log")
        log_frame.pack(fill="both", expand=True)
        self.log_text = tk.Text(log_frame, wrap="word", height=14)
        self.log_text.pack(side="left", fill="both", expand=True)
        scroll = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        scroll.pack(side="right", fill="y")
        self.log_text.configure(yscrollcommand=scroll.set)

    def _choose_file(self) -> None:
        path = filedialog.askopenfilename(
            title="Choisir un fichier .docx",
            filetypes=[("Word documents", "*.docx"), ("Tous les fichiers", "*.*")]
        )
        if path:
            self.input_path.set(path)

    def _choose_folder(self) -> None:
        path = filedialog.askdirectory(title="Choisir un dossier")
        if path:
            self.input_path.set(path)

    def _choose_output(self) -> None:
        path = filedialog.askdirectory(title="Choisir le dossier de sortie")
        if path:
            self.output_path.set(path)

    def _toggle_in_place(self) -> None:
        self.output_entry.configure(state="disabled" if self.in_place.get() else "normal")

    def _clear_log(self) -> None:
        self.log_text.delete("1.0", "end")

    def _log(self, msg: str) -> None:
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")

    def _thread_log(self, msg: str) -> None:
        self.log_queue.put(msg)

    def _pump_log_queue(self) -> None:
        try:
            while True:
                msg = self.log_queue.get_nowait()
                if msg == "__DONE__":
                    self.progress.stop()
                    self.run_button.configure(state="normal")
                else:
                    self._log(msg)
        except queue.Empty:
            pass
        self.after(120, self._pump_log_queue)

    def _inside_margin_value(self) -> Optional[float]:
        raw = self.inside_margin.get().strip().lower()
        if raw in {"", "auto"}:
            return None
        try:
            return float(raw)
        except ValueError:
            raise ValueError("Marge intérieure doit être 'auto' ou un nombre en pouces, ex. 0.56")

    def _start_processing(self) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showinfo("Traitement en cours", "Un traitement est déjà en cours.")
            return

        source = self.input_path.get().strip()
        if not source:
            messagebox.showerror("Source manquante", "Choisis un fichier .docx ou un dossier.")
            return

        input_root = Path(source).expanduser().resolve()
        output_root = Path(self.output_path.get().strip() or "kdp_6x9_bw").expanduser().resolve()

        if self.in_place.get():
            confirm = messagebox.askyesno(
                "Confirmer le mode in-place",
                "Le mode in-place écrase les fichiers originaux, avec une sauvegarde .bak automatique.\n\nContinuer ?"
            )
            if not confirm:
                return

        try:
            inside_margin = self._inside_margin_value()
        except ValueError as exc:
            messagebox.showerror("Marge invalide", str(exc))
            return

        cfg = KDPConfig(
            bleed=self.bleed.get(),
            page_count=int(self.page_count.get()),
            inside_margin_in=inside_margin,
            outside_margin_in=float(self.outside_margin.get()),
            top_margin_in=float(self.top_margin.get()),
            bottom_margin_in=float(self.bottom_margin.get()),
            body_size_pt=float(self.body_size.get()),
            body_alignment=self.body_alignment.get(),
            header_layout=self.header_layout.get(),
            body_font=self.body_font.get().strip() or "EB Garamond",
            heading_font=self.heading_font.get().strip() or "EB Garamond",
            footer_text=self.footer_text.get().strip() or "Univers-Cité King Klown",
            add_toc=self.add_toc.get(),
            add_header_footer=self.header_footer.get(),
            chapter_page_breaks=self.chapter_breaks.get(),
            clear_google_run_formatting=self.clear_google_formatting.get(),
            standardize_tables=self.standardize_tables_var.get(),
            add_callout_boxes=self.callout_boxes.get(),
            export_pdf=self.export_pdf_var.get(),
        )

        self.run_button.configure(state="disabled")
        self.progress.start(10)
        inside = cfg.inside_margin_in if cfg.inside_margin_in is not None else recommended_inside_margin_in(cfg.page_count)
        self._log("-" * 72)
        self._log(
            f"KDP 6x9 BW | page_count={cfg.page_count} | inside={inside:.3f} | "
            f"outside={cfg.outside_margin_in:.3f} | body={cfg.body_size_pt:.1f}pt | align={cfg.body_alignment}"
        )

        def worker() -> None:
            try:
                run_batch(
                    input_root=input_root,
                    output_root=output_root,
                    cfg=cfg,
                    recursive=self.recursive.get(),
                    in_place=self.in_place.get(),
                    suffix=self.suffix.get().strip() or "_KDP_6x9_BW",
                    dry_run=self.dry_run.get(),
                    logger=self._thread_log,
                )
            except Exception as exc:
                self._thread_log(f"ERREUR GÉNÉRALE : {exc}")
            finally:
                self._thread_log("__DONE__")

        self.worker_thread = threading.Thread(target=worker, daemon=True)
        self.worker_thread.start()


def launch_gui() -> None:
    app = KDPGUI()
    app.mainloop()


def main(argv: Optional[list[str]] = None) -> int:
    args = parse_args(argv)
    return cli_main(args)


if __name__ == "__main__":
    raise SystemExit(main())
