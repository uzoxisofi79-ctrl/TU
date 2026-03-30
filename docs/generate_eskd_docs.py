#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Генератор DOCX-документов с оформлением по ЕСКД / ЕСТД.

Создаёт два файла:
  TD_Voin-Med.docx  — Технологическая карта
  KD_Voin-Med.docx  — Конструкторская документация

Форматирование:
  - Лист А4 (210 × 297 мм)
  - Рамка ЕСКД: левое поле 20 мм (подшивка), остальные 5 мм
  - Основная надпись (Форма 2 / 2а)  ГОСТ 2.104-2006
    · первый лист  — 40 мм, полная форма
    · последующие  — 15 мм, сокращённая форма
  - Шрифт: Times New Roman
"""

from docx import Document
from docx.shared import Mm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import re, os

# ─────────────────────────────────────────────────────────────────────────────
# Geometry (all in mm unless noted)
# ─────────────────────────────────────────────────────────────────────────────
PG_W, PG_H = 210, 297
MAR_L, MAR_T, MAR_R = 20, 5, 5
MAR_B_FIRST = 45       # bottom margin first page  (40 mm block + 5 mm edge)
MAR_B_CONT  = 20       # bottom margin cont. pages (15 mm block + 5 mm edge)
FOOT_DIST   = 5        # footer distance from page bottom edge
TEXT_W      = PG_W - MAR_L - MAR_R   # 185 mm

FONT = 'Times New Roman'

# Title-block column widths (mm), 13 columns, total = 185 mm
# Group A (revision): 7+7+10+10+10 = 44
# Group B (signatures): 20+15+10+10 = 55
# Group C (doc info): 66
# Group D (numbering): 5+8+7 = 20
# Total: 44+55+66+20 = 185 ✓
TB_COLS = [7, 7, 10, 10, 10,  20, 15, 10, 10,  66,  5, 8, 7]
assert sum(TB_COLS) == TEXT_W, f"TB_COLS sum={sum(TB_COLS)}, expected {TEXT_W}"

# ─────────────────────────────────────────────────────────────────────────────
# Low-level XML helpers
# ─────────────────────────────────────────────────────────────────────────────
def mm2tw(mm):
    """Millimetres → twips (1440 twips/inch ÷ 25.4 mm/inch ≈ 56.693 tw/mm)."""
    return int(round(mm * 56.693))


def _ensure_child(parent, tag):
    el = parent.find(qn(tag))
    if el is None:
        el = OxmlElement(tag)
        parent.append(el)
    return el


def set_run_font(run, size=Pt(10), bold=False, italic=False):
    run.font.name = FONT
    run.font.size = size
    run.font.bold = bold
    run.font.italic = italic
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    for attr in ('w:ascii', 'w:hAnsi', 'w:cs'):
        rFonts.set(qn(attr), FONT)


def set_para_spacing(para, before=0, after=0, line=240):
    pPr = para._p.get_or_add_pPr()
    sp = pPr.find(qn('w:spacing'))
    if sp is None:
        sp = OxmlElement('w:spacing')
        pPr.append(sp)
    sp.set(qn('w:before'), str(before))
    sp.set(qn('w:after'),  str(after))
    sp.set(qn('w:line'),   str(line))
    sp.set(qn('w:lineRule'), 'auto')


def set_para_indent(para, first=0, left=0):
    pPr = para._p.get_or_add_pPr()
    ind = pPr.find(qn('w:ind'))
    if ind is None:
        ind = OxmlElement('w:ind')
        pPr.append(ind)
    if first:
        ind.set(qn('w:firstLine'), str(first))
    if left:
        ind.set(qn('w:left'), str(left))


def set_cell_borders(cell, top=True, bottom=True, left=True, right=True, sz=4):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for old in tcPr.findall(qn('w:tcBorders')):
        tcPr.remove(old)
    tcB = OxmlElement('w:tcBorders')
    for edge, show in [('top', top), ('bottom', bottom), ('start', left), ('end', right)]:
        b = OxmlElement(f'w:{edge}')
        if show:
            b.set(qn('w:val'), 'single')
            b.set(qn('w:sz'),  str(sz))
            b.set(qn('w:space'), '0')
            b.set(qn('w:color'), '000000')
        else:
            b.set(qn('w:val'), 'nil')
        tcB.append(b)
    tcPr.append(tcB)


def set_cell_width(cell, width_mm):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for old in tcPr.findall(qn('w:tcW')):
        tcPr.remove(old)
    w = OxmlElement('w:tcW')
    w.set(qn('w:w'),    str(mm2tw(width_mm)))
    w.set(qn('w:type'), 'dxa')
    tcPr.append(w)


def set_cell_vAlign(cell, val='center'):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    va = _ensure_child(tcPr, 'w:vAlign')
    va.set(qn('w:val'), val)


def set_cell_margins(cell, left_tw=28, right_tw=28, top_tw=0, bottom_tw=0):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for old in tcPr.findall(qn('w:tcMar')):
        tcPr.remove(old)
    tcM = OxmlElement('w:tcMar')
    for side, val in [('top', top_tw), ('left', left_tw), ('bottom', bottom_tw), ('right', right_tw)]:
        m = OxmlElement(f'w:{side}')
        m.set(qn('w:w'),    str(val))
        m.set(qn('w:type'), 'dxa')
        tcM.append(m)
    tcPr.append(tcM)


def cell_write(cell, text='', size=Pt(7), bold=False, italic=False,
               align=WD_ALIGN_PARAGRAPH.CENTER, va='center'):
    """Clear a cell and write text with given formatting."""
    for p in cell.paragraphs:
        p.clear()
    set_cell_vAlign(cell, va)
    set_cell_margins(cell)
    para = cell.paragraphs[0]
    para.alignment = align
    set_para_spacing(para, 0, 0, 240)
    if text:
        run = para.add_run(text)
        set_run_font(run, size=size, bold=bold, italic=italic)


def set_row_height(row, mm, exact=True):
    trPr = row._tr.get_or_add_trPr()
    for old in trPr.findall(qn('w:trHeight')):
        trPr.remove(old)
    trH = OxmlElement('w:trHeight')
    trH.set(qn('w:val'),   str(mm2tw(mm)))
    trH.set(qn('w:hRule'), 'exact' if exact else 'atLeast')
    trPr.append(trH)


def make_table(container, nrows, ncols, width_mm, col_widths_mm):
    """Create a table with no default borders and explicit column widths."""
    tbl = container.add_table(rows=nrows, cols=ncols, width=Mm(width_mm))
    tbl.alignment = WD_TABLE_ALIGNMENT.LEFT

    # Width
    tblXml  = tbl._tbl
    tblPr   = tblXml.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tblXml.insert(0, tblPr)
    for old in tblPr.findall(qn('w:tblW')):
        tblPr.remove(old)
    tW = OxmlElement('w:tblW')
    tW.set(qn('w:w'),    str(mm2tw(width_mm)))
    tW.set(qn('w:type'), 'dxa')
    tblPr.append(tW)

    # No default borders
    for old in tblPr.findall(qn('w:tblBorders')):
        tblPr.remove(old)
    tBrd = OxmlElement('w:tblBorders')
    for e in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        b = OxmlElement(f'w:{e}')
        b.set(qn('w:val'), 'nil')
        tBrd.append(b)
    tblPr.append(tBrd)

    # Grid
    for old in tblXml.findall(qn('w:tblGrid')):
        tblXml.remove(old)
    grid = OxmlElement('w:tblGrid')
    for w in col_widths_mm:
        gc = OxmlElement('w:gridCol')
        gc.set(qn('w:w'), str(mm2tw(w)))
        grid.append(gc)
    tblPr.addnext(grid)

    # Assign cell widths row-by-row
    for row in tbl.rows:
        for j, cell in enumerate(row.cells):
            if j < len(col_widths_mm):
                set_cell_width(cell, col_widths_mm[j])

    return tbl


# ─────────────────────────────────────────────────────────────────────────────
# Page setup
# ─────────────────────────────────────────────────────────────────────────────
def setup_section(section, bottom_mm):
    section.page_width      = Mm(PG_W)
    section.page_height     = Mm(PG_H)
    section.left_margin     = Mm(MAR_L)
    section.top_margin      = Mm(MAR_T)
    section.right_margin    = Mm(MAR_R)
    section.bottom_margin   = Mm(bottom_mm)
    section.header_distance = Mm(3)
    section.footer_distance = Mm(FOOT_DIST)
    section.different_first_page_header_footer = True


def add_page_borders(section):
    """
    ЕСКД рамка: линии в 20/5/5/5 мм от края листа.
    w:pgBorders offsetFrom='page'; w:space в пунктах (1 мм ≈ 2.835 pt).
    Левое поле 20 мм → ≈ 57 pt; остальные 5 мм → ≈ 14 pt.
    """
    sectPr = section._sectPr
    for old in sectPr.findall(qn('w:pgBorders')):
        sectPr.remove(old)
    pgB = OxmlElement('w:pgBorders')
    pgB.set(qn('w:offsetFrom'), 'page')
    for name, spc in (('top', 14), ('left', 57), ('bottom', 14), ('right', 14)):
        b = OxmlElement(f'w:{name}')
        b.set(qn('w:val'),   'single')
        b.set(qn('w:sz'),    '6')    # 0.75 pt line
        b.set(qn('w:space'), str(spc))
        b.set(qn('w:color'), '000000')
        pgB.append(b)
    sectPr.append(pgB)


# ─────────────────────────────────────────────────────────────────────────────
# ЕСКД основная надпись — первый лист (Форма 2, высота 40 мм)
# ─────────────────────────────────────────────────────────────────────────────
#
#  Col widths: A1=7 A2=7 A3=10 A4=10 A5=10 | B1=20 B2=15 B3=10 B4=10 | C=66 | D1=5 D2=8 D3=7
#  Rows (mm):   R0=8  R1=5  R2=5  R3=5  R4=5  R5=7  R6=5  = 40 mm
#
#  R0: [A1-A5 blank] Разраб  | name | sign | date | [C: DocName (rows 0-1)] | Лит. | Лист | Листов
#  R1: [A1-A5 blank] Пров.   | name | sign | date | [C: cont.]              | val  | val  | val
#  R2: [A1-A5 blank] blank   | name | sign | date | [C: DocDesig (rows 2-3)]| mrg  | mrg  | mrg
#  R3: [A1-A5 blank] Н.контр.| name | sign | date | [C: cont.]              | mrg  | mrg  | mrg
#  R4: [A1-A5 blank] Утв.    | name | sign | date | [C: Company]            | mrg  | mrg  | mrg
#  R5: [A1-A5 blank] blank   | name | sign | date | [C: blank]              | mrg  | mrg  | mrg
#  R6: Изм | Лист | №докум | Подп | Дата | [B1-B4 merged] | [C blank] | [D1-D3 merged]
#
def build_title_first(footer, doc_num, doc_name, company, lit='О',
                      sheet_val='1', total_val='', author=''):
    """Full ESKD Form-2 title block for the first page (40 mm)."""
    ROW_H = [8, 5, 5, 5, 5, 7, 5]   # total = 40 mm

    tbl = make_table(footer, nrows=7, ncols=13,
                     width_mm=TEXT_W, col_widths_mm=TB_COLS)

    for i, row in enumerate(tbl.rows):
        set_row_height(row, ROW_H[i])

    # ── merges ──────────────────────────────────────────────────────────────
    # C column (index 9): DocName spans rows 0-1
    tbl.cell(0, 9).merge(tbl.cell(1, 9))
    # C column: DocDesig spans rows 2-3
    tbl.cell(2, 9).merge(tbl.cell(3, 9))
    # D columns (10-12): rows 1-5 merged vertically (values under headers)
    for d in (10, 11, 12):
        tbl.cell(1, d).merge(tbl.cell(5, d))
    # Bottom row (R6): B merged horizontally (cols 5-8)
    tbl.cell(6, 5).merge(tbl.cell(6, 8))
    # Bottom row (R6): D merged horizontally (cols 10-12)
    tbl.cell(6, 10).merge(tbl.cell(6, 12))

    SL = Pt(6)   # label size
    SV = Pt(7)   # value size
    SN = Pt(8.5) # doc-name size

    # ── Group A (revision cols, rows 0-5, blank) ────────────────────────────
    for r in range(6):
        for c in range(5):
            cell_write(tbl.cell(r, c), '', size=SV)
            set_cell_borders(tbl.cell(r, c))

    # ── Group A row 6 (headers) ─────────────────────────────────────────────
    for c, hdr in enumerate(['Изм.', 'Лист', '№ докум.', 'Подп.', 'Дата']):
        cell_write(tbl.cell(6, c), hdr, size=SL, bold=True)
        set_cell_borders(tbl.cell(6, c))

    # ── Group B (signatures, rows 0-5) ──────────────────────────────────────
    roles = ['Разраб.', 'Пров.', '', 'Н.контр.', 'Утв.', '']
    names = [author,    '',      '', '',          '',      '']
    for r in range(6):
        cell_write(tbl.cell(r, 5), roles[r], size=SL, bold=bool(roles[r]),
                   align=WD_ALIGN_PARAGRAPH.LEFT)
        set_cell_borders(tbl.cell(r, 5))
        cell_write(tbl.cell(r, 6), names[r], size=SV,
                   align=WD_ALIGN_PARAGRAPH.LEFT)
        set_cell_borders(tbl.cell(r, 6))
        cell_write(tbl.cell(r, 7), '', size=SV)   # подпись
        set_cell_borders(tbl.cell(r, 7))
        cell_write(tbl.cell(r, 8), '', size=SV)   # дата
        set_cell_borders(tbl.cell(r, 8))

    # Group B row 6: merged blank
    cell_write(tbl.cell(6, 5), '', size=SV)
    set_cell_borders(tbl.cell(6, 5))

    # ── Group C (doc info) ──────────────────────────────────────────────────
    # Row 0 (merged 0-1): document name
    cell_write(tbl.cell(0, 9), doc_name, size=SN, bold=True)
    set_cell_borders(tbl.cell(0, 9))

    # Row 2 (merged 2-3): document designation
    cell_write(tbl.cell(2, 9), doc_num, size=SV, bold=True)
    set_cell_borders(tbl.cell(2, 9))

    # Row 4: company
    cell_write(tbl.cell(4, 9), company, size=SL)
    set_cell_borders(tbl.cell(4, 9))

    # Row 5: blank
    cell_write(tbl.cell(5, 9), '', size=SV)
    set_cell_borders(tbl.cell(5, 9))

    # Row 6: blank
    cell_write(tbl.cell(6, 9), '', size=SV)
    set_cell_borders(tbl.cell(6, 9))

    # ── Group D (numbering) ─────────────────────────────────────────────────
    # Row 0: headers
    for c, lbl in zip((10, 11, 12), ('Лит.', 'Лист', 'Листов')):
        cell_write(tbl.cell(0, c), lbl, size=SL, bold=True)
        set_cell_borders(tbl.cell(0, c))

    # Row 1 (merged 1-5): values
    cell_write(tbl.cell(1, 10), lit,        size=SV)
    set_cell_borders(tbl.cell(1, 10))
    cell_write(tbl.cell(1, 11), sheet_val,  size=SV)
    set_cell_borders(tbl.cell(1, 11))
    cell_write(tbl.cell(1, 12), total_val,  size=SV)
    set_cell_borders(tbl.cell(1, 12))

    # Row 6: merged blank
    cell_write(tbl.cell(6, 10), '', size=SV)
    set_cell_borders(tbl.cell(6, 10))

    return tbl


# ─────────────────────────────────────────────────────────────────────────────
# ЕСКД основная надпись — последующие листы (Форма 2а, высота 15 мм)
# ─────────────────────────────────────────────────────────────────────────────
#
#  Same 13 columns.  2 rows:
#  R0 (8 mm): [A blank] [B1-B4 merged blank] [C: doc_num+doc_name_short] [D1: Лист] [D2-3 blank]
#  R1 (7 mm): Изм | Лист | №докум | Подп | Дата | [B merged] | [C blank] | [D merged]
#
def build_title_cont(footer, doc_num, doc_name_short, sheet_val):
    """Shortened ESKD Form-2a title block for continuation pages (15 mm)."""
    ROW_H = [8, 7]   # total 15 mm

    tbl = make_table(footer, nrows=2, ncols=13,
                     width_mm=TEXT_W, col_widths_mm=TB_COLS)

    for i, row in enumerate(tbl.rows):
        set_row_height(row, ROW_H[i])

    # Row 0: B cols merged (5-8), D cols (11-12) merged into one
    tbl.cell(0, 5).merge(tbl.cell(0, 8))
    tbl.cell(0, 11).merge(tbl.cell(0, 12))

    # Row 1: B cols merged (5-8), C blank, D cols merged (10-12)
    tbl.cell(1, 5).merge(tbl.cell(1, 8))
    tbl.cell(1, 10).merge(tbl.cell(1, 12))

    SL = Pt(6)
    SV = Pt(7)

    # Row 0
    for c in range(5):
        cell_write(tbl.cell(0, c), '', size=SV)
        set_cell_borders(tbl.cell(0, c))

    cell_write(tbl.cell(0, 5), '', size=SV)
    set_cell_borders(tbl.cell(0, 5))

    label = f'{doc_num}  {doc_name_short}'
    cell_write(tbl.cell(0, 9), label, size=SV)
    set_cell_borders(tbl.cell(0, 9))

    cell_write(tbl.cell(0, 10), '', size=SV)
    set_cell_borders(tbl.cell(0, 10))
    cell_write(tbl.cell(0, 11), sheet_val, size=SV)
    set_cell_borders(tbl.cell(0, 11))

    # Row 1 headers
    for c, hdr in enumerate(['Изм.', 'Лист', '№ докум.', 'Подп.', 'Дата']):
        cell_write(tbl.cell(1, c), hdr, size=SL, bold=True)
        set_cell_borders(tbl.cell(1, c))

    cell_write(tbl.cell(1, 5), '', size=SV)
    set_cell_borders(tbl.cell(1, 5))
    cell_write(tbl.cell(1, 9), '', size=SV)
    set_cell_borders(tbl.cell(1, 9))
    cell_write(tbl.cell(1, 10), '', size=SV)
    set_cell_borders(tbl.cell(1, 10))

    return tbl


# ─────────────────────────────────────────────────────────────────────────────
# Document styling helpers
# ─────────────────────────────────────────────────────────────────────────────
def add_heading(doc, text, level):
    """Add heading with Times New Roman, no extra spacing."""
    sizes  = {1: Pt(13), 2: Pt(12), 3: Pt(11)}
    aligns = {1: WD_ALIGN_PARAGRAPH.CENTER,
              2: WD_ALIGN_PARAGRAPH.LEFT,
              3: WD_ALIGN_PARAGRAPH.LEFT}
    p = doc.add_paragraph()
    p.alignment = aligns.get(level, WD_ALIGN_PARAGRAPH.LEFT)
    set_para_spacing(p, before=120 if level == 1 else 80,
                     after=60 if level == 1 else 40)
    run = p.add_run(text)
    set_run_font(run, size=sizes.get(level, Pt(10)), bold=True)
    return p


def add_body_para(doc, text, bold=False, first_line_indent=True):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    set_para_spacing(p, before=0, after=40, line=276)
    if first_line_indent:
        set_para_indent(p, first=mm2tw(10))
    run = p.add_run(text)
    set_run_font(run, size=Pt(10), bold=bold)
    return p


def add_list_item(doc, text, level=1, numbered=False, num_val=''):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    left_tw = mm2tw(10 * level)
    set_para_spacing(p, before=0, after=20, line=260)
    set_para_indent(p, left=left_tw, first=-mm2tw(5))
    bullet = f'{num_val}' if numbered else '–'
    run = p.add_run(f'{bullet}  {text}')
    set_run_font(run, size=Pt(10))
    return p


def add_md_table(doc, header_row, data_rows):
    """Render a markdown table as a DOCX table."""
    ncols = len(header_row)
    if ncols == 0:
        return

    tbl = doc.add_table(rows=1 + len(data_rows), cols=ncols)
    tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
    tbl.style = 'Table Grid'

    # Fix table width
    tblXml = tbl._tbl
    tblPr  = tblXml.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tblXml.insert(0, tblPr)
    for old in tblPr.findall(qn('w:tblW')):
        tblPr.remove(old)
    tW = OxmlElement('w:tblW')
    tW.set(qn('w:w'),    str(mm2tw(TEXT_W)))
    tW.set(qn('w:type'), 'dxa')
    tblPr.append(tW)

    # Table layout: auto or fixed
    tblLayout = OxmlElement('w:tblLayout')
    tblLayout.set(qn('w:type'), 'autofit')
    tblPr.append(tblLayout)

    def write_cell(cell, text, bold=False):
        """Write cell text, parsing inline **bold** markers."""
        cell.text = ''
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        set_para_spacing(p, 0, 0, 240)
        set_cell_margins(cell, left_tw=56, right_tw=56, top_tw=28, bottom_tw=28)
        # Parse **bold** spans
        parts = re.split(r'(\*\*[^*]+\*\*)', text)
        for part in parts:
            if part.startswith('**') and part.endswith('**'):
                run = p.add_run(part[2:-2])
                set_run_font(run, size=Pt(9), bold=True)
            else:
                run = p.add_run(part)
                set_run_font(run, size=Pt(9), bold=bold)

    # Header
    for j, hdr in enumerate(header_row):
        write_cell(tbl.rows[0].cells[j], _strip_md(hdr), bold=True)

    # Data rows
    for i, row_data in enumerate(data_rows):
        for j in range(ncols):
            val = row_data[j] if j < len(row_data) else ''
            write_cell(tbl.rows[i + 1].cells[j], val)

    # Spacing after table
    p = doc.add_paragraph()
    set_para_spacing(p, 0, 40, 240)


def _strip_md(text):
    """Remove basic markdown formatting characters."""
    text = re.sub(r'\*\*(.*?)\*\*', r'\1', text)
    text = re.sub(r'\*(.*?)\*',     r'\1', text)
    text = re.sub(r'`(.*?)`',       r'\1', text)
    text = text.strip()
    return text


# ─────────────────────────────────────────────────────────────────────────────
# Markdown → DOCX content renderer
# ─────────────────────────────────────────────────────────────────────────────
def render_md(doc, md_text):
    """
    Parse a markdown string and add content to *doc*.
    Handles: # headings, tables, bullet lists, numbered lists,
             blockquotes (>), horizontal rules, bold inline, paragraphs.
    """
    lines = md_text.splitlines()
    i = 0

    while i < len(lines):
        raw  = lines[i]
        line = raw.strip()

        # ── heading ──────────────────────────────────────────────────────────
        m = re.match(r'^(#{1,3})\s+(.*)', line)
        if m:
            level = len(m.group(1))
            text  = _strip_md(m.group(2))
            add_heading(doc, text, level)
            i += 1
            continue

        # ── horizontal rule ──────────────────────────────────────────────────
        if re.match(r'^-{3,}$', line) or re.match(r'^={3,}$', line):
            p = doc.add_paragraph()
            set_para_spacing(p, 20, 20, 240)
            i += 1
            continue

        # ── markdown table ───────────────────────────────────────────────────
        if line.startswith('|') and '|' in line[1:]:
            table_lines = []
            while i < len(lines) and lines[i].strip().startswith('|'):
                table_lines.append(lines[i].strip())
                i += 1
            # Parse
            rows = []
            for tl in table_lines:
                cols = [c.strip() for c in tl.strip('|').split('|')]
                rows.append(cols)
            # Drop separator rows (cells that are only dashes/colons)
            rows = [r for r in rows
                    if not all(re.match(r'^[-:]+$', c) for c in r if c)]
            if rows:
                add_md_table(doc, rows[0], rows[1:])
            continue

        # ── blockquote (>) ───────────────────────────────────────────────────
        if line.startswith('>'):
            text = _strip_md(line.lstrip('> '))
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            set_para_spacing(p, 40, 40, 260)
            set_para_indent(p, left=mm2tw(10))
            pPr = p._p.get_or_add_pPr()
            pBdr = OxmlElement('w:pBdr')
            bdr  = OxmlElement('w:left')
            bdr.set(qn('w:val'),   'single')
            bdr.set(qn('w:sz'),    '8')
            bdr.set(qn('w:space'), '4')
            bdr.set(qn('w:color'), '000000')
            pBdr.append(bdr)
            pPr.append(pBdr)
            run = p.add_run(text)
            set_run_font(run, size=Pt(10), italic=True)
            i += 1
            continue

        # ── numbered list ────────────────────────────────────────────────────
        m = re.match(r'^(\s*)(\d+)\.\s+(.*)', raw)
        if m:
            indent_str = m.group(1)
            num_str    = m.group(2)
            text       = _strip_md(m.group(3))
            level      = len(indent_str) // 3 + 1
            add_list_item(doc, text, level=level, numbered=True,
                          num_val=num_str + '.')
            i += 1
            continue

        # ── bullet list ──────────────────────────────────────────────────────
        m = re.match(r'^(\s*)[-*]\s+(.*)', raw)
        if m:
            indent_str = m.group(1)
            text       = _strip_md(m.group(2))
            level      = len(indent_str) // 2 + 1
            add_list_item(doc, text, level=level)
            i += 1
            continue

        # ── code block (``` ... ```) ─────────────────────────────────────────
        if line.startswith('```'):
            i += 1
            block_lines = []
            while i < len(lines) and not lines[i].strip().startswith('```'):
                block_lines.append(lines[i])
                i += 1
            i += 1  # skip closing ```
            # Render as preformatted paragraph (Courier-style)
            if block_lines:
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                set_para_spacing(p, 0, 0, 240)
                set_para_indent(p, left=mm2tw(10))
                for bl_line in block_lines:
                    run = p.add_run(bl_line + '\n')
                    run.font.name = 'Courier New'
                    run.font.size = Pt(8)
            continue

        # ── blank line ───────────────────────────────────────────────────────
        if not line:
            i += 1
            continue

        # ── normal paragraph (may contain **bold**) ──────────────────────────
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        set_para_spacing(p, 0, 40, 276)
        set_para_indent(p, first=mm2tw(10))

        # Split on **…** bold markers
        parts = re.split(r'(\*\*[^*]+\*\*)', line)
        for part in parts:
            if part.startswith('**') and part.endswith('**'):
                run = p.add_run(part[2:-2])
                set_run_font(run, size=Pt(10), bold=True)
            else:
                run = p.add_run(part)
                set_run_font(run, size=Pt(10))

        i += 1

    return doc


# ─────────────────────────────────────────────────────────────────────────────
# Document builder
# ─────────────────────────────────────────────────────────────────────────────
def build_document(md_path, out_path, doc_num, doc_name, doc_name_short,
                   company, author='', lit='О'):
    """
    Build an ESKD-formatted DOCX from a markdown source file.

    Parameters
    ----------
    md_path        : path to the source .md file
    out_path       : path to write the output .docx
    doc_num        : document designation (обозначение документа)
    doc_name       : full document name for title block
    doc_name_short : short name for continuation pages
    company        : organisation name
    author         : developer's name (Разраб.)
    lit            : document literal (Лит.), default 'О'
    """
    with open(md_path, 'r', encoding='utf-8') as f:
        md_text = f.read()

    doc = Document()

    # ── Section 0: first page ────────────────────────────────────────────────
    sec0 = doc.sections[0]
    setup_section(sec0, bottom_mm=MAR_B_FIRST)
    add_page_borders(sec0)

    # First-page footer: full title block
    fp_footer = sec0.first_page_footer
    build_title_first(fp_footer, doc_num=doc_num, doc_name=doc_name,
                      company=company, lit=lit,
                      sheet_val='1', total_val='', author=author)

    # Default (continuation) footer: short title block
    def_footer = sec0.footer
    build_title_cont(def_footer, doc_num=doc_num,
                     doc_name_short=doc_name_short, sheet_val='')

    # Add page-number field in a tiny run after the title block
    for footer_obj in (fp_footer, def_footer):
        p = footer_obj.add_paragraph()
        set_para_spacing(p, 0, 0, 240)
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = p.add_run()
        run.font.name = FONT
        run.font.size = Pt(7)
        fldChar1 = OxmlElement('w:fldChar')
        fldChar1.set(qn('w:fldCharType'), 'begin')
        instrText = OxmlElement('w:instrText')
        instrText.text = ' PAGE '
        fldChar2 = OxmlElement('w:fldChar')
        fldChar2.set(qn('w:fldCharType'), 'end')
        run._r.append(fldChar1)
        run._r.append(instrText)
        run._r.append(fldChar2)

    # ── Render content ───────────────────────────────────────────────────────
    # Remove default empty paragraph so content starts cleanly
    if doc.paragraphs:
        first_p = doc.paragraphs[0]
        first_p.text = ''
        set_para_spacing(first_p, 0, 0, 240)

    render_md(doc, md_text)

    # ── Save ─────────────────────────────────────────────────────────────────
    doc.save(out_path)
    print(f'Saved: {out_path}')


# ─────────────────────────────────────────────────────────────────────────────
# Entry point
# ─────────────────────────────────────────────────────────────────────────────
if __name__ == '__main__':
    HERE = os.path.dirname(os.path.abspath(__file__))

    # ── ТД ──────────────────────────────────────────────────────────────────
    build_document(
        md_path       = os.path.join(HERE, 'TD_tekhnologicheskaya_karta.md'),
        out_path      = os.path.join(HERE, 'TD_Voin-Med.docx'),
        doc_num       = 'ТД-ВМ-001',
        doc_name      = ('Жгут медицинский кровоостанавливающий\n'
                         'одноразового использования «Воин-Мед»\n'
                         'Технологическая карта'),
        doc_name_short= 'ТД Воин-Мед. Технологическая карта',
        company       = '________________',
        author        = '________________',
        lit           = 'О',
    )

    # ── КД ──────────────────────────────────────────────────────────────────
    build_document(
        md_path       = os.path.join(HERE, 'KD_konstruktorskaya_dokumentaciya.md'),
        out_path      = os.path.join(HERE, 'KD_Voin-Med.docx'),
        doc_num       = 'КД-ВМ-001',
        doc_name      = ('Жгут медицинский кровоостанавливающий\n'
                         'одноразового использования «Воин-Мед»\n'
                         'Конструкторская документация'),
        doc_name_short= 'КД Воин-Мед. Конструкторская документация',
        company       = '________________',
        author        = '________________',
        lit           = 'О',
    )
