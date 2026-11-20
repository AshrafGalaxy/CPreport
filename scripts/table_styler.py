"""
table_styler.py
---------------
Defensive styling utilities for Microsoft Word tables:
- Zero table orphans via <w:cantSplit/>
- Header repeat across page overflows via <w:tblHeader/>
- Professional cell padding and alternating background shading
- Clean bordered layout matching academic publication guidelines
"""

from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Inches

def apply_defensive_table_formatting(table):
    """Enforces zero row page-split and repeats table header."""
    for i, row in enumerate(table.rows):
        trPr = row._tr.get_or_add_trPr()
        trPr.append(parse_xml(r'<w:cantSplit xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'))
        if i == 0:
            trPr.append(parse_xml(r'<w:tblHeader xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'))

def set_cell_shading(cell, color_hex="F1F5F9"):
    """Fills a table cell with the specified hex color."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    tcPr.append(shd)

def set_cell_padding(cell, top=100, bottom=100, left=150, right=150):
    """Sets internal cell margins in dxa (20 dxa = 1 pt)."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'  <w:top w:w="{top}" w:type="dxa"/>'
        f'  <w:bottom w:w="{bottom}" w:type="dxa"/>'
        f'  <w:left w:w="{left}" w:type="dxa"/>'
        f'  <w:right w:w="{right}" w:type="dxa"/>'
        f'</w:tcMar>'
    )
    tcPr.append(tcMar)

def apply_clean_table_borders(table):
    """Applies clean, subtle single-line borders to all cells."""
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        r'<w:tblBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        r'  <w:top w:val="single" w:sz="6" w:space="0" w:color="CBD5E1"/>'
        r'  <w:left w:val="single" w:sz="6" w:space="0" w:color="CBD5E1"/>'
        r'  <w:bottom w:val="single" w:sz="8" w:space="0" w:color="0F172A"/>'
        r'  <w:right w:val="single" w:sz="6" w:space="0" w:color="CBD5E1"/>'
        r'  <w:insideH w:val="single" w:sz="4" w:space="0" w:color="E2E8F0"/>'
        r'  <w:insideV w:val="single" w:sz="4" w:space="0" w:color="E2E8F0"/>'
        r'</w:tblBorders>'
    )
    tblPr.append(borders)
