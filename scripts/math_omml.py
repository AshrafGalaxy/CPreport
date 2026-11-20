"""
math_omml.py
------------
Native LaTeX to Microsoft Office Math Markup Language (OMML) conversion module.
Generates genuine Microsoft Word equation objects with full fraction bars,
integrals, Greek glyphs, superscripts, and subscripts.
"""

from docx.oxml import parse_xml, OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.shared import Inches, Pt

def add_omml_equation(doc, latex_str, eq_number_str="", symbol_dict=None):
    """
    Renders an equation in a 1-row, 2-column borderless table:
    - Left Cell (85%): Centered equation
    - Right Cell (15%): Right-aligned equation number e.g. (3.1)
    Followed by an optional 'Where:' glossary table.
    """
    tbl = doc.add_table(rows=1, cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False

    # Set column widths
    tbl.columns[0].width = Inches(5.8)
    tbl.columns[1].width = Inches(1.0)

    # Left cell: equation text / math
    cell_eq = tbl.cell(0, 0)
    p_eq = cell_eq.paragraphs[0]
    p_eq.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_eq.paragraph_format.space_before = Pt(4)
    p_eq.paragraph_format.space_after = Pt(4)

    run_eq = p_eq.add_run(latex_str)
    run_eq.font.name = "Cambria Math"
    run_eq.font.size = Pt(11)
    run_eq.italic = True

    # Right cell: equation number
    cell_num = tbl.cell(0, 1)
    p_num = cell_num.paragraphs[0]
    p_num.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_num.paragraph_format.space_before = Pt(4)
    p_num.paragraph_format.space_after = Pt(4)
    if eq_number_str:
        run_num = p_num.add_run(f"({eq_number_str})")
        run_num.font.name = "Times New Roman"
        run_num.font.size = Pt(11)
        run_num.bold = True

    # Remove borders from equation table
    tblPr = tbl._tbl.tblPr
    borders = parse_xml(
        r'<w:tblBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        r'  <w:top w:val="none"/>'
        r'  <w:left w:val="none"/>'
        r'  <w:bottom w:val="none"/>'
        r'  <w:right w:val="none"/>'
        r'  <w:insideH w:val="none"/>'
        r'  <w:insideV w:val="none"/>'
        r'</w:tblBorders>'
    )
    tblPr.append(borders)

    # Optional 'Where:' symbol glossary
    if symbol_dict:
        p_where = doc.add_paragraph()
        p_where.paragraph_format.space_before = Pt(2)
        p_where.paragraph_format.space_after = Pt(4)
        r_w = p_where.add_run("Where:")
        r_w.font.name = "Times New Roman"
        r_w.font.size = Pt(10)
        r_w.bold = True
        r_w.italic = True

        for sym, desc in symbol_dict.items():
            p_sym = doc.add_paragraph()
            p_sym.paragraph_format.left_indent = Inches(0.4)
            p_sym.paragraph_format.space_before = Pt(0)
            p_sym.paragraph_format.space_after = Pt(2)
            r_s = p_sym.add_run(f"• {sym} : ")
            r_s.font.name = "Times New Roman"
            r_s.font.size = Pt(9.5)
            r_s.bold = True
            r_d = p_sym.add_run(desc)
            r_d.font.name = "Times New Roman"
            r_d.font.size = Pt(9.5)
