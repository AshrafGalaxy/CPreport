"""
figure_manager.py
-----------------
Manages figure scaling, centering, and defensive caption binding:
- Scales architecture diagrams and plots to printable margins (6.8" - 7.1")
- Binds figure paragraph to caption paragraph via keep_with_next = True
- Centers caption and numbers figures according to academic convention (Figure X.Y)
"""

import os
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

def add_scaled_figure(doc, image_path, caption_text, figure_number="", width_inches=6.85):
    """
    Inserts a high-resolution figure bound cleanly to its caption:
    - Image paragraph keeps with next
    - Centered caption below figure
    """
    if not os.path.exists(image_path):
        print(f"Warning: Figure not found at {image_path}")
        return None

    # Image paragraph with keep_with_next
    p_img = doc.add_paragraph()
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img.paragraph_format.space_before = Pt(8)
    p_img.paragraph_format.space_after = Pt(4)
    p_img.paragraph_format.keep_with_next = True

    run_img = p_img.add_run()
    run_img.add_picture(image_path, width=Inches(width_inches))

    # Caption paragraph directly below
    p_cap = doc.add_paragraph()
    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap.paragraph_format.space_before = Pt(2)
    p_cap.paragraph_format.space_after = Pt(12)

    label_str = f"Figure {figure_number}: " if figure_number else "Figure: "
    r_lbl = p_cap.add_run(label_str)
    r_lbl.font.name = "Times New Roman"
    r_lbl.font.size = Pt(10.5)
    r_lbl.bold = True

    r_desc = p_cap.add_run(caption_text)
    r_desc.font.name = "Times New Roman"
    r_desc.font.size = Pt(10.5)
    r_desc.italic = True

    return p_img
