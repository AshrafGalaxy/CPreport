---
name: academic-course-project-report
description: Generates publication-grade, fully compliant academic course project reports adhering strictly to institutional engineering guidelines (such as the VIT Pune Course Project format). Handles front matter, certificate pages, native OMML math equations, zero-orphan tables, scaled-in figures, and accurate 2-pass Table of Contents page synchronization into both DOCX and PDF formats.
---

# Academic Course Project Report Generation Skill

This skill provides the comprehensive standard, structural rules, and automated compilation pipeline for generating engineering course project reports that adhere 100% to strict institutional guidelines (specifically modeled after the Vishwakarma Institute of Technology, Pune B.Tech Course Project Format).

Use this skill whenever:
- The user requests a course project report, capstone report, semester project documentation, or academic submission for ANY engineering course (e.g., Soft Computing, Machine Learning, Operating Systems, Computer Networks, Database Systems, Cloud Computing, etc.).
- The output must strictly follow formal formatting: Title page, Certificate, Project Details page, Acknowledgement, Abstract, Table of Contents, List of Figures, List of Tables, Chapter Hierarchy, IEEE References, and Appendices.

---

## 1. Golden Format Specifications (Strict Requirements)

All generated documents MUST adhere to these exact typographic and structural constraints:

| Element | Specification | Rationale / Rules |
| :--- | :--- | :--- |
| **Typography (Body)** | Times New Roman, 12pt, Regular | Standard academic requirement across all chapters. |
| **Line Spacing** | 1.5 Lines, 6pt Space After, 0pt Before | Mandatory for readability and faculty grading annotations. |
| **Paragraph Alignment** | Full Justification (`JUSTIFY`) | Straight left and right margins. All hard line breaks inside paragraphs must be stripped to prevent Word spacing stretch. |
| **Page Margins** | Left: **1.25"** (or 1.5" for binding)<br>Right: **1.0"**<br>Top: **1.0"**<br>Bottom: **1.0"** | Leaves sufficient gutter on the left side for spiral or hardbound binding without clipping text. |
| **Heading 1 (Chapter)** | Times New Roman, 16pt, Bold, Centered/Left | E.g., `CHAPTER 1: INTRODUCTION` (All caps, starts on a fresh page). |
| **Heading 2 (Section)** | Times New Roman, 14pt, Bold, Left | E.g., `1.1 Problem Motivation` |
| **Heading 3 (Sub-section)** | Times New Roman, 12pt, Bold, Left | E.g., `1.1.1 Boundary Constraints` |
| **Figure Captions** | Times New Roman, 10–11pt, Centered, **Below Figure** | Format: `Figure X.Y: Description`. Sized to 6.8"–7.1" printable width. Must have `keep_with_next` on image paragraph. |
| **Table Captions** | Times New Roman, 10–11pt, Centered, **Above Table** | Format: `Table X.Y: Description`. Must have `keep_with_next` on caption. |
| **Equations** | Native OMML Math, Centered | Enclosed in a 2-column borderless grid: equation centered on left, equation number `(X.Y)` right-aligned. Accompanied by a "Where:" symbol definition list. |
| **Table Layout** | `<w:cantSplit/>` on all rows<br>`<w:tblHeader/>` on Row 0 | Prevents single-row page splits across page breaks. Repeats header row automatically if a table splits. |
| **Headers & Footers** | Zero running headers if template prohibits | Clean pages without intrusive top headers unless requested. Page numbers placed in footer. |

---

## 2. Standard Document Hierarchy

Every generated report MUST follow this exact sequential structure:

```
├── FRONT MATTER (Roman numerals: i, ii, iii... Cover page unnumbered)
│   ├── 1. Title / Cover Page (College crest, institute trust, title, degree, student names & PRNs/Rolls, guide, academic year)
│   ├── 2. Certificate Page (Formal bonafide statement, guide signature, HOD signature, place, date)
│   ├── 3. Project Details Page (Group number, member table with roll numbers, guide signature block)
│   ├── 4. Acknowledgement & Declaration
│   ├── 5. Executive Summary / Abstract
│   ├── 6. Table of Contents (With exact physical page numbers)
│   ├── 7. List of Figures (With exact page numbers)
│   ├── 8. List of Tables (With exact page numbers)
│   └── 9. List of Abbreviations (3-column clean table)
│
├── BODY CHAPTERS (Arabic numerals: 1, 2, 3...)
│   ├── Chapter 1: Introduction (Background, Motivation, Problem Statement, Objectives, Scope)
│   ├── Chapter 2: Literature Review / Related Work (Comparative matrix, research gaps)
│   ├── Chapter 3: System Methodology & Theoretical Formulation (Mathematical formulations, algorithms)
│   ├── Chapter 4: System Architecture & Design (Block diagrams, module interactions, database schema)
│   ├── Chapter 5: Implementation Details & Verification (Algorithms, libraries, test cases)
│   ├── Chapter 6: Results, Performance Evaluation & Discussion (Metrics, plots, benchmarks, ablation)
│   └── Chapter 7: Conclusion & Future Scope (Summary of contributions, limitations, future work)
│
└── BACK MATTER
    ├── References (IEEE Citation Format: [1], [2], numbered sequentially as cited in text)
    └── Appendices (Mathematical proofs, full rule bases, test scripts, telemetry samples)
```

---

## 3. The 5 Core Engineering Implementations (Why It Works)

### Pillar 1: Native OMML Math Equations
Do NOT paste raw LaTeX code or low-resolution raster images into Word. Convert LaTeX to MathML using `latex2mathml`, then transform MathML to native Microsoft Office Math (`OMML`) via `MML2OMML.XSL` or insert Office Open XML math runs (`<m:oMathPara>`).
Place each equation inside a 1-row, 2-column borderless table:
- Left cell (85% width): Centered equation.
- Right cell (15% width): Right-aligned equation number `(3.1)`.

### Pillar 2: Zero Table Orphans
Tables that break awkwardly across pages look amateurish. Enforce two critical XML attributes via `python-docx`:
```python
# Prevent row from splitting in half across a page boundary
for row in table.rows:
    trPr = row._tr.get_or_add_trPr()
    trPr.append(parse_xml(r'<w:cantSplit xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'))

# Repeat header on subsequent pages if table overflows
header_trPr = table.rows[0]._tr.get_or_add_trPr()
header_trPr.append(parse_xml(r'<w:tblHeader xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'))
```

### Pillar 3: Defensive Caption and Image Binding
To prevent a figure from appearing at the bottom of a page with its caption orphaned at the top of the next page:
```python
p_img = doc.add_paragraph()
p_img.paragraph_format.keep_with_next = True
p_img.add_run().add_picture(img_path, width=Inches(6.8))

p_cap = doc.add_paragraph('Figure 4.1: Architecture Data Flow', style='Caption')
p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
```

### Pillar 4: Whitespace Normalization Before Justification
When using `WD_ALIGN_PARAGRAPH.JUSTIFY` in Word, any embedded `\n` or double spaces within a paragraph causes Word's justification engine to stretch remaining words across the entire line. Always normalize paragraphs:
```python
clean_text = ' '.join(raw_text.split())
p = doc.add_paragraph(clean_text)
p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
```

### Pillar 5: Two-Pass Accurate Page Number Synchronization
Never hardcode Table of Contents page numbers.
- **Pass 1:** Generate the initial `.docx` and convert to `.pdf` via Word COM or PyMuPDF.
- **Scan:** Use `pypdf` or `pymupdf` (fitz) to search for Heading strings and extract their actual physical page numbers.
- **Pass 2:** Inject the exact physical page numbers into the TOC, List of Figures, and List of Tables before finalizing the document.

---

## 4. Reusable Compilation Workflow

When building a report for any course:
1. Define metadata in a clean dictionary or `report_config.json`:
   - `course_name`, `project_title`, `group_no`, `team_members` (names, PRNs, roll numbers).
   - `guide_name`, `guide_salutation` ("Prof. Dr."), `hod_name`, `academic_year`, `submission_date`.
2. Author chapters in modular Markdown files under `chapters/`.
3. Generate or save high-resolution figures under `figures/` (at least 300 DPI).
4. Run the master compilation script with Python 3.12 (`py -3.12 build_report.py`).
5. Verify output using the PDF page audit checklist.
