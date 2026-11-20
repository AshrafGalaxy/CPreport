---
name: academic-course-project-report
description: Generates publication-grade, fully compliant academic course project reports adhering strictly to institutional engineering guidelines (specifically modeled after the VIT Pune Course Project format). Handles front matter, certificate pages, native OMML math equations, zero-orphan tables, scaled-in figures, PyMuPDF page-by-page visual auditing, and accurate 2-pass Table of Contents page synchronization into both DOCX and PDF formats.
---

# Academic Course Project Report Generation Skill (VIT Pune Format)

This skill provides the comprehensive standard, structural rules, and automated compilation pipeline for generating engineering course project reports that adhere 100% to strict institutional guidelines (modeled after the **Vishwakarma Institute of Technology, Pune B.Tech Course Project Format**).

Use this skill whenever:
- The user requests a course project report, capstone report, semester project documentation, or academic submission for ANY course (e.g., Soft Computing, Machine Learning, Operating Systems, Computer Networks, Database Systems, Cloud Computing, etc.).
- The output must strictly follow formal formatting: Title page, Certificate, Project Details page, Acknowledgement, Abstract, Table of Contents, List of Figures, List of Tables, Chapter Hierarchy, IEEE References, and Appendices.

---

## 1. Golden Format Specifications (Strict Requirements)

All generated documents MUST adhere to these exact typographic and structural constraints:

| Element | Specification | Rationale / Rules |
| :--- | :--- | :--- |
| **Typography (Body)** | Times New Roman, 12pt, Regular | Standard academic requirement across all chapters. |
| **Line Spacing** | 1.5 Lines, 6pt Space After, 0pt Before | Mandatory for readability and faculty grading annotations. |
| **Paragraph Alignment** | Full Justification (`JUSTIFY`) | Straight left and right margins. All hard line breaks inside paragraphs must be stripped to prevent Word spacing stretch. |
| **Page Margins** | Left: **1.25"** (binding gutter)<br>Right: **1.0"**<br>Top: **1.0"**<br>Bottom: **1.0"** | Leaves sufficient gutter on the left side for spiral or hardbound binding without clipping text. |
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
│   ├── 1. Title / Cover Page (College crest, trust, title in bold red, degree, student names & PRNs/Rolls, guide, academic year)
│   ├── 2. Certificate Page (Formal bonafide statement, guide signature, HOD signature, place, date)
│   ├── 3. Project Details Page (Group number, member table with roll numbers, guide signature block)
│   ├── 4. Acknowledgement & Declaration
│   ├── 5. Executive Summary / Abstract
│   ├── 6. Table of Contents (With exact physical page numbers)
│   ├── 7. List of Figures (With exact page numbers)
│   ├── 8. List of Tables (With exact page numbers)
│   └── 9. List of Abbreviations (4-column clean table)
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

## 3. The 4-Stage Agent Execution & Verification Lifecycle

To achieve 100% compliance without formatting defects, the agent executes four distinct phases:

### Phase 1: Construction Engine (Python-docx + OMML)
- Builds document using `python-docx` with defensive OpenXML tags.
- Converts LaTeX to native Microsoft Office Math (`OMML`) XML using `latex2mathml` and `MML2OMML.XSL`.
- Enforces `<w:cantSplit/>` on all table rows and `<w:tblHeader/>` on row 0.
- Binds figure paragraphs to captions via `keep_with_next = True`.

### Phase 2: Native Word COM Export (True Pagination)
- Uses Microsoft Word automation (`win32com.client`) to open the `.docx` and export the initial `.pdf` via `doc.SaveAs(PDF_PATH, FileFormat=17)`.
- Guarantees identical pagination between Word and PDF.

### Phase 3: Page-by-Page Visual & Textual Audit (PyMuPDF)
- Opens the PDF using `fitz` (PyMuPDF) and audits:
  1. **Page 1 (Cover):** Verifies red title span (`0xC00000`/`0xFF0000`), VIT crest, student names and PRNs.
  2. **Page 2 (Certificate):** Confirms Guide & HOD signatures, Place, and Date fit on Page 2 with ZERO spillover to Page 3.
  3. **Page 3 (Project Details):** Verifies group number and roll numbers are sorted properly.
  4. **Screenshot Extraction:** Extracts high-resolution PNG images of every page (`page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0))`) into an audit folder to visually verify that there are zero orphaned headings, zero split tables, and zero running headers.

### Phase 4: Two-Pass Table of Contents Synchronization
- Scans the generated PDF to extract the exact physical page numbers for all headings, figures, and tables.
- Injects these physical page numbers into the Word document's Table of Contents, List of Figures, and List of Tables.
- Re-exports and verifies the final PDF.
