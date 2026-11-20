# VIT Pune Course Project Report Generator & Format Standard

A production-grade report generator and format standard designed to produce **100% compliant, publication-grade engineering course project reports** adhering strictly to the official guidelines of **Vishwakarma Institute of Technology (VIT Pune)**, Department of Computer Engineering / CSE (Data Science).

---

## 📑 Overview

Academic project reports often suffer from formatting degradations:
- Left margin clipping when bound (due to missing 1.25" gutter).
- Awkward paragraph line-stretching when applying justification in Microsoft Word.
- Ugly or unreadable rasterized math equations instead of native formulas.
- Tables splitting mid-row across page breaks with missing headers.
- Trailing 2-line pages, orphaned headings, and captions detached from figures.

This repository provides an **automated, space-aware engineering compilation pipeline** along with **reusable templates, custom AI skills, and reference outputs** so that students and AI agents can produce flawless academic reports for **any VIT course** in minutes.

---

## 🚀 Quick Start (In 3 Simple Steps)

### 1. Configure Project Metadata
Edit `config/report_config.json`:
- **Course Name:** e.g., *Soft Computing, Operating Systems, Machine Learning, Computer Networks, DBMS*
- **Project Title:** Title of your project
- **Group Number:** Group allocation number
- **Team Members:** Names, Roll Numbers, and PRNs
- **Faculty Guide:** Name, designation, and salutation (`Prof. Dr.`)
- **Institution:** Trust name, college name, department, academic year, and submission date

### 2. Add Chapter Content
Drop your project write-ups into `chapters/`:
- `01_introduction.md`
- `02_literature_review.md`
- `03_methodology.md`
- `04_system_architecture.md`
- `05_implementation.md`
- `06_results_and_discussion.md`
- `07_conclusion_and_future_scope.md`
- `08_references.md`

### 3. Compile Report
Run the master compiler script with Python 3.12:
```bash
python scripts/build_report.py
```
Output: Generates `Course_Project_Report.docx` and `Course_Project_Report.pdf`.

---

## 📐 Golden Format Specification (VIT Pune B.Tech Format)

| Element | Specification | Technical Enforcement |
| :--- | :--- | :--- |
| **Typography (Body)** | Times New Roman, 12pt, Regular | Unified across all paragraphs. |
| **Line Spacing** | 1.5 Lines, 6pt Space After, 0pt Before | Readable line height for faculty evaluation. |
| **Paragraph Alignment** | Full Justification (`JUSTIFY`) | Strips accidental hard breaks before justification to avoid word gaps. |
| **Margins** | **Left: 1.25"** (binding gutter)<br>**Right: 1.0"** \| **Top: 1.0"** \| **Bottom: 1.0"** | Prevents spiral/hardbound binding from clipping text. |
| **Heading 1 (Chapter)** | Times New Roman, 16pt, Bold, Centered or Left | Starts on a fresh page (`CHAPTER X: TITLE`). |
| **Heading 2 (Section)** | Times New Roman, 14pt, Bold, Left | E.g., `1.1 Problem Statement`. |
| **Heading 3 (Sub-section)** | Times New Roman, 12pt, Bold, Left | E.g., `1.1.1 Boundary Constraints`. |
| **Equations** | Native OMML Math, Centered | 2-column borderless table: formula centered, `(X.Y)` right-aligned. |
| **Figure Captions** | Times New Roman, 10–11pt, Centered, **Below Figure** | Bound to image using `keep_with_next = True`. |
| **Table Captions** | Times New Roman, 10–11pt, Centered, **Above Table** | Bound to table using `keep_with_next = True`. |
| **Table Layout** | Anti-Orphan Rows & Repeated Headers | `<w:cantSplit/>` on all rows; `<w:tblHeader/>` on Row 0. |

---

## 🧮 Mathematical Formulations (Native OMML Math)

Rather than embedding low-resolution raster images or raw LaTeX code, equations are rendered as **genuine Microsoft Office Math (`OMML`)** elements using `latex2mathml`:
- Rendered natively with smooth fraction bars, summation signs, integrals, and Greek glyphs.
- Each equation is presented in a borderless table with right-aligned numbering `(X.Y)`.
- Followed by a structured *'Where:'* symbol definition breakdown specifying units.

---

## 🏛️ Official Institutional Logo

Support for the official VIT Pune crest is built into the cover and certificate layouts:
- **VIT Pune Logo:** `assets/VIT.png`
- Configurable height (calibrated to 1.2" centered) without distorting page flow.

---

## 🔍 Agent Verification & Page Audit Process

A key innovation of this repository is the **automated page-by-page visual and textual audit** performed by the AI agent during report creation:
1. **Pass 1:** Generate initial Word document and export draft PDF via Word COM.
2. **Page Extraction:** Extract each physical PDF page as a high-resolution 144–300 DPI screenshot (`page.get_pixmap()`) using PyMuPDF (`fitz`).
3. **Automated Audit:** Verify that the Certificate Page fits completely on Page 2 with zero spillover, check title red highlights, verify student roll numbers, and detect any table or heading orphans.
4. **Pass 2 Synchronization:** Extract exact physical page numbers for all sections and inject them into the Table of Contents, List of Figures, and List of Tables.

For a full step-by-step breakdown of how the agent executes this, see [`docs/AGENT_WORKFLOW_AND_PAGE_AUDIT_GUIDE.md`](docs/AGENT_WORKFLOW_AND_PAGE_AUDIT_GUIDE.md).

---

## 📂 Repository Structure

```
├── README.md                              <-- Project guide and overview
├── SETUP_GUIDE.md                         <-- Quick setup instructions
├── requirements.txt                       <-- Python dependencies (python-docx, Pillow, etc.)
├── .gitignore                             <-- Clean ignore rules
│
├── config/
│   └── report_config.json                 <-- Project metadata configuration
│
├── assets/
│   └── VIT.png                            <-- Official VIT Pune crest
│
├── skills/
│   └── academic-course-project-report/
│       └── SKILL.md                       <-- Antigravity / AI Agent custom skill
│
├── docs/
│   ├── FORMAT_SPECIFICATION.md            <-- In-depth formatting rulebook
│   └── AGENT_WORKFLOW_AND_PAGE_AUDIT_GUIDE.md  <-- Agent page audit & verification guide
│
├── scripts/
│   ├── build_report.py                    <-- Master report compiler
│   ├── math_omml.py                       <-- LaTeX to native OMML converter
│   ├── table_styler.py                    <-- Defensive table formatting
│   ├── figure_manager.py                  <-- Figure scaling & caption binding
│   └── curated_all_scripts.py             <-- Complete 15-script reference compendium
│
├── chapters/                              <-- Modular chapter templates
│   ├── 01_introduction.md
│   ├── 02_literature_review.md
│   ├── 03_methodology.md
│   ├── 04_system_architecture.md
│   ├── 05_implementation.md
│   ├── 06_results_and_discussion.md
│   ├── 07_conclusion_and_future_scope.md
│   └── 08_references.md
│
└── reference/                             <-- Golden Standard Reference Documents
    ├── CP_Report_Format.docx              <-- Institutional raw format template
    ├── CP_Report_Format.pdf               <-- Institutional format PDF guide
    ├── Range_Intelligence_Course_Project_Report.docx  <-- 54-page reference Word doc
    └── Range_Intelligence_Course_Project_Report.pdf   <-- 54-page reference PDF
```

---

## 🤖 Using as an AI Agent Skill

Copy `skills/academic-course-project-report/SKILL.md` to your agent skills folder:
`~/.gemini/config/skills/academic-course-project-report/SKILL.md`

Any coding agent can then generate perfect reports automatically with a single prompt:
> *"Generate a course project report for my Database Systems project following the academic-course-project-report skill format."*
