# Academic Course Project Report Generator & Format Standard

A universal, production-grade report generator and format standard designed to produce **100% compliant, publication-grade engineering course project reports** adhering strictly to institutional requirements (including **Vishwakarma Institute of Technology (VIT Pune)**, **IITs**, **NITs**, and autonomous universities).

---

## 📑 Overview

Academic project reports often suffer from formatting degradations:
- Left margin clipping when bound (due to missing 1.25" gutter).
- Awkward paragraph line-stretching when applying justification in Microsoft Word.
- Ugly or unreadable rasterized math equations instead of native formulas.
- Tables splitting mid-row across page breaks with missing headers.
- Trailing 2-line pages, orphaned headings, and captions detached from figures.

This repository provides an **automated, space-aware engineering compilation pipeline** along with **reusable templates, custom AI skills, and reference outputs** so that students and AI agents can produce flawless academic reports for **any course** in minutes.

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

## 📐 Golden Format Specification

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

## 🏛️ Institutional Logo Support

Support for official college crests is built into the cover and certificate layouts:
- **VIT Pune Logo:** `assets/VIT.png`
- **IIT / University Crest:** Replace or configure in `assets/IIT_Logo_Template.png`
- Configurable height (typically 1.2"–1.4" centered) without distorting page flow.

---

## 📂 Repository Structure

```
├── README.md                              <-- Project guide and overview
├── config/
│   └── report_config.json                 <-- Project metadata configuration
├── skills/
│   └── academic-course-project-report/
│       └── SKILL.md                       <-- Antigravity / AI Agent custom skill
├── docs/
│   └── FORMAT_SPECIFICATION.md            <-- In-depth formatting rulebook
├── scripts/
│   ├── build_report.py                    <-- Master report compiler
│   ├── math_omml.py                       <-- LaTeX to native OMML converter
│   ├── table_styler.py                    <-- Defensive table formatting
│   └── figure_manager.py                  <-- Figure scaling & caption binding
├── chapters/                              <-- Modular chapter templates
│   ├── 01_introduction.md
│   ├── 02_literature_review.md
│   ├── 03_methodology.md
│   ├── 04_system_architecture.md
│   ├── 05_implementation.md
│   ├── 06_results_and_discussion.md
│   ├── 07_conclusion_and_future_scope.md
│   └── 08_references.md
├── assets/
│   ├── VIT.png                            <-- Official VIT crest
│   └── IIT_Logo_Template.png              <-- Generic / IIT crest template
└── reference/
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
