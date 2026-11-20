# Academic Course Project Report Format Specification

This document defines the strict typographic, architectural, and layout guidelines for undergraduate and postgraduate engineering course project reports. Adherence to these guidelines ensures full compliance with institutional standards of **Vishwakarma Institute of Technology (VIT Pune)**.

---

## 1. Page Setup & Document Layout

### 1.1 Margins (Binding Gutter Aware)
- **Left Margin:** `1.25 inches` (31.75 mm) — Essential binding gutter allowance.
- **Right Margin:** `1.0 inch` (25.4 mm)
- **Top Margin:** `1.0 inch` (25.4 mm)
- **Bottom Margin:** `1.0 inch` (25.4 mm)
- **Header Distance:** `0.5 inches` (12.7 mm)
- **Footer Distance:** `0.5 inches` (12.7 mm)

### 1.2 Headers and Footers
- **Front Matter Pages:** Unnumbered cover page; lower-case Roman numerals (`i`, `ii`, `iii`...) for subsequent front matter.
- **Chapter Pages:** Arabic numerals (`1`, `2`, `3`...) centered or right-aligned in the footer.
- **Running Headers:** Suppressed unless explicitly mandated by the department.

---

## 2. Typography & Text Hierarchy

| Level | Element | Typeface | Size | Weight | Alignment | Spacing |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| **L0** | Report Title (Cover) | Times New Roman | 16–18 pt | Bold | Centered | 18 pt after |
| **L1** | Chapter Title | Times New Roman | 16 pt | Bold | Left/Center | 12 pt before, 14 pt after |
| **L2** | Main Section (e.g. 1.1) | Times New Roman | 14 pt | Bold | Left | 10 pt before, 6 pt after |
| **L3** | Sub-Section (e.g. 1.1.1) | Times New Roman | 12 pt | Bold | Left | 6 pt before, 4 pt after |
| **Body** | Regular Paragraph | Times New Roman | 12 pt | Regular | Justified | 0 pt before, 6 pt after, 1.5 line spacing |
| **Cap** | Figure/Table Caption | Times New Roman | 10–11 pt | Regular/Bold | Centered | 4 pt before, 6 pt after |
| **Foot** | Table Cell / Footnote | Times New Roman | 9–10 pt | Regular | Left/Center | 2 pt padding |

### 2.1 Paragraph Justification Rules
To avoid Word's notorious "stretched words" defect when using full justification (`WD_ALIGN_PARAGRAPH.JUSTIFY`):
1. **Never use hard line breaks (`\n` or `<w:br/>`)** inside body paragraphs.
2. Clean all text strings by replacing internal whitespace with a single space: `' '.join(text.split())`.
3. Set line spacing strictly to **1.5 lines** (multiplied or exact 18pt).

---

## 3. Mathematical Equations (Native OMML Math)

### 3.1 Strict Anti-Rasterization Rule
Never paste equations as low-resolution images or raw LaTeX text strings. Convert all equations into native Microsoft Office Math (`OMML`) XML blocks.

### 3.2 Two-Column Borderless Grid Alignment
Each displayed equation must be housed inside a 1-row, 2-column borderless table:
- **Left Column (85% width):** Centered equation in native math typography.
- **Right Column (15% width):** Right-aligned equation number `(X.Y)`.
- Follow each equation with a structured *'Where:'* symbol glossary specifying definitions and units.

---

## 4. Tables: Defensive Formatting & Anti-Orphan Rules

To guarantee professional table rendering across page breaks:
1. **Row Protection:** Apply `<w:cantSplit/>` to all rows. Rows will never split horizontally across a page break.
2. **Repeated Header:** Apply `<w:tblHeader/>` to Row 0 so column titles cleanly appear at the top of subsequent pages.
3. **Caption Binding:** Place table caption **ABOVE** the table and set `keep_with_next = True`.
4. **Padding & Shading:** Header rows shaded light gray (`#E2E8F0` or `#F1F5F9`). Internal cell padding: Top/Bottom 100 dxa (5pt), Left/Right 150 dxa (7.5pt).

---

## 5. Figures & Screenshots: Sizing and Binding

1. **Resolution:** All diagrams, GIS maps, and screenshots must be at least **300 DPI**.
2. **Width Scaling:** Scale large architectural diagrams to full printable width (`6.80"` to `7.10"`).
3. **Caption Binding:** Always set `keep_with_next = True` on the image paragraph and place the caption **BELOW** the figure.

---

## 6. Document Structure Sequence

```
1. Cover / Title Page (Unnumbered)
2. Certificate Page
3. Project Details Page
4. Declaration & Acknowledgements
5. Executive Summary / Abstract
6. Table of Contents
7. List of Figures
8. List of Tables
9. List of Abbreviations
10. Chapters 1 to 7
11. References (IEEE Style [1], [2]...)
12. Appendices
```
