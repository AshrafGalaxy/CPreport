# Setup & Workflow Automation Guide

## Prerequisites
- Python 3.10, 3.11, or 3.12
- Dependencies: `pip install -r requirements.txt`

## 1. Quick Local Execution
To compile a report locally from your terminal:
```bash
python scripts/build_report.py
```
This script reads `config/report_config.json`, reads all chapters in `chapters/`, and generates:
- `Course_Project_Report.docx`

## 2. Using with Antigravity / Coding Agents
Place the skill file:
`skills/academic-course-project-report/SKILL.md`
into your Antigravity skills directory (`~/.gemini/config/skills/academic-course-project-report/SKILL.md`).

Whenever you ask the agent to create or format a report for any course, it will automatically load this skill and format everything to this exact institutional standard.
