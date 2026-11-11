# Excel Diff Aggregator User Guide PDF Generation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build HTML/CSS templates and execution script to generate `ExcelDiffAggregator_QuickGuide.pdf` (1-page) and `ExcelDiffAggregator_UserManual.pdf` (multi-page) for ExcelDiffAggregator.

**Architecture:** Use Modern HTML5 & CSS3 with print media directives (`@media print`, `@page`) styled according to `design.md` (Soft Modern UI theme), and execute MS Edge Headless CLI (`msedge --headless --print-to-pdf=... --no-pdf-header-footer`) via Windows Batch script to export pristine PDFs.

**Tech Stack:** HTML5, CSS3 (Print CSS), Windows Batch Scripting, MS Edge Headless PDF Exporter.

## Global Constraints

- Quick Guide MUST fit into exactly 1 A4 page when printed.
- User Manual MUST be multi-page with clear section headers, tables, callouts, and clean page breaks.
- Color palette MUST match `design.md` (Primary `#6366F1` / `#2563EB`, Header Slate `#0F172A`, Surface White `#FFFFFF`, Soft Borders `#E2E8F0`).

---

### Task 1: Create HTML/CSS Template for Quick Start Guide

**Files:**
- Create: `docs/manuals/quick_guide.html`

- [ ] **Step 1: Write HTML markup and print CSS for Quick Start Guide**

Write `docs/manuals/quick_guide.html` with responsive layout, 1-page print CSS constraints, 3-step usage diagram, 4-feature card grid, and button reference table.

---

### Task 2: Create HTML/CSS Template for Standard User Manual

**Files:**
- Create: `docs/manuals/user_manual.html`

- [ ] **Step 1: Write HTML markup and print CSS for Standard User Manual**

Write `docs/manuals/user_manual.html` with multi-page structure, TOC, Chapters 1 through 5, feature details, FAQ, and exception handling notes.

---

### Task 3: Create and Run PDF Generation Batch Script

**Files:**
- Create: `scripts/generate_manual_pdfs.bat`
- Output: `ExcelDiffAggregator_QuickGuide.pdf`
- Output: `ExcelDiffAggregator_UserManual.pdf`

- [ ] **Step 1: Create batch script for Edge Headless PDF printing**

Write `scripts/generate_manual_pdfs.bat` calling `msedge.exe` with `--headless`, `--print-to-pdf`, and `--no-pdf-header-footer` flags for both HTML files.

- [ ] **Step 2: Execute batch script to generate PDFs**

Run `scripts/generate_manual_pdfs.bat` in working directory `d:\project\test`.

- [ ] **Step 3: Verify output PDF files existence and sizes**

Confirm that `ExcelDiffAggregator_QuickGuide.pdf` and `ExcelDiffAggregator_UserManual.pdf` are generated successfully.
