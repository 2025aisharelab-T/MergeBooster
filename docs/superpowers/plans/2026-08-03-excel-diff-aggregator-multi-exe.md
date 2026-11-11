# 엑셀 변경사항 다중 취합기 (Excel Multi-Diff Aggregator GUI App) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 다중 수정본 엑셀 파일 선택, CheckedListBox 체크 선택 및 순차 누적 취합 기능을 지원하는 C# WinForms 독립 실행 파일(`ExcelDiffAggregator.exe`) 빌드

**Architecture:** C# WinForms UI (`CheckedListBox` + Buttons) + BackgroundWorker 비동기 다중 파일 루프 + Excel COM Interop 엔진.

**Tech Stack:** C# (.NET Framework 4.0 / WinForms), Native Windows `csc.exe` compiler.

## Global Constraints
- Target Binary: `ExcelDiffAggregator.exe`
- Compiler: `C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe` (C# 5 syntax)
- Highlight Color: Light Yellow `#FFF2CC` (RGB: 255, 242, 204)
- Multi-File CheckedListBox selection support.

---

### Task 1: 다중 파일 지원 C# UI 및 엔진 소스 수정 (`ExcelDiffEngine.cs`, `MainForm.cs`)

**Files:**
- Modify: `src/cs/ExcelDiffEngine.cs`
- Modify: `src/cs/MainForm.cs`

- [ ] **Step 1: Update `src/cs/ExcelDiffEngine.cs`**
- Support multiple modified file paths array (`string[] modPaths`)
- Loop through each modified file, accumulating diffs on the base workbook
- Record modified file name in cell comment

- [ ] **Step 2: Update `src/cs/MainForm.cs`**
- Replace single modified file textbox with `CheckedListBox lstModFiles`
- Add file operation buttons (`➕ 파일 추가...`, `➖ 선택 삭제`, `🧹 전체 비우기`)
- Pass list of checked file paths to `BackgroundWorker`

- [ ] **Step 3: Commit Task 1**

```bash
git add src/cs/ExcelDiffEngine.cs src/cs/MainForm.cs
git commit -m "feat(cs): add multi-file CheckedListBox selection and sequential aggregation engine"
```

---

### Task 2: `ExcelDiffAggregator.exe` 컴파일 및 최종 검증

**Files:**
- Execute: `scripts/build_exe.bat`
- Build Output: `ExcelDiffAggregator.exe`

- [ ] **Step 1: Run `scripts/build_exe.bat`**

Run: `scripts\build_exe.bat`
Expected: Successful build of updated `ExcelDiffAggregator.exe`

- [ ] **Step 2: Commit Task 2**

```bash
git add ExcelDiffAggregator.exe
git commit -m "build: re-compile multi-file ExcelDiffAggregator.exe binary"
```
