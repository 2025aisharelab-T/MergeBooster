# 메모/음영 제거 기능 및 프리미엄 UI 구현 계획 (Implementation Plan)

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 원본 엑셀 파일의 메모 및 음영 일괄 제거 기능 구현 및 C# WinForms UI를 화려하고 세련된 모던 다크/슬레이트 프리미엄 테마로 전면 재설계.

**Architecture:** C# COM Interop 음영/메모 제거 엔진 + WinForms STA Threading 및 Modern Flat Style UI Component.

**Tech Stack:** C# (.NET Framework 4.0 / WinForms), Native Windows `csc.exe` compiler.

## Global Constraints
- Target Binary: `ExcelDiffAggregator.exe`
- Compiler: `C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe`
- New Feature: `ResetHighlightsAndComments` Engine Method & `🧹 메모 및 음영 제거` UI Button

---

### Task 1: `ExcelDiffEngine.cs`에 메모 및 음영 제거 메서드 추가

**Files:**
- Modify: `src/cs/ExcelDiffEngine.cs`

- [ ] **Step 1: `ResetHighlightsAndComments` 메서드 구현**

`ExcelDiffEngine.cs`에 `ResetHighlightsAndComments` static 메서드 추가:
- 원본 엑셀 파일 열기.
- 모든 시트(또는 지정 시트)의 `UsedRange.Interior.ColorIndex = -4142` (xlNone) 지정.
- `UsedRange.ClearComments()` 실행.
- `wbBase.Save()` 후 닫기.

---

### Task 2: `MainForm.cs` 프리미엄 모던 UI 전면 재설계 및 이벤틀 연결

**Files:**
- Modify: `src/cs/MainForm.cs`

- [ ] **Step 1: UI 테마 및 레이아웃 모던화**

- 상단 헤더 패널 (Deep Slate Navy `#0F172A`, 화려한 그래픽 텍스트)
- 카드형 GroupBox 스타일링
- `🧹 메모 및 음영 제거` 전용 버튼 추가 (Purple Accent `#7C3AED`)
- 버튼 클릭 시 `BackgroundWorker`를 통해 `ResetHighlightsAndComments` 호출.
- 콘솔 테마 가독성 향상.

---

### Task 3: 바이너리 컴파일 및 최종 검증

**Files:**
- Execute: `scripts/build_exe.bat`
- Output: `ExcelDiffAggregator.exe`

- [ ] **Step 1: `scripts/build_exe.bat` 실행 및 빌드 확인**
