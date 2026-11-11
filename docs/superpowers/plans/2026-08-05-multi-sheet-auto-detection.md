# 엑셀 다중 시트 자동 인식 및 취합 기능 구현 계획 (Implementation Plan)

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 엑셀 원본 파일 내의 다중 워크시트를 자동으로 탐색하여 수정본 엑셀 파일 내의 동일한 이름의 시트들과 순차적으로 비교/취합하는 기능 구현

**Architecture:** C# COM Interop 시트 순회 비교 루프 및 MainForm UI 기본값 업데이트, VBA `mod_DiffEngine.bas` 동일 시트 매칭 순회 로직 구현.

**Tech Stack:** C# (.NET Framework 4.0 / WinForms), Native Windows `csc.exe` compiler, VBA.

## Global Constraints
- Target Binary: `ExcelDiffAggregator.exe`
- Compiler: `C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe`
- Auto Detection Triggers: 빈 값 `""`, `"(전체 시트 자동 인식)"`, `"(전체 시트)"`, `"(모든 시트)"`
- Single First Sheet Trigger: `"(첫번째 시트)"`

---

### Task 1: C# 엔진 및 UI 다중 시트 순회 구현

**Files:**
- Modify: `src/cs/ExcelDiffEngine.cs`
- Modify: `src/cs/MainForm.cs`

- [ ] **Step 1: `src/cs/ExcelDiffEngine.cs`에 시트 순회 및 동명 시트 매칭 로직 추가**

`ExcelDiffEngine.cs`의 `ProcessDiffMultiple` 내부를 수정하여:
`sheetName`이 빈 문자열이거나 `"(전체 시트 자동 인식)"`, `"(전체 시트)"`, `"(모든 시트)"` 일 경우 원본 워크북의 모든 시트를 순회하며, 수정본 워크북에 동명 시트가 존재하는 경우 비교 처리.

```csharp
bool isAutoAllSheets = string.IsNullOrEmpty(sheetName) || 
                      sheetName == "(전체 시트 자동 인식)" || 
                      sheetName == "(전체 시트)" || 
                      sheetName == "(모든 시트)";
bool isFirstSheetOnly = sheetName == "(첫번째 시트)";
```

- [ ] **Step 2: `src/cs/MainForm.cs` 기본 시트명 텍스트 변경**

`txtSheetName.Text` 기본값을 `"(전체 시트 자동 인식)"`으로 변경.

---

### Task 2: VBA 엔지니어링 모듈 다중 시트 지원 구현

**Files:**
- Modify: `src/vba/mod_DiffEngine.bas`

- [ ] **Step 1: `src/vba/mod_DiffEngine.bas` 수정**

`ProcessDiffAndMerge`에 `SheetName`이 `(전체 시트 자동 인식)` 등의 자동 인식 패턴일 때 모든 시트를 순회하는 로직 추가.

---

### Task 3: 바이너리 재빌드 및 시뮤레이션 검증

**Files:**
- Modify: `tests/test_simulation.py`
- Execute: `scripts/build_exe.bat`
- Output: `ExcelDiffAggregator.exe`

- [ ] **Step 1: `scripts/build_exe.bat` 실행하여 `ExcelDiffAggregator.exe` 컴파일**

- [ ] **Step 2: `tests/test_simulation.py`에 다중 시트 생성 로직 추가 후 실행 검증**

```python
# test_simulation.py에 시트2("부서정보") 추가
```
