# Pastel Highlight Options & Modified Value Comments Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a 5-color pastel shading palette selector and update cell comments to show modified values alongside previous values during Excel aggregation and selective reset.

**Architecture:** Update `ExcelDiffEngine.cs` to accept dynamic OLE highlight color integers, detect all 5 pastel colors during selective reset, and format comments with modified values. Update `MainForm.cs` with a pastel color ComboBox and updated checkbox label.

**Tech Stack:** C#, .NET Framework 4.8 / Windows Forms, Excel COM Interop.

## Global Constraints

- Pastel Yellow: `#FFF2CC` (OLE: 13431551)
- Pastel Mint: `#E2EFDA` (OLE: 14348258)
- Pastel Blue: `#DDEBF7` (OLE: 16247517)
- Pastel Pink: `#FCE4D6` (OLE: 14083324)
- Pastel Purple: `#E8E2F0` (OLE: 15786728)
- Checkbox copy: `변경(수정값) 셀 메모(주석) 표시`
- Comment format: `[{modFileName}] 변경(수정값): {vMod} (기존: {vBase})\n수정일시: yyyy-MM-dd HH:mm`

---

### Task 1: Update ExcelDiffEngine for Pastel Colors and New Comment Format

**Files:**
- Modify: `d:\project\test\src\cs\ExcelDiffEngine.cs`

**Interfaces:**
- Consumes: `ProcessDiffMultiple(string basePath, string[] modPaths, string sheetName, bool applyHighlight, bool addComment, int highlightColorOle, LogHandler logger)`
- Consumes: `ResetHighlightsAndComments(string basePath, string sheetName, LogHandler logger)`

- [ ] **Step 1: Update ProcessDiffMultiple signature and color application**

In `ExcelDiffEngine.cs`:
Add `int highlightColorOle` parameter to `ProcessDiffMultiple` and `ProcessDiffMultipleInternal`.
In `ProcessDiffMultipleInternal`:
Use `highlightColorOle` for `cell.Interior.Color = highlightColorOle;`.
Update comment creation logic:
```csharp
string commentText = string.Format("[{0}] 변경(수정값): {1} (기존: {2})\n수정일시: {3}",
    modFileName,
    vMod,
    vBase,
    DateTime.Now.ToString("yyyy-MM-dd HH:mm"));
```

- [ ] **Step 2: Update ResetHighlightsAndComments for 5 pastel colors and new comment text**

In `ResetHighlightsAndCommentsInternal`:
Create an array of pastel colors:
```csharp
HashSet<long> pastelColors = new HashSet<long> {
    13431551L, // #FFF2CC Yellow
    14348258L, // #E2EFDA Mint
    16247517L, // #DDEBF7 Blue
    14083324L, // #FCE4D6 Pink
    15786728L  // #E8E2F0 Purple
};
```
In cell loop:
Check `if (pastelColors.Contains(colorVal))` -> reset `interior.ColorIndex = -4142`.

In comment loop:
Check `if (cText != null && (cText.Contains("변경(수정값):") || cText.Contains("기존 값:")))` -> delete comment.

- [ ] **Step 3: Verify ExcelDiffEngine compilation**

Run: `cmd /c scripts\build_exe.bat`
Expected: Build successful!

---

### Task 2: Update MainForm UI for Color Selector and Checkbox Label

**Files:**
- Modify: `d:\project\test\src\cs\MainForm.cs`

- [ ] **Step 1: Add Pastel Color ComboBox and update checkbox label**

In `MainForm.cs`:
1. Add `ComboBox cmbHighlightColor` to `grpOpts`.
2. Populate `cmbHighlightColor` with items:
   - `파스텔 노랑 (#FFF2CC)`
   - `파스텔 민트 (#E2EFDA)`
   - `파스텔 블루 (#DDEBF7)`
   - `파스텔 핑크 (#FCE4D6)`
   - `파스텔 퍼플 (#E8E2F0)`
3. Set `cmbHighlightColor.SelectedIndex = 0`.
4. Update `chkComment.Text = "변경(수정값) 셀 메모(주석) 표시";`.
5. Helper method `GetSelectedPastelColorOle()` to map selected index to OLE color code.

- [ ] **Step 2: Pass selected color to worker and ProcessDiffMultiple**

In `BtnRun_Click` and `Worker_DoWork`:
Pass `highlightColorOle` to `ExcelDiffEngine.ProcessDiffMultiple(...)`.

- [ ] **Step 3: Compile and Test**

Run: `cmd /c scripts\build_exe.bat`
Expected: Build successful! Created ExcelDiffAggregator.exe.
