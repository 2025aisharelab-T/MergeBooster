# Pastel Highlight Options & Modified Value Comments Spec

## Summary
Add support for selecting 5 soft pastel shading colors for aggregated excel cells, and update cell comment format to explicitly show the modified (new) value along with the previous value.

## 1. Pastel Shading Color Palette
Users can select one of 5 pastel colors from the UI dropdown:

| Option Name | Hex Code | RGB Values | OLE Color Value |
| :--- | :--- | :--- | :--- |
| 파스텔 노랑 (기본) | `#FFF2CC` | (255, 242, 204) | 13431551 |
| 파스텔 민트 | `#E2EFDA` | (226, 239, 218) | 14348258 |
| 파스텔 블루 | `#DDEBF7` | (221, 235, 247) | 16247517 |
| 파스텔 핑크 | `#FCE4D6` | (252, 228, 214) | 14083324 |
| 파스텔 퍼플 | `#E8E2F0` | (232, 226, 240) | 15786728 |

### UI Design (MainForm.cs)
- Replace static label/checkbox text with a color selection ComboBox.
- Place `cmbHighlightColor` next to `chkHighlight`.
- Default selection: `파스텔 노랑 (#FFF2CC)`.

### Engine Integration (ExcelDiffEngine.cs)
- `ProcessDiffMultiple` takes `int highlightColorOle` parameter instead of hardcoded yellow.
- `ResetHighlightsAndComments` detects cells matching ANY of the 5 pastel colors so selective reset removes any pastel shading applied during past aggregations.

## 2. Modified Value Cell Comments
### UI Checkbox
- Update checkbox label in `MainForm.cs` from `이전 값 셀 메모(주석) 남기기` to `변경(수정값) 셀 메모(주석) 표시`.

### Comment Text Format
When a cell value is modified during aggregation, the comment text will be formatted as:
```text
[{modFileName}] 변경(수정값): {vMod} (기존: {vBase})
수정일시: yyyy-MM-dd HH:mm
```

### Selective Reset Integration
- `ResetHighlightsAndComments` removes comments containing `"변경(수정값):"` or `"기존 값:"` to support both new and legacy aggregated comments without affecting user custom comments.

## Verification Plan
1. Compile executable using `scripts/build_exe.bat`.
2. Run test verifying pastel color selection, comment text formatting, and selective reset behavior.
