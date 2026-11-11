# 메모/음영 제거 기능 및 프리미엄 UI 디자인 설계서

## 1. 개요 (Overview)
본 설계는 사용자 요구사항에 따라 엑셀 취합 결과 문서의 노란색 배경 음영 및 셀 메모(주석)를 일괄 제거하는 기능을 추가하고, C# WinForms 애플리케이션(`MainForm.cs`)의 전반적인 UI/UX 디자인을 현대적이고 화려하며 프리미엄한 스타일로 개편합니다.

## 2. 주요 기능 및 디자인 사양 (Specifications)

### 2.1 메모 및 음영 제거 엔진 (`ExcelDiffEngine.cs`)
- `ResetHighlightsAndComments(string basePath, string sheetName, LogHandler logger)` 메서드 신설.
- 지정된 원본 파일의 모든 워크시트(또는 선택된 시트)에 대해 `UsedRange.Interior.ColorIndex = xlNone` 및 `UsedRange.ClearComments()` 적용 후 자동 저장.
- 저장 시 안전한 COM 레퍼런스 해제 및 예외 억제 처리.

### 2.2 프리미엄 모던 UI 디자인 개편 (`MainForm.cs`)
- **색상 팔레트 (Color Palette)**:
  - Header & Accent: Deep Navy/Slate (`#0F172A`, `#1E293B`) 및 Indigo-Blue (`#3B82F6`)
  - Primary Action Button (`🚀 취합 실행`): Gradient 감성의 Bright Indigo (`#2563EB`)
  - Reset Action Button (`🧹 메모&음영 제거`): Purple/Amber Accent (`#7C3AED` 또는 `#D97706`)
  - File Operation Buttons (`➕ 파일 추가`, `➖ 선택 삭제`, `🧹 비우기`): Emerald Green (`#059669`), Rose (`#E11D48`), Slate (`#475569`)
  - Log Console: Slate Black (`#090D16`) 배경에 Neon Emerald (`#34D399`) 콘솔 폰트 적용.
- **레이아웃 및 카드 UI**:
  - 기존 일반 GroupBox를 모던 플랫 카드로 스타일링.
  - 가시성이 극대화된 크기와 간격, 둥근 버튼 및 명확한 아이콘 표기.

## 3. 구현 계획 (Implementation Steps)
1. `src/cs/ExcelDiffEngine.cs`에 `ResetHighlightsAndComments` 메서드 추가.
2. `src/cs/MainForm.cs`에 프리미엄 테마 적용 및 메모/음영 제거 버튼 생성 및 이벤트를 바인딩.
3. `scripts/build_exe.bat`로 `ExcelDiffAggregator.exe` 컴파일 및 최종 동작 검증.
