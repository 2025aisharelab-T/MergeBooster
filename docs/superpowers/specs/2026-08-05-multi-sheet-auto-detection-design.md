# 엑셀 전체 시트 자동 인식 및 다중 시트 취합 기능 설계서

## 1. 개요 (Overview)
엑셀 파일에 여러 개(다중)의 워크시트가 존재하는 경우, 사용자가 시트명을 일일이 지정하지 않아도 원본(Base) 엑셀 파일 내의 모든 시트를 자동으로 인식하고 수정본(Mod) 엑셀 파일의 동일 시트와 비교하여 변경사항을 순차 취합할 수 있도록 개선합니다.

## 2. 변경 요구사항 (Requirements)
1. **시트 자동 인식 기본화**:
   - UI의 비교 시트명 기본값을 `"(전체 시트 자동 인식)"`으로 변경합니다.
   - 빈 값 `""` 또는 `"(전체 시트 자동 인식)"`, `"(전체 시트)"`, `"(모든 시트)"`가 지정될 경우 전체 시트 순회 모드로 동작합니다.
2. **동일 시트 자동 매칭 및 스킵 처리**:
   - 원본 워크북의 모든 시트(`wbBase.Worksheets`)를 순회합니다.
   - 수정본 워크북(`wbMod`)에 동명의 시트가 존재할 경우, 두 시트 간 셀 단위 diff를 비교하여 하이라이트 및 주석을 누적 반영합니다.
   - 수정본 워크북에 해당 시트가 존재하지 않으면 해당 시트는 스킵(Skip)하고 로그 메시지를 출력합니다.
3. **하위 호환성 지원**:
   - 특정 시트명이 입력된 경우 해당 시트만 비교합니다.
   - `"(첫번째 시트)"`로 명시 입력 시 첫 번째 시트만 비교합니다.

## 3. 상세 변경 내역 (Detailed Changes)

### 3.1 C# Excel Engine (`src/cs/ExcelDiffEngine.cs`)
- `ProcessDiffMultiple` 내부의 시트 처리 구조 변경:
  - `bool processAllSheets`: `sheetName`이 빈값이거나 `(전체 시트 자동 인식)`, `(전체 시트)`, `(모든 시트)`인지 판단.
  - `bool isFirstSheetOnly`: `sheetName`이 `(첫번째 시트)`인지 판단.
  - 각 수정본 파일(`modPath`) 처리 시, `processAllSheets`가 true이면 `wbBase.Worksheets` 전체를 순회하며 `wbMod.Worksheets`에서 동명 시트를 탐색 후 diff 수행.

### 3.2 C# UI (`src/cs/MainForm.cs`)
- `txtSheetName.Text` 기본값을 `"(첫번째 시트)"`에서 `"(전체 시트 자동 인식)"`으로 변경.

### 3.3 VBA Engine (`src/vba/mod_DiffEngine.bas`)
- `ProcessDiffAndMerge` 내부 시트 매칭 로직을 `(전체 시트 자동 인식)` 조건 지원하도록 수정.

## 4. 검증 계획 (Verification Plan)
1. `scripts/build_exe.bat` 실행을 통해 `ExcelDiffAggregator.exe` 성공적인 재컴파일 확인.
2. 테스트 파이썬 스크립트(`tests/test_simulation.py`) 수정/확장: 다중 시트가 포함된 원본 및 수정본 샘플 엑셀 파일 생성 후 동작 테스트 진행.
