# 엑셀 변경사항 다중 취합기 (Excel Multi-Diff Aggregator GUI App) 설계서

- **작성일**: 2026-08-03
- **형식**: Windows 단일 독립 실행 파일 (`ExcelDiffAggregator.exe`)
- **기술 스택**: C# .NET Framework 4.0 / WinForms / Excel COM Interop (Dynamic Binding)
- **목적**: 여러 담당자가 개별적으로 수정한 엑셀 파일들을 다중 선택/체크하여 원본 문서에 순차적으로 자동 반영 및 노란색 하이라이트 적용.

---

## 1. UI 레이아웃 업데이트 명세

### 1.1 메인 윈도우 구성
- **타이틀**: `📊 엑셀 변경사항 다중 취합기 (Excel Multi-Diff Aggregator)`
- **크기**: `680 x 620` (고정 크기, 모던 Flat UI 스타일)

### 1.2 주요 컨트롤
| 영역 | 컨트롤 | 설명 |
| :--- | :--- | :--- |
| **원본 파일** | `txtBasePath` & `btnSelectBase` | 원본(기준) 엑셀 파일 지정 |
| **수정본 파일 목록** | `lstModFiles` (`CheckedListBox`) | 다중 첨부된 수정본 파일 목록 및 체크박스 선택기 |
| | `btnAddModFiles` | `[ ➕ 파일 추가... ]` (OpenFileDialog MultiSelect = true) |
| | `btnRemoveSelected` | `[ ➖ 선택 삭제 ]` |
| | `btnClearAll` | `[ 🧹 전체 비우기 ]` |
| **옵션** | `txtSheetName`, `chkHighlight`, `chkComment` | 시트명, 노란색 배경 강조, 주석 기록 |
| **실행 및 로그** | `btnRun`, `txtLog` | `[ 🚀 체크된 N개 파일 순차 취합 실행 ]` 및 실시간 로그 |

---

## 2. 핵심 로직 명세

1. **다중 파일 순차 비교 처리**:
   - `lstModFiles.CheckedItems`를 순회하며 체크된 파일 경로 리스트 취득.
   - 각 수정본 파일에 대하여 `ExcelDiffEngine.ProcessDiffMultiple` 호출.
2. **누적 업데이트 & 주석명 세부화**:
   - 수정본 A ➔ 원본 반영 ➔ 수정본 B ➔ 원본 누적 반영.
   - 이전 값 주석 기록 시 `[파일명: A.xlsx] 기존 값: XXX` 형식으로 출처 명시.
3. **C# 5 호환성 유지**:
   - `csc.exe` (.NET Framework 4.0) 호환 구문 사용.
