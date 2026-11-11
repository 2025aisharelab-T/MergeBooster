# 엑셀 변경사항 자동 취합기 (Excel Diff Aggregator GUI App) 설계서

- **작성일**: 2026-08-03
- **형식**: Windows 단일 독립 실행 파일 (`ExcelDiffAggregator.exe`)
- **기술 스택**: C# .NET Framework 4.0 / Windows Forms / Excel COM Interop (Dynamic Binding)
- **목적**: 필터 및 숨겨진 행(Hidden Rows)에 영향받지 않고, 원본 엑셀 문서 대비 수정본 엑셀 문서의 변경된 셀만 추출하여 원본에 자동 반영 및 연한 노란색 하이라이트 적용.

---

## 1. 사용자 인터페이스 (GUI Layout)

### 1.1 메인 윈도우 구성
- **타이틀**: `📊 엑셀 변경사항 자동 취합기 (Excel Diff Aggregator)`
- **크기**: `640 x 520` (고정 크기, 모던 Flat UI 스타일)

### 1.2 주요 컨트롤
| 위치 | 컨트롤 이름 | 유형 | 설명 |
| :--- | :--- | :--- | :--- |
| **상단** | `txtBasePath` | TextBox & Button | 원본(기준) 엑셀 파일 경로 선택 (`.xlsx`) |
| | `txtModPath` | TextBox & Button | 수정본 엑셀 파일 경로 선택 (`.xlsx`) |
| **중단** | `txtSheetName` | TextBox | 비교 대상 시트명 (기본값: `(첫번째 시트)`) |
| | `chkHighlight` | CheckBox | 노란색 하이라이트 적용 여부 (기본: Checked) |
| | `chkComment` | CheckBox | 이전 값 메모(주석) 추가 여부 (기본: Checked) |
| **하단** | `btnRun` | Button | `[ ⚡ 변경사항 취합 및 덮어쓰기 실행 ]` |
| | `lblStatus` & `txtLog` | Label & TextBox | 진행 상황 표시 및 결과 로그 출력 |

---

## 2. 핵심 처리 로직 (C# Architecture)

1. **배경 스레드 비동기 처리 (`BackgroundWorker`)**:
   - UI 멈춤(Freezing) 방지를 위해 취합 작업을 별도 배경 스레드에서 수행.
   - 작업 진행률 및 상태를 실시간으로 `txtLog`에 출력.
2. **Dynamic Excel COM Interop**:
   - `Type.GetTypeFromProgID("Excel.Application")`을 사용하여 Excel 어플리케이션 비가시화(`Visible = false`) 실행.
   - `ScreenUpdating = false`, `DisplayAlerts = false`로 빠른 속도 보장.
3. **2차원 메모리 배열 비교**:
   - `UsedRange.Value`를 `object[,]` 배열로 읽어 메모리 상에서 초고속 비교.
   - 필터나 숨겨진 행에 상관없이 절대 행/열 좌표 `(r, c)` 기준으로 1:1 비교.
4. **셀 업데이트 및 서식 지정**:
   - `vMod != "" && vMod != vBase` 조건 시 `Cells[r, c].Value = vMod`.
   - 하이라이트 색상: `Interior.Color = ColorTranslator.ToOle(Color.FromArgb(255, 242, 204))`.
   - 주석 추가 옵션: `AddComment("기존 값: " + vBase)`.
5. **리소스 해제 및 예외 처리**:
   - `Marshal.ReleaseComObject`로 Excel COM 프로세스 완전 종료 보장.

---

## 3. 빌드 및 배포 방법

- **컴파일러**: `C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe`
- **타겟**: `/target:winexe` `/r:System.Windows.Forms.dll` `/r:System.Drawing.dll` `/r:Microsoft.CSharp.dll`
- **배포물**: `ExcelDiffAggregator.exe` (단일 실행 파일, 외부 DLL 필요 없음)
