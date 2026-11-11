# 엑셀 변경사항 자동 취합기 사용자 가이드 및 PDF 생성 설계서

- **작성일**: 2026-08-05
- **목적**: `ExcelDiffAggregator` 프로그램의 사용자 편의성을 극대화하기 위해 1장짜리 퀵 스타트 가이드와 상세 표준 사용 매뉴얼을 작성하고, 고품질 PDF 파일로 자동 변환하여 배포.

---

## 1. 개요 및 생성 대상 문서

사용자의 요구에 따라 다음 2가지 형태의 PDF 이용 가이드를 생성합니다.

| 문서 구분 | 파일명 | 페이지 분량 | 목적 및 대상 |
| :--- | :--- | :--- | :--- |
| **퀵 스타트 가이드** | `ExcelDiffAggregator_QuickGuide.pdf` | A4 1페이지 고정 | 빠른 사용법 숙지가 필요한 실무자용 한 눈에 보는 가이드 |
| **표준 사용 매뉴얼** | `ExcelDiffAggregator_UserManual.pdf` | A4 다중 페이지 (3~4P) | 기능 상세 설명, 설치/실행법, FAQ 및 예외 처리가 포함된 종합 가이드 |

---

## 2. 디자인 사양 및 스타일 규칙

`design.md`에 정의된 Visual Design System (Soft & Rounded Modern UI)을 그대로 적용합니다.

- **컬러 팔레트**:
  - Primary Accent: `#6366F1` (Soft Indigo) / `#2563EB` (Bright Indigo)
  - Dark Header & Slate: `#0F172A` (Navy Slate) / `#1E293B`
  - Card Surface: `#FFFFFF` (Pure White)
  - Background: `#F8FAFC` (Light Slate)
  - Text Muted: `#64748B`
  - Border: `#E2E8F0`
- **typography**: Modern Sans-Serif (`Segoe UI`, `Malgun Gothic`, `Noto Sans KR`, sans-serif)
- **레이아웃**:
  - `@media print` 쿼리를 적용하여 인쇄 시 헤더/푸터 깔끔 처리 및 여백 제어.
  - 퀵가이드: `page-break-inside: avoid;`로 1페이지 완벽 수록.
  - 표준매뉴얼: 섹션별 `page-break-after: always;` 또는 자연스러운 단락 구분 적용.

---

## 3. 문서별 세부 목차 및 구성

### 3.1 퀵 스타트 가이드 (`ExcelDiffAggregator_QuickGuide.pdf`)
1. **프로그램 헤더**: 타이틀, 서브타이틀 및 핵심 가치 요약 ("필터/숨겨진 행 무관 1:1 자동 취합기")
2. **3단계 초간단 사용법 (Visual Flow)**:
   - `Step 1`: 원본(Base) 엑셀 파일 및 수정본(Mod) 엑셀 파일들 등록
   - `Step 2`: 비교 시트명 확인 (기본값: `(전체 시트 자동 인식)`) 및 주석/음영 옵션 선택
   - `Step 3`: `🚀 취합 실행` 버튼 클릭 후 결과 확인
3. **핵심 기능 4선 요약 카드**:
   - 다중 수정본 일괄 병합
   - 전체 시트 자동 매칭
   - 절대 좌표 1:1 비교
   - `🧹 메모 & 음영 제거` 원클릭 초기화
4. **버튼 & 기능 퀵 참고표**

### 3.2 표준 사용 매뉴얼 (`ExcelDiffAggregator_UserManual.pdf`)
1. **제1장: 프로그램 소개 및 주요 기능**
   - 개발 배경 및 기존 복사/붙여넣기 방식과의 차이점 (필터/숨겨진 행 관련 오류 완벽 해결)
   - 주요 제공 기능 상세 목록
2. **제2장: 시스템 요구사항 및 실행 환경**
   - 지원 OS (Windows 7/10/11) 및 엑셀 버전 (Excel 2010 이상)
   - 독립 실행 파일 (`ExcelDiffAggregator.exe`) 사용법
3. **제3장: 화면 구성 및 기능 상세 설명**
   - 메인 GUI 윈도우 각 영역별 명칭 및 역할 (원본 파일 선택, 수정본 파일 목록, 옵션 체크박스, 콘솔 로그)
4. **제4장: 상세 활용 가이드**
   - 단일 시트 vs 다중 시트 자동 인식 동작 방식
   - 변경사항 셀 하이라이트(노란색 음영) 및 이전 값 메모(주석) 확인법
   - 병합 결과물 저장 규칙 (원본 파일 자동 덮어쓰기/저장)
   - 취합 후 서식 복원 (`메모 & 음영 제거` 기능)
5. **제5장: 자주 묻는 질문 (FAQ) 및 예외 조치**
   - 시트명 불일치 시 동작
   - 엑셀 파일 열림 상태 및 COM 프로세스 처리
   - 수식 셀 및 서식 비교 범위 안내

---

## 4. PDF 자동 생성 생성 파이프라인

1. **HTML/CSS 소스 문서 작성**:
   - `docs/manuals/quick_guide.html`
   - `docs/manuals/user_manual.html`
2. **MS Edge Headless 빌드 스크립트 작성**:
   - `scripts/generate_manual_pdfs.bat` 실행 시 Edge Headless CLI 호출:
     - `msedge --headless --print-to-pdf=ExcelDiffAggregator_QuickGuide.pdf --no-pdf-header-footer docs/manuals/quick_guide.html`
     - `msedge --headless --print-to-pdf=ExcelDiffAggregator_UserManual.pdf --no-pdf-header-footer docs/manuals/user_manual.html`
3. **결과물 검증 및 루트 배포**:
   - 루트 디렉터리에 `ExcelDiffAggregator_QuickGuide.pdf` 및 `ExcelDiffAggregator_UserManual.pdf` 최종 생성 확인.

---

## 5. 검증 계획 (Verification Plan)

1. `scripts/generate_manual_pdfs.bat` 실행을 통한 PDF 자동 생성 검증.
2. 생성된 PDF 파일 존재 여부 및 파일 크기 확인.
3. 퀵가이드가 정확히 1페이지 내에 딱 맞게 출력되었는지 레이아웃 검증.
