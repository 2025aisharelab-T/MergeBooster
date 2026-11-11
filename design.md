---
name: Modern Rounded Soft UI
colors:
  primary: "#6366F1"      # Soft Indigo
  primary-hover: "#4F46E5"
  secondary: "#EC4899"    # Soft Pink/Rose Accent
  background: "#F8FAFC"   # Light Slate Background
  surface: "#FFFFFF"      # White Card Surface
  text-main: "#0F172A"    # Deep Navy Black
  text-muted: "#64748B"   # Muted Gray Text
  border: "#E2E8F0"       # Light Subtle Border
rounded:
  sm: "8px"
  md: "12px"
  lg: "20px"
  xl: "28px"
  full: "9999px"
shadows:
  sm: "0 1px 3px rgba(0,0,0,0.05)"
  md: "0 4px 12px rgba(0,0,0,0.05)"
  lg: "0 10px 25px -5px rgba(0,0,0,0.08)"
---

# Visual Design System: Soft & Rounded Modern UI

## 1. Visual Theme & Philosophy
- **Concept:** Friendly, Modern, and Soft (둥글둥글하고 친근한 모던 UI).
- **Core Rule:** 각진 직사각형 요소(곡률 0px)를 전면 배제하고, 시각적 피로도를 줄이기 위해 넉넉한 여백(Padding)과 부드러운 라운딩, 은은한 그림자(Soft Elevation)를 적용합니다.

## 2. Component Style Rules

### Buttons
- **Primary Button:** `border-radius: 9999px` (알약 스타일) 또는 최소 `16px`.
- **Secondary / Ghost Button:** `border-radius: 12px`, 배경색은 은은한 패스텔 톤이나 투명 배경에 두꺼운 라운드 보더 적용.
- **Hover/Active Effect:** Click/Hover 시 미세하게 확대되는 Scale 효과(`scale(1.02)`) 및 은은한 Shadow 증가.

### Cards & Containers
- **Corner Radius:** `border-radius: 20px` ~ `24px`.
- **Border & Shadow:** 강한 검은색 테두리 대신 `1px solid #E2E8F0`와 `shadow-md`를 결합하여 입체감 형성.
- **Padding:** 컴포넌트 내부 여백을 넉넉히 주어 답답함을 해소 (최소 `24px` 이상 권장).

### Input & Form Fields
- **Corner Radius:** `border-radius: 12px` ~ `16px`.
- **Focus State:** Focus 시 테두리 색상이 Primary Color로 바뀌면서 `ring` 형태의 둥근 외곽선 강조.

### Badges, Chips & Tags
- **Corner Radius:** `border-radius: 9999px` (Complete Pill shape).
- **Typography:** `font-weight: 600`, 소문자/대문자 조합과 보스턴 스타일 패딩 적용.

## 3. Do's and Don'ts
- **DO:**
  - 모든 대화상자(Modal), 카드(Card), 버튼(Button)에 최소 `12px` 이상의 곡률(`border-radius`)을 유지할 것.
  - 배경과 컨테이너 구분을 위해 강한 보더 라인 대신 명도 차이(Light Slate vs Pure White)를 활용할 것.
- **DON'T:**
  - `0px` 모서리의 각진 직사각형 사각형 패널 사용 금지.
  - 진하고 짙은 검은색 그림자(`rgba(0,0,0,0.3)` 이상) 사용 금지.