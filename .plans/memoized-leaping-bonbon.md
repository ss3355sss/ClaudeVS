# Fix: Terminal resize and text duplication

## Context

Two reported issues:
1. Panel 리사이즈 시 텍스트가 새 크기에 맞게 출력되지 않음
2. 텍스트 중복 출력 → 스크롤 무한 증가

**Root cause**: ConPTY columns가 `120`으로 하드코딩되어 있어서, 터미널 컨트롤의 실제 렌더링 너비와 불일치 발생. ConPTY가 120컬럼 포맷으로 출력을 보내지만, 터미널 렌더러가 다른 너비로 표시하면 줄바꿈이 발생하면서 텍스트가 중복/겹쳐 보이고 스크롤이 비정상적으로 늘어남.

## Changes

### 1. `ConPtyTerminal.cs` — `Resize` 메서드 (line 600)
- `columns = 120;` 하드코딩 제거 → 호출자가 전달한 columns 값 그대로 사용

### 2. `ClaudeTerminalControl.xaml.cs` — `ClaudeTerminalControl_SizeChanged` (line 257)
- `uint columns = 120;` 제거
- 실제 `ActualWidth`에서 동적으로 columns 계산: `columns = (uint)(actualWidth / charWidth)`
- `charWidth`는 기존 `UpdateTerminalMaxWidth` 하드코딩 값에서 역산: `fontSize * 0.75` (Consolas 모노스페이스 기준)

### 3. `ClaudeTerminalControl.xaml.cs` — `CreateNewAgentTab` (line 574)
- `Border`의 `HorizontalAlignment = Left` → `Stretch`로 변경하여 패널 너비에 맞게 확장

### 4. `ClaudeTerminalControl.xaml.cs` — `UpdateTerminalMaxWidth` (line 1410)
- 하드코딩된 폰트별 고정 너비 제거
- 더 이상 Border 너비를 제한하지 않음 (Stretch로 대체)

### 5. `ClaudeTerminalControl.xaml.cs` — `RefreshTimer_Tick` (line 1018)
- 매 500ms마다 `SetTheme` 호출 제거 → 불필요한 재렌더링 방지 (중복 출력 가능성 차단)

### 6. `ClaudeTerminalControl.xaml.cs` — `ConPtyTerminal` 초기 생성 (line 187)
- `new ConPtyTerminal(rows: 30, columns: 120)` → 초기 columns도 패널 너비 기반으로 계산하되, 아직 크기를 모를 수 있으므로 기본값 유지하고 첫 `NeedsResizeAfterOutput`에서 보정

## Verification
- Build.bat으로 빌드 성공 확인
- VS Experimental Instance에서 ClaudeVS 패널 열기
- 패널 너비를 좁게/넓게 조절하면서 텍스트 출력이 패널 너비에 맞게 조정되는지 확인
- 텍스트 중복 없이 스크롤이 정상 동작하는지 확인
