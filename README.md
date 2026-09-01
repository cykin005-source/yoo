# EMS 알람 자동화 (PowerShell + UI Automation)

사내 EMS 웹 화면에서 "장비명+알람코드 조회 → 알람명 수정 → 저장"을 CSV에
적힌 여러 건에 대해 반복 처리하는 스크립트입니다. 추가 프로그램 설치나
관리자 권한이 필요 없습니다(Windows 기본 PowerShell + .NET UI Automation만
사용).

## 아직 남은 작업 (중요)

`Update-EmsAlarms.ps1` 상단의 `$Config` 값들이 전부 `TODO`로 채워져 있습니다.
UIATreeInspector 같은 도구로 확인한 실제 값으로 아래 항목을 채운 뒤 사용하세요.

- `$EmsWindowTitleContains`: EMS가 열린 Edge 창의 제목에 포함된 문자열
- `EquipmentNameBox`, `AlarmCodeBox`, `SearchButton`, `AlarmNameBox`, `SaveButton`: 각 컨트롤의 AutomationId/Name/ControlType
- `ResultAlarmCodeDisplay`: 조회 결과 표에서 알람코드가 표시되는 요소(결과 로드 판단 + 행 구분에 사용)
- `SaveSuccessIndicator` / `SaveErrorIndicator`: 저장 성공/실패를 알리는 메시지 요소(선택, 없으면 비워둬도 동작은 함)
- `$RowContainerAncestorLevels`: 결과 알람코드 요소에서 몇 단계 위로 올라가야 "그 행 전체"(알람명 입력창 + 저장 버튼 포함)가 나오는지. 테스트하면서 맞는 값을 찾으면 됩니다.

## 실행 전 준비물

1. Microsoft Edge에서 EMS 페이지를 미리 열어둘 것
2. `data.csv`를 스크립트와 같은 폴더에 준비 (컬럼: `장비명,알람코드,새_알람명`)
3. 위 "아직 남은 작업"의 `$Config` 값을 실제 화면에 맞게 채워넣기

## 실행 방법

PowerShell(5.1 권장)을 열고:

```powershell
cd <스크립트가 있는 폴더>
# 최초 1회, 실행 정책 때문에 막히는 경우에만 (관리자 권한 불필요, 현재 세션에만 적용)
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

# 먼저 1~2건만 테스트
.\Update-EmsAlarms.ps1 -TestMode

# 문제없으면 전체 실행
.\Update-EmsAlarms.ps1
```

## 실행 중 안전 정지

스크립트와 같은 폴더에 `STOP.txt`라는 빈 텍스트 파일을 만들면, 처리 중이던
행을 끝낸 뒤(중간에 끊지 않음) 다음 행 시작 전에 결과를 저장하고 정상
종료합니다. 탐색기에서 빈 텍스트 파일을 만들어 이름만 `STOP.txt`로
바꾸면 됩니다.

## 결과 확인

- `result.csv`: 행별 처리결과(`성공`/`실패`)와 실패사유
- `run.log`: 실행 로그(타임스탬프 포함)

## 알아두면 좋은 주의사항

- PowerShell 7에서는 설치 방식에 따라 UI Automation 관련 어셈블리 로드가
  실패할 수 있습니다. 문제가 생기면 Windows PowerShell 5.1로 실행해보세요.
- EMS 화면이 React/Vue 등으로 만들어진 경우, 값을 넣어도 화면엔 보이지만
  내부 상태가 갱신되지 않는 경우가 드물게 있습니다. 스크립트는 입력 직후
  값을 재확인하고 저장 전에도 한 번 더 확인하지만, 처음 몇 건은 실제 EMS
  화면에서 결과가 맞게 저장됐는지 눈으로 꼭 확인하세요.
