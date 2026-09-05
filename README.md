# EMS 알람 자동화 (PowerShell + UI Automation)

사내 EMS 웹 화면에서 "장비명+알람코드 조회 → 알람명 수정 → 저장 → 저장 후
재조회 검증"을 CSV에 적힌 여러 건에 대해 반복 처리하는 스크립트입니다.
추가 프로그램 설치나 관리자 권한이 필요 없습니다(Windows 기본 PowerShell +
.NET UI Automation만 사용).

## 파일 구성

- `Update-EmsAlarms.ps1`: 실제 자동화를 수행하는 메인 스크립트
- `Get-EmsUiaTree.ps1`: 창 하나를 통째로 골라 그 안의 모든 요소(Name,
  AutomationId, ClassName, 화면 위치, 지원 패턴 등)를 트리 구조 그대로
  텍스트 파일로 뽑아주는 진단 도구 (Accessibility Insights, UIATreeInspector
  등 별도 설치 프로그램을 쓸 수 없는 환경을 위한 대체 도구 — Windows 기본
  PowerShell만으로 동작)
- `Get-ElementAtCursor.ps1`: 마우스를 원하는 요소 위에 놓고 Enter만 누르면
  그 요소(+그 요소를 감싸는 부모/틀 몇 단계)의 정보를 바로 보여주는 도구.
  `Get-EmsUiaTree.ps1`로 뽑은 전체 목록에서 "이게 내가 원하는 그 요소가
  맞나?"를 대조하기 어려울 때 사용
- `data.csv`: 처리할 데이터 샘플 템플릿
- `.gitignore`: 실행 중 생기는 STOP.txt/result.csv/run.log/uia_tree_*.txt 등 제외

## 1단계 — 화면 요소 정보 뽑기 (아직 안 했다면 먼저 이걸부터)

EMS 화면의 입력창/버튼 Name·AutomationId를 아직 모른다면, EMS 페이지를 연
상태에서 `Get-EmsUiaTree.ps1`을 먼저 실행하세요.

```powershell
.\Get-EmsUiaTree.ps1
```

실행하면 현재 열려있는 **엣지(msedge.exe) 창만** 목록으로 `[번호] 프로세스=... PID=... 제목='...'`
형태로 출력됩니다(다른 프로그램 창은 이 자동화와 무관하므로 안 나옴). 그중 EMS가
열려있는 Edge 창의 **번호를 입력**하면(제목만으로 자동 판단하지 않고 직접 눈으로
확인하고 고르는 방식), 그 창 안의 모든 요소가 같은 폴더에 `uia_tree_YYYYMMDD_HHMMSS.txt` 파일로
저장됩니다. Name/AutomationId가 비어있는 이미지 버튼 같은 요소도 빠짐없이
포함됩니다.

파일은 `|`(파이프)로 열이 구분되어 있어 메모장에서 Ctrl+F로 찾아봐도 되고,
엑셀에서 "텍스트 나누기(구분 기호: |)"로 열어도 됩니다. 첫 줄이 열 이름
(`Depth|ControlType|Name|AutomationId|ClassName|HelpText|IsEnabled|BoundingRect|InvokePattern|ValuePattern`)
입니다. 이 안에서 원하는 입력창/버튼 줄을 찾아 Name/AutomationId 값을
`Update-EmsAlarms.ps1`의 `$Config`에 옮겨 적으면 됩니다.

화면이 복잡해서 결과가 너무 많으면 아래 옵션을 참고하세요.

```powershell
# 클릭하거나 입력 가능한 요소만 보기
.\Get-EmsUiaTree.ps1 -OnlyInteractable

# Name이나 ClassName에 특정 문자열이 들어간 요소만 보기
.\Get-EmsUiaTree.ps1 -Filter "알람"

# 너무 깊이 들어가지 않고 얕게만 훑어보기
.\Get-EmsUiaTree.ps1 -MaxDepth 8
```

## 1-2단계 — 특정 요소가 맞는지 바로 대조하기 (선택)

`uia_tree_*.txt`가 너무 커서 원하는 요소를 찾기 어려우면, `Get-ElementAtCursor.ps1`을
쓰세요. 전체를 뒤질 필요 없이 **마우스로 직접 가리키기만** 하면 됩니다.

```powershell
.\Get-ElementAtCursor.ps1
```

실행 후 확인하고 싶은 요소(버튼/입력창/이미지 등) 위에 마우스를 놓고, 콘솔
창으로 다시 포커스를 옮겨 Enter를 누르면 그 요소의 정보가 바로 출력됩니다.
그 요소를 감싸는 부모(틀)도 몇 단계 위까지 같이 보여줘서, 이미지 버튼처럼
Name/AutomationId가 비어있는 요소도 부모 쪽에 식별 정보가 있는지 바로 확인할
수 있습니다. `q` 입력 후 Enter로 종료합니다.

## 2단계 — `$Config` 채우기 (`Update-EmsAlarms.ps1` 상단)

전부 `TODO`로 되어 있는 아래 항목을 실제 값으로 채워야 합니다.

- `$EmsWindowTitleContains`: EMS가 열린 Edge 창의 제목에 포함된 문자열
- `EquipmentNameBox`, `AlarmCodeBox`, `SearchButton`, `AlarmNameBox`, `SaveButton`: 각 컨트롤의 AutomationId/Name/ControlType
- `AlarmNameEditButton`: 알람명을 고치기 전에 눌러야 하는 수정 아이콘(gif 이미지). 이 아이콘은 Name/AutomationId/ClassName이 전부 비어있는 것으로 확인되어, ControlType(`Image`)만으로 그 행 안에서 찾도록 되어 있습니다(행 안에 이미지가 이거 하나뿐이라는 전제). 이름/아이디로는 못 찾으니 이 항목은 값을 채울 필요 없이 그대로 두면 됩니다. (실제 클릭은 이미지 자체가 아니라, 이미지의 부모 중 InvokePattern을 지원하는 첫 번째 요소에 대해 자동으로 수행됩니다 - 이미지 자체는 클릭 이벤트가 없고 그걸 감싸는 부모 쪽에 클릭 동작이 걸려있는 것으로 확인됨)
- `ResultAlarmCodeDisplay`: 조회 결과 표에서 알람코드가 표시되는 요소(결과 로드 판단 + 행 구분 + 저장 전 확인에 사용)
- `SaveSuccessIndicator`: 저장 후 나타나는 "저장되었습니다" 류의 성공 메시지 요소 (필수 — 이게 안 뜨면 실패로 기록됨)
- `$RowContainerAncestorLevels`: 결과 알람코드 요소에서 몇 단계 위로 올라가야 "그 행 전체"(알람명 입력창 + 저장 버튼 포함)가 나오는지. 테스트하면서 맞는 값을 찾으면 됩니다.

## 실행 전 준비물

1. Microsoft Edge에서 EMS 페이지를 미리 열어둘 것 (최소화하지 말 것 — 최소화 상태에서는 렌더링이 늦어질 수 있음. 다른 창 뒤에 두는 건 괜찮음)
2. `data.csv`를 스크립트와 같은 폴더에 준비 (컬럼: `장비명,알람코드,새_알람명`)
3. 위 "2단계"의 `$Config` 값을 실제 화면에 맞게 채워넣기

## 실행 방법

**반드시 Windows PowerShell 5.1(`powershell.exe`)로 실행하세요.**
PowerShell 7(`pwsh.exe`)은 UI Automation 관련 어셈블리가 기본 제공되지
않아 동작이 불안정할 수 있습니다.

```powershell
cd <스크립트가 있는 폴더>
# 최초 1회, 실행 정책 때문에 막히는 경우에만 (관리자 권한 불필요, 현재 세션에만 적용)
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

# 먼저 1~2건만 테스트
.\Update-EmsAlarms.ps1 -TestMode

# 문제없으면 전체 실행
.\Update-EmsAlarms.ps1
```

실행 중에는 각 행을 처리할 때마다 콘솔에 `N/전체건수 처리 중: 장비명, 알람코드`
형태로 진행 상황이 출력됩니다.

## 처리 흐름 (한 행당)

1. 장비명 입력 → 알람코드 입력 → 조회
2. 결과 로드 대기 → 해당 행의 수정 아이콘(gif) 클릭 → 알람명 입력창이 편집 가능해질 때까지 대기 → 새_알람명 입력
3. 저장 전, 화면에 조회된 알람코드가 이 행의 알람코드와 일치하는지 재확인
4. 저장 버튼 클릭 → "저장되었습니다" 류의 성공 메시지가 뜨는지 확인(5초 대기)
5. **같은 장비명+알람코드로 다시 조회 → 화면에 표시된 알람명이 새_알람명과
   실제로 일치하는지 확인.** 성공 메시지만으로는 성공 처리하지 않습니다
   (화면은 성공했다고 뜨는데 실제 저장은 안 되는 경우를 걸러내기 위함).
6. 4~5를 모두 통과해야 최종 "성공"으로 기록

## 실행 중 안전 정지

스크립트와 같은 폴더에 `STOP.txt`라는 빈 텍스트 파일을 만들면, 처리 중이던
행을 끝낸 뒤(중간에 끊지 않음) 다음 행 시작 전에 결과를 저장하고 정상
종료합니다. 탐색기에서 빈 텍스트 파일을 만들어 이름만 `STOP.txt`로
바꾸면 됩니다.

## 결과 확인

- `result.csv`: 행별 처리결과(`성공`/`실패`)와 실패사유
- `run.log`: 실행 로그(타임스탬프 + 진행 상황 포함)

## 알아두면 좋은 주의사항

- 대상 PC에서 `Add-Type -AssemblyName UIAutomationClient` / `UIAutomationTypes`
  실행이 확인됐고, 실행 정책/Constrained Language Mode 차단도 없는 것으로
  확인됐습니다(단, Windows PowerShell 5.1 기준). PowerShell 7에서는 위
  어셈블리 로드가 실패할 수 있으니 문제가 생기면 5.1로 실행해보세요.
- EMS 화면이 React/Vue 등으로 만들어진 경우, 값을 넣어도 화면엔 보이지만
  내부 상태가 갱신되지 않는 경우가 드물게 있습니다. 이런 경우를 대비해
  입력 직후 값 재확인 + 저장 전 재확인 + **저장 후 재조회 검증**까지
  3중으로 확인하도록 만들어져 있습니다. 그래도 처음 몇 건은 실제 EMS
  화면에서 결과가 맞게 저장됐는지 눈으로 꼭 확인하세요.
