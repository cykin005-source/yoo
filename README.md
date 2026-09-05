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

## 2단계 — `$Config` (이미 실제 값으로 채워져 있음)

실제 화면에서 확인한 값으로 `Update-EmsAlarms.ps1`의 `$Config`가 이미 채워져
있습니다. 화면이 바뀌면 아래 항목만 수정하면 됩니다.

- `$EmsWindowTitleContains` = `"Search Error Code"`, `$PopupWindowTitleContains` = `"Search and Select List of Values"`
- `EquipmentNameBox`(SEmNo), `AlarmCodeBox`(SPlcErrCode): 장비명/에러코드 입력창
- `LookupLink`: "찾아보기" 링크 (Hyperlink, Name="Search: 설비 에러 코드") — 클릭하면 값 선택 팝업이 뜸
- `PopupRadioItem`(RadioButton, Name="Select"), `PopupConfirmButton`(Button, Name="Select"): 팝업에서 후보 선택 + 확정. 팝업에는 항상 첫 번째 라디오만 선택하도록 되어 있습니다(정확한 코드로 검색했기 때문에 첫 번째가 항상 정답)
- `SearchButton`(Find): 실제 조회 버튼
- `EditIconLink`(SearchAlarmCdTable:Update:0): 연필 모양 아이콘 — 클릭하면 같은 창 안에서 "Update Error Code" 화면으로 전환됨
- `AlarmNameBox`(cPlcErrDesc), `StatusBox`(NStatus): 둘 다 새_알람명과 같은 값이 입력됨
- `GenerateButton`(NGenerate), `SaveButton`(SaveButton), `ErrorListLink`(XXEMSSTD052)

## 실행 전 준비물

1. Microsoft Edge에서 EMS의 "Search Error Code" 화면을 미리 열어둘 것 (최소화하지 말 것)
2. `data.csv`를 스크립트와 같은 폴더에 준비 (컬럼: `장비명,알람코드,새_알람명`)

## 실행 방법

**반드시 Windows PowerShell 5.1(`powershell.exe`)로 실행하세요.**

```powershell
cd <스크립트가 있는 폴더>
# 최초 1회, 실행 정책 때문에 막히는 경우에만
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

# 먼저 1~2건만 테스트
.\Update-EmsAlarms.ps1 -TestMode

# 문제없으면 전체 실행
.\Update-EmsAlarms.ps1
```

실행하면 콘솔 로그와 함께, **실시간으로 갱신되는 결과 창(표)** 이 하나 뜹니다.
행을 처리할 때마다 이 창에 바로 한 줄씩 추가됩니다.

## 처리 흐름 (한 행당)

1. 장비명 + 알람코드 입력 → "찾아보기" 클릭 → 값 선택 팝업에서 첫 번째 후보 선택 → 확정
2. "조회" 클릭 → 연필 아이콘 클릭 → "Update Error Code" 화면으로 전환
3. 에러명(cPlcErrDesc) + 상태(NStatus)에 새_알람명 입력 → "생성" 클릭
4. 생성된 텍스트에 새_알람명이 포함되는지 확인(저장 전 확인) → "저장" 클릭
5. "Error List"로 복귀 → **1~2번을 처음부터 다시 반복해서 재조회** →
   화면에 표시된 에러명이 새_알람명과 실제로 일치하는지 최종 확인
6. 일치해야 최종 "성공"으로 기록, 그 자리에서 다시 "Error List"로 복귀 후 다음 행 진행

## 실행 중 안전 정지

스크립트와 같은 폴더에 `STOP.txt`라는 빈 텍스트 파일을 만들면, 처리 중이던
행을 끝낸 뒤(중간에 끊지 않음) 다음 행 시작 전에 멈춥니다. (클릭식 일시정지
버튼은 없음 — 파일 방식만 사용)

## 결과 확인

- `result.csv`: 행마다 즉시 추가 저장(장비명/알람코드/새_알람명/검증완료/소요시간초/처리결과/실패사유)
- 화면의 실시간 결과 창: 같은 내용을 표 형태로 실시간 표시 (처리 완료 후에도 창은 닫힐 때까지 유지됨)
- `run.log`: 실행 로그(타임스탬프 + 진행 상황 포함)

## 알아두면 좋은 주의사항

- 대상 PC에서 `Add-Type -AssemblyName UIAutomationClient` / `UIAutomationTypes`
  실행이 확인됐고, 실행 정책/Constrained Language Mode 차단도 없는 것으로
  확인됐습니다(단, Windows PowerShell 5.1 기준).
- 이 PC에서는 `LegacyIAccessiblePattern` 타입 자체가 없거나 못 찾는 것으로
  확인되어, 관련 코드는 리플렉션으로 안전하게 우회하도록 되어 있습니다
  (없으면 조용히 "미지원"으로 처리하고 계속 진행).
- "Search Error Code" ↔ "Update Error Code"는 같은 창(같은 PID) 안에서
  내용만 바뀌는 것으로 확인되어 창을 한 번만 고정해서 계속 씁니다. 값 선택
  팝업("Search and Select List of Values")만 별도의 새 창으로 처리합니다.
