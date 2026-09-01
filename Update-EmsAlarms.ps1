<#
    EMS 알람코드/알람명 일괄 수정 자동화 스크립트
    ---------------------------------------------
    - PowerShell(Windows PowerShell 5.1 권장) + .NET UI Automation(UIA) 만 사용.
    - 추가 모듈/패키지 설치 없음. 관리자 권한 불필요.
    - 좌표클릭 / SendKeys / Tab 이동 사용하지 않음. 전부 UIA 패턴 기반.

    !!! 사용 전 필수 작업 !!!
    아래 "0. 환경설정" 구역의 $Config 값들이 전부 TODO 로 되어 있습니다.
    UIATreeInspector 등으로 확인한 실제 Name/AutomationId/ControlType 값으로
    반드시 채워 넣은 뒤 사용하세요. 채우지 않으면 스크립트가 요소를 찾지 못하고
    타임아웃으로 실패 처리됩니다.
#>

[CmdletBinding()]
param(
    # 테스트 모드: data.csv 전체가 아니라 앞의 $TestModeRows 건만 처리
    [switch]$TestMode,
    [int]$TestModeRows = 2
)

# ===================================================================
# 0. 환경설정 - EMS 화면이 바뀌면 이 구역만 수정하면 됩니다.
# ===================================================================

# EMS 페이지가 열려있는 Edge "창"을 찾기 위한 기준(창 제목에 포함된 문자열).
# 예: 창 제목이 "EMS 알람관리 - Microsoft Edge" 라면 "EMS" 로 설정.
$EmsWindowTitleContains = "TODO: EMS 창 제목 일부"

# 각 컨트롤 식별 정보. AutomationId 를 아는 경우 AutomationId 를 우선 사용하고,
# 모르면 Name + ControlType 조합으로 찾습니다. 둘 다 비어있지 않으면 AND 조건으로 좁혀집니다.
# ControlType 은 System.Windows.Automation.ControlType 의 정적 필드명을 문자열로 적으세요.
# (자주 쓰는 값: Edit, Button, Text, DataItem, ListItem, Group, Pane, Document)
$Config = @{
    EquipmentNameBox = @{ AutomationId = "TODO"; Name = "TODO"; ControlType = "Edit" }   # 장비명 입력창
    AlarmCodeBox     = @{ AutomationId = "TODO"; Name = "TODO"; ControlType = "Edit" }   # 알람코드 입력창
    SearchButton     = @{ AutomationId = "TODO"; Name = "조회"; ControlType = "Button" } # 조회 버튼
    AlarmNameBox     = @{ AutomationId = "TODO"; Name = "TODO"; ControlType = "Edit" }   # 알람명 입력창(결과 행 내부)
    SaveButton       = @{ AutomationId = "TODO"; Name = "저장"; ControlType = "Button" } # 저장 버튼(결과 행 내부)

    # 조회 결과가 로드됐는지 + 어느 행인지 판단하기 위한 "결과 알람코드 표시" 요소.
    # 보통 결과 표(그리드)에 알람코드 값이 텍스트로 표시되는 셀입니다.
    ResultAlarmCodeDisplay = @{ AutomationId = "TODO"; Name = "TODO"; ControlType = "Text" }

    # (선택) 저장 성공/실패를 알려주는 메시지 요소가 있다면 채우세요. 없으면 TODO 로 두고,
    # 아래 로직은 "예외 없이 저장 버튼 클릭 완료 = 성공(확인필요)" 로 보수적으로 기록합니다.
    SaveSuccessIndicator = @{ AutomationId = "TODO"; Name = "TODO"; ControlType = "Text" }
    SaveErrorIndicator   = @{ AutomationId = "TODO"; Name = "TODO"; ControlType = "Text" }
}

# ResultAlarmCodeDisplay 요소에서 몇 단계 위로 올라가야 "그 행 전체"(알람명 입력창 +
# 저장 버튼을 함께 포함하는 컨테이너)가 나오는지. 실제 EMS 표 구조를 보고 조정하세요.
# (모르면 3~5 사이 값으로 시험해보면서 맞는 값을 찾으면 됩니다.)
$RowContainerAncestorLevels = 4

$TimeoutSec = 10          # 폴링 대기 최대 시간(초)
$PollingIntervalMs = 300  # 폴링 간격(ms)

# ===================================================================
# 경로 설정
# ===================================================================
$ScriptDir    = Split-Path -Parent $MyInvocation.MyCommand.Path
$DataCsvPath  = Join-Path $ScriptDir "data.csv"
$ResultCsvPath = Join-Path $ScriptDir "result.csv"
$StopFlagPath = Join-Path $ScriptDir "STOP.txt"
$LogPath      = Join-Path $ScriptDir "run.log"

# ===================================================================
# 1. UIA 어셈블리 로드 (Windows 기본 포함, 설치 불필요)
# ===================================================================
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

function Write-Log {
    param([string]$Message)
    $line = "[{0:yyyy-MM-dd HH:mm:ss}] {1}" -f (Get-Date), $Message
    Write-Host $line
    Add-Content -Path $LogPath -Value $line
}

# ===================================================================
# 2. 조건(Condition) 생성 헬퍼
# ===================================================================

function Get-ControlTypeByName {
    param([string]$Name)
    $field = [System.Windows.Automation.ControlType].GetField($Name, [System.Reflection.BindingFlags]::Public -bor [System.Reflection.BindingFlags]::Static)
    if (-not $field) { throw "알 수 없는 ControlType 이름: $Name" }
    return $field.GetValue($null)
}

# $Config 항목(해시테이블) 하나를 받아서 AutomationId/Name/ControlType 중
# TODO 가 아닌 값들만으로 AND 조건을 만듭니다.
function New-ConditionFromConfig {
    param([hashtable]$Item)

    $conditions = New-Object System.Collections.Generic.List[System.Windows.Automation.Condition]

    if ($Item.AutomationId -and $Item.AutomationId -ne "TODO") {
        $conditions.Add((New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::AutomationIdProperty, $Item.AutomationId)))
    }
    if ($Item.Name -and $Item.Name -ne "TODO") {
        $conditions.Add((New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::NameProperty, $Item.Name)))
    }
    if ($Item.ControlType -and $Item.ControlType -ne "TODO") {
        $ct = Get-ControlTypeByName $Item.ControlType
        $conditions.Add((New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::ControlTypeProperty, $ct)))
    }

    if ($conditions.Count -eq 0) {
        throw "요소 조건이 비어있습니다. `$Config 값을 채워주세요."
    }
    if ($conditions.Count -eq 1) {
        return $conditions[0]
    }
    return New-Object System.Windows.Automation.AndCondition($conditions.ToArray())
}

# ===================================================================
# 3. 대상 창 찾기 - PID 로 고정 (창 활성화/포커스 이동 없이)
# ===================================================================

function Find-EmsWindow {
    param([string]$TitleContains)

    $root = [System.Windows.Automation.AutomationElement]::RootElement
    $windowCondition = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
        [System.Windows.Automation.ControlType]::Window)

    $candidates = $root.FindAll([System.Windows.Automation.TreeScope]::Children, $windowCondition)

    $matches = @()
    foreach ($w in $candidates) {
        try {
            if ($w.Current.Name -like "*$TitleContains*") {
                $matches += $w
            }
        } catch { }
    }

    if ($matches.Count -eq 0) {
        throw "제목에 '$TitleContains' 를 포함하는 창을 찾지 못했습니다. EMS 페이지가 열려 있는지 확인하세요."
    }
    if ($matches.Count -gt 1) {
        Write-Log "경고: 제목이 일치하는 창이 ${($matches.Count)}개 발견됨. 첫 번째 창을 사용합니다."
        foreach ($m in $matches) { Write-Log "  후보 창: '$($m.Current.Name)' (PID=$($m.Current.ProcessId))" }
    }

    return $matches[0]
}

# ===================================================================
# 4. 요소 찾기 / 대기 (고정 Sleep 대신 폴링)
# ===================================================================

function Find-ElementNow {
    param(
        [System.Windows.Automation.AutomationElement]$Parent,
        [System.Windows.Automation.Condition]$Condition,
        [System.Windows.Automation.TreeScope]$Scope = [System.Windows.Automation.TreeScope]::Descendants
    )
    return $Parent.FindFirst($Scope, $Condition)
}

function Wait-UIAElement {
    param(
        [System.Windows.Automation.AutomationElement]$Parent,
        [System.Windows.Automation.Condition]$Condition,
        [int]$TimeoutSec = 10
    )
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while ($sw.Elapsed.TotalSeconds -lt $TimeoutSec) {
        $el = Find-ElementNow -Parent $Parent -Condition $Condition
        if ($el) { return $el }
        Start-Sleep -Milliseconds $PollingIntervalMs
    }
    return $null
}

# 알람코드 값이 일치하는 결과 요소를 찾을 때까지 대기 (조회 결과 로드 판단 + 행 식별 겸용)
function Wait-ResultRow {
    param(
        [System.Windows.Automation.AutomationElement]$Window,
        [string]$ExpectedAlarmCode,
        [int]$TimeoutSec = 10
    )
    $baseCondition = New-ConditionFromConfig $Config.ResultAlarmCodeDisplay
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while ($sw.Elapsed.TotalSeconds -lt $TimeoutSec) {
        $candidates = $Window.FindAll([System.Windows.Automation.TreeScope]::Descendants, $baseCondition)
        foreach ($c in $candidates) {
            try {
                $val = $null
                $vp = $null
                if ($c.TryGetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern, [ref]$vp)) {
                    $val = $vp.Current.Value
                } else {
                    $val = $c.Current.Name
                }
                if ($val -eq $ExpectedAlarmCode) {
                    return $c
                }
            } catch { }
        }
        Start-Sleep -Milliseconds $PollingIntervalMs
    }
    return $null
}

# 결과 요소에서 위로 $Levels 단계 올라가 "행 컨테이너"를 반환
function Get-RowContainer {
    param(
        [System.Windows.Automation.AutomationElement]$Element,
        [int]$Levels
    )
    $walker = [System.Windows.Automation.TreeWalker]::ControlViewWalker
    $current = $Element
    for ($i = 0; $i -lt $Levels; $i++) {
        $parent = $walker.GetParent($current)
        if (-not $parent) { break }
        $current = $parent
    }
    return $current
}

# ===================================================================
# 5. 값 입력 (ValuePattern) - 입력 후 재확인 포함
# ===================================================================

function Set-UiaValue {
    param(
        [System.Windows.Automation.AutomationElement]$Element,
        [string]$Value
    )
    $pattern = $Element.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern)
    $pattern.SetValue($Value)

    # 실제로 반영됐는지 재확인 (일부 웹 프레임워크는 SetValue 가 화면엔 보이되
    # 내부 상태에 반영 안 되는 경우가 있어, 최소한 "표시값"은 확인합니다)
    Start-Sleep -Milliseconds 150
    $actual = $Element.GetCurrentPropertyValue([System.Windows.Automation.ValuePattern]::ValueProperty)
    if ($actual -ne $Value) {
        throw "값 입력 검증 실패 (입력='$Value', 실제='$actual')"
    }
}

# ===================================================================
# 6. 버튼 클릭 (InvokePattern 우선, 미지원 시 LegacyIAccessiblePattern)
# ===================================================================

function Invoke-UiaClick {
    param([System.Windows.Automation.AutomationElement]$Element)

    $invokePattern = $null
    if ($Element.TryGetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern, [ref]$invokePattern)) {
        $invokePattern.Invoke()
        return
    }

    # InvokePattern 미지원 요소에 대한 대안: LegacyIAccessiblePattern.DoDefaultAction()
    # (오래된 웹 컨트롤/커스텀 컴포넌트가 IInvokeProvider 를 구현하지 않을 때 사용)
    $legacyPattern = $null
    if ($Element.TryGetCurrentPattern([System.Windows.Automation.LegacyIAccessiblePattern]::Pattern, [ref]$legacyPattern)) {
        $legacyPattern.DoDefaultAction()
        return
    }

    throw "이 요소는 InvokePattern과 LegacyIAccessiblePattern을 모두 지원하지 않습니다."
}

# ===================================================================
# 7. 안전 정지(STOP.txt) 확인
# ===================================================================

function Test-StopRequested {
    return (Test-Path -LiteralPath $StopFlagPath)
}

# ===================================================================
# 8. 한 행 처리
# ===================================================================

function Process-Row {
    param(
        [System.Windows.Automation.AutomationElement]$Window,
        [pscustomobject]$Row
    )

    $result = [pscustomobject]@{
        장비명   = $Row.장비명
        알람코드 = $Row.알람코드
        처리결과 = "실패"
        실패사유 = ""
    }

    try {
        # 1) 장비명 입력
        $cond = New-ConditionFromConfig $Config.EquipmentNameBox
        $el = Find-ElementNow -Parent $Window -Condition $cond
        if (-not $el) { throw "장비명 입력창을 찾지 못했습니다." }
        Set-UiaValue -Element $el -Value $Row.장비명

        # 2) 알람코드 입력
        $cond = New-ConditionFromConfig $Config.AlarmCodeBox
        $el = Find-ElementNow -Parent $Window -Condition $cond
        if (-not $el) { throw "알람코드 입력창을 찾지 못했습니다." }
        Set-UiaValue -Element $el -Value $Row.알람코드

        # 3) 조회 버튼 클릭
        $cond = New-ConditionFromConfig $Config.SearchButton
        $el = Find-ElementNow -Parent $Window -Condition $cond
        if (-not $el) { throw "조회 버튼을 찾지 못했습니다." }
        Invoke-UiaClick -Element $el

        # 4) 조회 결과 로드 대기 + 해당 행(알람코드 일치) 컨테이너 확보
        $resultCell = Wait-ResultRow -Window $Window -ExpectedAlarmCode $Row.알람코드 -TimeoutSec $TimeoutSec
        if (-not $resultCell) { throw "조회 결과 대기 시간 초과(${TimeoutSec}초) 또는 알람코드 불일치." }
        $rowContainer = Get-RowContainer -Element $resultCell -Levels $RowContainerAncestorLevels

        # 5) 알람명 입력창 (해당 행 범위 안에서만 검색 - 중복 이름 대응)
        $cond = New-ConditionFromConfig $Config.AlarmNameBox
        $el = Find-ElementNow -Parent $rowContainer -Condition $cond
        if (-not $el) { throw "행 내에서 알람명 입력창을 찾지 못했습니다. RowContainerAncestorLevels 값을 조정해보세요." }
        Set-UiaValue -Element $el -Value $Row.새_알람명

        # 6) 저장 전 최종 확인: 화면에 조회된 알람코드가 지금 처리 중인 행과 일치하는지 재확인
        $vp = $null
        $currentAlarmCode = $null
        if ($resultCell.TryGetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern, [ref]$vp)) {
            $currentAlarmCode = $vp.Current.Value
        } else {
            $currentAlarmCode = $resultCell.Current.Name
        }
        if ($currentAlarmCode -ne $Row.알람코드) {
            throw "저장 전 확인 실패: 화면 알람코드='$currentAlarmCode', 기대값='$($Row.알람코드)'. 저장하지 않고 실패 처리합니다."
        }

        # 7) 저장 버튼 클릭 (행 범위 안에서만 검색 - 중복 이름 대응)
        $cond = New-ConditionFromConfig $Config.SaveButton
        $el = Find-ElementNow -Parent $rowContainer -Condition $cond
        if (-not $el) { throw "행 내에서 저장 버튼을 찾지 못했습니다." }
        Invoke-UiaClick -Element $el

        # 8) 성공/실패 판단
        #    SaveSuccessIndicator / SaveErrorIndicator 를 설정했다면 그걸로 판단하고,
        #    설정 안 했다면(TODO 상태) 예외 없이 여기까지 왔다는 것만으로 "성공(확인필요)" 처리.
        if ($Config.SaveErrorIndicator.Name -ne "TODO" -or $Config.SaveErrorIndicator.AutomationId -ne "TODO") {
            $errCond = New-ConditionFromConfig $Config.SaveErrorIndicator
            $errEl = Wait-UIAElement -Parent $Window -Condition $errCond -TimeoutSec 3
            if ($errEl) {
                throw "저장 실패 메시지 감지됨: $($errEl.Current.Name)"
            }
        }

        $result.처리결과 = "성공"
        Write-Log "성공: 장비명='$($Row.장비명)' 알람코드='$($Row.알람코드)'"
    }
    catch {
        $result.처리결과 = "실패"
        $result.실패사유 = $_.Exception.Message
        Write-Log "실패: 장비명='$($Row.장비명)' 알람코드='$($Row.알람코드)' 사유='$($_.Exception.Message)'"
    }

    return $result
}

# ===================================================================
# 9. 메인 실행부
# ===================================================================

Write-Log "===== EMS 알람 자동화 시작 ====="

# 이전 실행의 STOP.txt 잔재 정리
if (Test-Path -LiteralPath $StopFlagPath) {
    Remove-Item -LiteralPath $StopFlagPath -Force
    Write-Log "이전 실행에서 남은 STOP.txt 를 삭제했습니다."
}

if (-not (Test-Path -LiteralPath $DataCsvPath)) {
    Write-Log "오류: data.csv 를 찾을 수 없습니다 ($DataCsvPath). 스크립트를 종료합니다."
    exit 1
}

$rows = Import-Csv -Path $DataCsvPath -Encoding UTF8
if ($TestMode) {
    Write-Log "테스트 모드: 앞의 $TestModeRows 건만 처리합니다."
    $rows = $rows | Select-Object -First $TestModeRows
}

Write-Log "대상 EMS 창을 찾는 중 (제목에 '$EmsWindowTitleContains' 포함)..."
$window = Find-EmsWindow -TitleContains $EmsWindowTitleContains
$targetPid = $window.Current.ProcessId
Write-Log "대상 창 확보. PID=$targetPid, 제목='$($window.Current.Name)'"

$results = New-Object System.Collections.Generic.List[pscustomobject]

foreach ($row in $rows) {

    if (Test-StopRequested) {
        Write-Log "STOP.txt 감지됨. 현재까지 결과를 저장하고 종료합니다."
        break
    }

    $r = Process-Row -Window $window -Row $row
    $results.Add($r)
}

$results | Export-Csv -Path $ResultCsvPath -NoTypeInformation -Encoding UTF8

$successCount = ($results | Where-Object { $_.처리결과 -eq "성공" }).Count
$failCount = ($results | Where-Object { $_.처리결과 -eq "실패" }).Count
Write-Log "===== 종료. 성공 $successCount 건 / 실패 $failCount 건. 결과: $ResultCsvPath ====="
