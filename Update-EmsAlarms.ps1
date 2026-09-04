<#
    EMS 알람코드/알람명 일괄 수정 자동화 스크립트
    ---------------------------------------------
    - Windows PowerShell 5.1(powershell.exe) 기준. PowerShell 7(pwsh.exe)은
      UIAutomationClient 참조 어셈블리가 기본 제공되지 않아 동작이 불안정할 수
      있으므로 사용하지 않습니다.
    - .NET UI Automation(UIA) 만 사용. 추가 모듈/패키지 설치, 관리자 권한 불필요.
    - 좌표클릭 / SendKeys / Tab 이동 사용하지 않음. 전부 UIA 패턴 기반.
    - 저장 후에는 같은 조건으로 재조회하여 실제로 값이 반영됐는지까지 확인합니다
      (저장 성공 메시지만으로는 성공 처리하지 않음).

    !!! 사용 전 필수 작업 !!!
    아래 "0. 환경설정" 구역의 $Config 값들이 전부 TODO 로 되어 있습니다.
    실제 화면의 요소 정보를 아직 모른다면, 같은 폴더의 Get-EmsUiaTree.ps1 을
    먼저 실행해 요소 목록을 뽑아보세요(별도 설치 프로그램 없이 동작합니다).
    확인한 Name/AutomationId/ControlType 값으로 $Config 를 채운 뒤 사용하세요.
    채우지 않으면 스크립트가 요소를 찾지 못하고 타임아웃으로 실패 처리됩니다.
#>

[CmdletBinding()]
param(
    # 테스트 모드: data.csv 전체가 아니라 앞의 $TestModeRows 건만 처리
    [switch]$TestMode,
    [int]$TestModeRows = 2
)

# 실행 환경 확인: 반드시 Windows PowerShell 5.1(Desktop 에디션)에서 실행할 것.
# PowerShell 7(Core 에디션)에서는 UIAutomationClient 로드/동작이 불안정할 수 있습니다.
if ($PSVersionTable.PSEdition -ne 'Desktop') {
    Write-Warning "이 스크립트는 Windows PowerShell 5.1(powershell.exe) 기준으로 검증되었습니다. 현재 PSEdition='$($PSVersionTable.PSEdition)' 입니다. 문제가 생기면 powershell.exe 로 실행해보세요."
}

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

    # 저장 후 나타나는 "저장되었습니다" 류의 성공 메시지 요소. 필수입니다.
    # (성공 메시지만으로는 최종 성공 처리하지 않고, 이후 재조회 검증까지 통과해야 최종 성공)
    SaveSuccessIndicator = @{ AutomationId = "TODO"; Name = "TODO"; ControlType = "Text" }
}

# ResultAlarmCodeDisplay 요소에서 몇 단계 위로 올라가야 "그 행 전체"(알람명 입력창 +
# 저장 버튼을 함께 포함하는 컨테이너)가 나오는지. 실제 EMS 표 구조를 보고 조정하세요.
# (모르면 3~5 사이 값으로 시험해보면서 맞는 값을 찾으면 됩니다.)
$RowContainerAncestorLevels = 4

$TimeoutSec = 10             # 조회 결과 대기 타임아웃(초)
$SaveConfirmTimeoutSec = 5   # 저장 성공 메시지 대기 타임아웃(초)
$PollingIntervalMs = 300     # 폴링 간격(ms)

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
#    주의: EMS 창은 최소화하지 말 것. Chromium 기반 브라우저는 최소화 상태에서
#    렌더링/접근성 트리 갱신이 늦어질 수 있음(다른 창 뒤에 두는 것은 무방).
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
        Write-Log "경고: 제목이 일치하는 창이 $($matches.Count)개 발견됨. 첫 번째 창을 사용합니다."
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
                $val = Get-ElementDisplayValue -Element $c
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

# ValuePattern 을 지원하면 그 값을, 아니면 Name 속성을 화면 표시값으로 간주
function Get-ElementDisplayValue {
    param([System.Windows.Automation.AutomationElement]$Element)
    $vp = $null
    if ($Element.TryGetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern, [ref]$vp)) {
        return $vp.Current.Value
    }
    return $Element.Current.Name
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
# 8. 장비명+알람코드로 조회 실행 (1~3단계 공통 로직: 최초 조회 / 저장 후 재조회 둘 다 사용)
# ===================================================================

function Invoke-EmsSearch {
    param(
        [System.Windows.Automation.AutomationElement]$Window,
        [string]$EquipmentName,
        [string]$AlarmCode
    )

    $cond = New-ConditionFromConfig $Config.EquipmentNameBox
    $el = Find-ElementNow -Parent $Window -Condition $cond
    if (-not $el) { throw "장비명 입력창을 찾지 못했습니다." }
    Set-UiaValue -Element $el -Value $EquipmentName

    $cond = New-ConditionFromConfig $Config.AlarmCodeBox
    $el = Find-ElementNow -Parent $Window -Condition $cond
    if (-not $el) { throw "알람코드 입력창을 찾지 못했습니다." }
    Set-UiaValue -Element $el -Value $AlarmCode

    $cond = New-ConditionFromConfig $Config.SearchButton
    $el = Find-ElementNow -Parent $Window -Condition $cond
    if (-not $el) { throw "조회 버튼을 찾지 못했습니다." }
    Invoke-UiaClick -Element $el
}

# 조회 실행 + 결과 로드 대기 + 행 컨테이너 확보를 한 번에 처리
function Invoke-EmsSearchAndGetRow {
    param(
        [System.Windows.Automation.AutomationElement]$Window,
        [string]$EquipmentName,
        [string]$AlarmCode
    )

    Invoke-EmsSearch -Window $Window -EquipmentName $EquipmentName -AlarmCode $AlarmCode

    $resultCell = Wait-ResultRow -Window $Window -ExpectedAlarmCode $AlarmCode -TimeoutSec $TimeoutSec
    if (-not $resultCell) { throw "조회 결과 대기 시간 초과(${TimeoutSec}초) 또는 알람코드 불일치." }

    return Get-RowContainer -Element $resultCell -Levels $RowContainerAncestorLevels
}

# ===================================================================
# 9. 한 행 처리 (조회 → 알람명 수정 → 저장 → 저장 후 재조회 검증)
# ===================================================================

function Process-Row {
    param(
        [System.Windows.Automation.AutomationElement]$Window,
        [pscustomobject]$Row,
        [int]$Index,
        [int]$Total
    )

    Write-Log ("{0}/{1} 처리 중: {2}, {3}" -f $Index, $Total, $Row.장비명, $Row.알람코드)

    $result = [pscustomobject]@{
        장비명   = $Row.장비명
        알람코드 = $Row.알람코드
        처리결과 = "실패"
        실패사유 = ""
    }

    try {
        # 1~4) 장비명/알람코드 입력 → 조회 → 결과 로드 대기 → 행 컨테이너 확보
        $rowContainer = Invoke-EmsSearchAndGetRow -Window $Window -EquipmentName $Row.장비명 -AlarmCode $Row.알람코드

        # 5) 알람명 입력창 (해당 행 범위 안에서만 검색 - 중복 이름 대응)
        $cond = New-ConditionFromConfig $Config.AlarmNameBox
        $alarmNameEl = Find-ElementNow -Parent $rowContainer -Condition $cond
        if (-not $alarmNameEl) { throw "행 내에서 알람명 입력창을 찾지 못했습니다. RowContainerAncestorLevels 값을 조정해보세요." }
        Set-UiaValue -Element $alarmNameEl -Value $Row.새_알람명

        # 6) 저장 전 최종 확인: 화면에 조회된 알람코드가 지금 처리 중인 행과 일치하는지 재확인
        $cond = New-ConditionFromConfig $Config.ResultAlarmCodeDisplay
        $codeEl = Find-ElementNow -Parent $rowContainer -Condition $cond
        $currentAlarmCode = if ($codeEl) { Get-ElementDisplayValue -Element $codeEl } else { $null }
        if ($currentAlarmCode -ne $Row.알람코드) {
            throw "저장 전 확인 실패: 화면 알람코드='$currentAlarmCode', 기대값='$($Row.알람코드)'. 저장하지 않고 실패 처리합니다."
        }

        # 7) 저장 버튼 클릭 (행 범위 안에서만 검색 - 중복 이름 대응)
        $cond = New-ConditionFromConfig $Config.SaveButton
        $saveEl = Find-ElementNow -Parent $rowContainer -Condition $cond
        if (-not $saveEl) { throw "행 내에서 저장 버튼을 찾지 못했습니다." }
        Invoke-UiaClick -Element $saveEl

        # 8) 저장 성공 메시지 폴링 확인 (타임아웃: $SaveConfirmTimeoutSec)
        $cond = New-ConditionFromConfig $Config.SaveSuccessIndicator
        $successEl = Wait-UIAElement -Parent $Window -Condition $cond -TimeoutSec $SaveConfirmTimeoutSec
        if (-not $successEl) {
            throw "저장 성공 메시지를 ${SaveConfirmTimeoutSec}초 내에 확인하지 못했습니다."
        }

        # 9) 저장 후 재조회 검증 (가장 중요): 성공 메시지만 믿지 않고, 같은 조건으로 다시 조회해서
        #    화면에 실제로 표시되는 알람명이 CSV의 새_알람명과 일치하는지 확인해야 최종 성공.
        $verifyRowContainer = Invoke-EmsSearchAndGetRow -Window $Window -EquipmentName $Row.장비명 -AlarmCode $Row.알람코드

        $cond = New-ConditionFromConfig $Config.AlarmNameBox
        $verifyNameEl = Find-ElementNow -Parent $verifyRowContainer -Condition $cond
        if (-not $verifyNameEl) { throw "저장 후 재조회 결과에서 알람명 요소를 찾지 못했습니다." }

        $displayedName = Get-ElementDisplayValue -Element $verifyNameEl
        if ($displayedName -ne $Row.새_알람명) {
            throw "저장 후 검증 실패: 화면 알람명='$displayedName', 기대값='$($Row.새_알람명)'. (성공 메시지는 떴으나 실제 반영은 안 됐을 수 있음)"
        }

        $result.처리결과 = "성공"
        Write-Log "성공: 장비명='$($Row.장비명)' 알람코드='$($Row.알람코드)' → 새_알람명='$($Row.새_알람명)' (재조회 검증 통과)"
    }
    catch {
        $result.처리결과 = "실패"
        $result.실패사유 = $_.Exception.Message
        Write-Log "실패: 장비명='$($Row.장비명)' 알람코드='$($Row.알람코드)' 사유='$($_.Exception.Message)'"
    }

    return $result
}

# ===================================================================
# 10. 메인 실행부
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
$total = $rows.Count
$index = 0

foreach ($row in $rows) {
    $index++

    if (Test-StopRequested) {
        Write-Log "STOP.txt 감지됨. 현재까지 결과를 저장하고 종료합니다."
        break
    }

    $r = Process-Row -Window $window -Row $row -Index $index -Total $total
    $results.Add($r)
}

$results | Export-Csv -Path $ResultCsvPath -NoTypeInformation -Encoding UTF8

$successCount = ($results | Where-Object { $_.처리결과 -eq "성공" }).Count
$failCount = ($results | Where-Object { $_.처리결과 -eq "실패" }).Count
Write-Log "===== 종료. 성공 $successCount 건 / 실패 $failCount 건. 결과: $ResultCsvPath ====="
