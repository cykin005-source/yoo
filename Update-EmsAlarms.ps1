<#
    EMS 알람코드/알람명 일괄 수정 자동화 스크립트
    ---------------------------------------------
    - Windows PowerShell 5.1(powershell.exe) 기준. PowerShell 7(pwsh.exe)은
      UIAutomationClient 참조 어셈블리가 기본 제공되지 않아 동작이 불안정할 수
      있으므로 사용하지 않습니다.
    - .NET UI Automation(UIA) 만 사용. 추가 모듈/패키지 설치, 관리자 권한 불필요.
    - 좌표클릭 / SendKeys / Tab 이동 사용하지 않음. 전부 UIA 패턴 기반.

    [실제 화면 흐름 - 실사용자 확인 완료]
    1) "Search Error Code" 창에서 장비명(SEmNo) + 설비 에러 코드(SPlcErrCode) 입력
    2) 찾아보기 링크 클릭 -> "Search and Select List of Values" 팝업이 뜸
       (팝업 안의 라디오버튼은 항상 첫 번째가 정답 - 정확한 코드로 검색했으므로)
    3) 팝업에서 첫 번째 라디오 선택 -> Select 버튼으로 확정 -> 팝업 닫힘
    4) "조회" 버튼 클릭
    5) 연필 모양 아이콘(Hyperlink, Invoke 지원 확인됨) 클릭
       -> 같은 창 안에서 "Update Error Code" 화면으로 전환됨(새 창 아님)
    6) 에러명(cPlcErrDesc) + 상태(NStatus)에 동일한 새 알람명 입력
    7) "생성" 버튼 클릭 -> 생성된 텍스트에 새 알람명이 포함되는지 확인(저장 전 확인)
    8) "저장" 버튼 클릭
    9) "Error List" 링크 클릭 -> "Search Error Code" 화면으로 복귀
    10) 저장이 실제로 반영됐는지 재검증: 1~5번을 다시 반복해서 에러명이
        새 알람명과 일치하는지 확인 -> 다시 Error List로 복귀

    처리 결과는 result.csv 파일에 행마다 즉시 추가 저장되고, 동시에 화면에 뜨는
    실시간 결과 창(표)에도 즉시 반영됩니다.
#>

[CmdletBinding()]
param(
    # 테스트 모드: data.csv 전체가 아니라 앞의 $TestModeRows 건만 처리
    [switch]$TestMode,
    [int]$TestModeRows = 2
)

if ($PSVersionTable.PSEdition -ne 'Desktop') {
    Write-Warning "이 스크립트는 Windows PowerShell 5.1(powershell.exe) 기준으로 검증되었습니다. 현재 PSEdition='$($PSVersionTable.PSEdition)' 입니다."
}

# ===================================================================
# 0. 환경설정 - EMS 화면이 바뀌면 이 구역만 수정하면 됩니다.
# ===================================================================

$EmsWindowTitleContains    = "Search Error Code"                 # 메인 창(장비명/에러코드 조회 화면)
$PopupWindowTitleContains  = "Search and Select List of Values"  # 값 선택 팝업 창

# ControlType 은 System.Windows.Automation.ControlType 의 정적 필드명을 문자열로 적으면 됩니다.
$Config = @{
    EquipmentNameBox   = @{ AutomationId = "SEmNo";                          Name = "TODO";                     ControlType = "Edit" }
    AlarmCodeBox       = @{ AutomationId = "SPlcErrCode";                    Name = "TODO";                     ControlType = "Edit" }
    # 알람코드 입력 후 자동으로 채워지는 설비 에러명 표시란. 이 값이 실제로
    # 채워졌는지 확인한 뒤에 찾아보기를 눌러야 타이밍 충돌이 없음(고정 딜레이 대신
    # 이 값이 채워질 때까지 폴링).
    AutoFilledErrDescBox = @{ AutomationId = "SPlcErrDesc";                  Name = "TODO";                     ControlType = "Edit" }
    # "찾아보기" 링크. SPlcErrDesc 값이 채워지면 이 링크의 Name 자체가 바뀌는
    # 것으로 보여, Name은 더 이상 매칭에 안 쓰고 참고용으로만 남겨둠. 실제로는
    # ControlType(Hyperlink)만으로 후보를 모은 뒤 알람코드 입력창과 위치가
    # 가장 가까운 것을 찾음(Wait-ForNearestLookupLink 참고).
    LookupLink         = @{ AutomationId = "TODO";                           Name = "Search: 설비 에러 코드";   ControlType = "Hyperlink" }
    PopupRadioItem     = @{ AutomationId = "TODO";                           Name = "Select";                   ControlType = "RadioButton" }
    PopupConfirmButton = @{ AutomationId = "TODO";                           Name = "Select";                   ControlType = "Button" }
    SearchButton       = @{ AutomationId = "Find";                          Name = "조회";                     ControlType = "Button" }
    EditIconLink       = @{ AutomationId = "SearchAlarmCdTable:Update:0";    Name = "TODO";                     ControlType = "Hyperlink" }
    AlarmNameBox       = @{ AutomationId = "cPlcErrDesc";                    Name = "TODO";                     ControlType = "Edit" }
    StatusBox          = @{ AutomationId = "NStatus";                       Name = "TODO";                     ControlType = "Edit" }
    GenerateButton     = @{ AutomationId = "NGenerate";                     Name = "생성";                     ControlType = "Button" }
    # "생성" 클릭 후 표준 에러 설명이 표시되는 곳(span id=NStandardErrorDesc). 이
    # 안에 새_알람명이 들어있으면 성공으로 판단. 실제 UIA ControlType이 Text가
    # 아닐 수도 있어(예: Group/Pane) ControlType 조건은 걸지 않고 AutomationId로만 찾음.
    GeneratedTextDisplay = @{ AutomationId = "NStandardErrorDesc";          Name = "TODO";                     ControlType = "TODO" }
    SaveButton         = @{ AutomationId = "SaveButton";                    Name = "저장";                     ControlType = "Button" }
    ErrorListLink      = @{ AutomationId = "XXEMSSTD052";                   Name = "Error List";               ControlType = "Hyperlink" }
}

$TimeoutSec             = 10   # 요소/화면전환 대기 타임아웃(초)
$GenerateVerifyTimeoutSec = 5  # "생성" 후 확인 텍스트 대기 타임아웃(초)
$PollingIntervalMs       = 300 # 폴링 간격(ms)

# ===================================================================
# 경로 설정
# ===================================================================
$ScriptDir     = Split-Path -Parent $MyInvocation.MyCommand.Path
$DataCsvPath   = Join-Path $ScriptDir "data.csv"
$ResultCsvPath = Join-Path $ScriptDir "result.csv"
$StopFlagPath  = Join-Path $ScriptDir "STOP.txt"
$LogPath       = Join-Path $ScriptDir "run.log"

# ===================================================================
# 1. 어셈블리 로드 (Windows 기본 포함, 설치 불필요)
# ===================================================================
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Write-Log {
    param([string]$Message)
    $line = "[{0:yyyy-MM-dd HH:mm:ss}] {1}" -f (Get-Date), $Message
    Write-Host $line
    Add-Content -Path $LogPath -Value $line
}

# 진단용: 클릭 직전에 실제로 어떤 요소를 골랐는지(위치/이름/AutomationId 등)
# 로그에 남긴다. 손으로 확인한 요소와 같은 것인지 대조하는 용도.
function Write-ElementDebugInfo {
    param(
        [System.Windows.Automation.AutomationElement]$Element,
        [string]$Label
    )
    $ct = "?"; $name = ""; $autoId = ""; $className = ""; $rectStr = ""
    try { $ct = ($Element.Current.ControlType.ProgrammaticName -replace '^ControlType\.', '') } catch { }
    try { $name = $Element.Current.Name } catch { }
    try { $autoId = $Element.Current.AutomationId } catch { }
    try { $className = $Element.Current.ClassName } catch { }
    try {
        $r = $Element.Current.BoundingRectangle
        $rectStr = "{0:F0},{1:F0},{2:F0},{3:F0}" -f $r.X, $r.Y, $r.Width, $r.Height
    } catch { }
    Write-Log "[진단:$Label] ControlType=$ct Name='$name' AutomationId='$autoId' ClassName='$className' BoundingRect(X,Y,W,H)=$rectStr"
}

# ===================================================================
# 2. 조건(Condition) / 패턴 헬퍼
# ===================================================================

function Get-ControlTypeByName {
    param([string]$Name)
    $field = [System.Windows.Automation.ControlType].GetField($Name, [System.Reflection.BindingFlags]::Public -bor [System.Reflection.BindingFlags]::Static)
    if (-not $field) { throw "알 수 없는 ControlType 이름: $Name" }
    return $field.GetValue($null)
}

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

    if ($conditions.Count -eq 0) { throw "요소 조건이 비어있습니다. `$Config 값을 채워주세요." }
    if ($conditions.Count -eq 1) { return $conditions[0] }
    return New-Object System.Windows.Automation.AndCondition($conditions.ToArray())
}

# [System.Windows.Automation.LegacyIAccessiblePattern] 처럼 대괄호로 타입을 직접 쓰면,
# 그 타입이 이 PC의 .NET 환경에 없을 때 try/catch로도 못 막는 에러가 나는 것이 실제
# 확인됨. 리플렉션으로 "있으면 쓰고 없으면 조용히 실패"하게 우회.
function Get-PatternObjectSafe {
    param([string]$PatternClassName)
    try {
        $asm = [System.Windows.Automation.AutomationElement].Assembly
        $type = $asm.GetType("System.Windows.Automation.$PatternClassName")
        if (-not $type) { return $null }
        $field = $type.GetField("Pattern", [System.Reflection.BindingFlags]::Public -bor [System.Reflection.BindingFlags]::Static)
        if (-not $field) { return $null }
        return $field.GetValue($null)
    } catch {
        return $null
    }
}

# ===================================================================
# 3. 창 찾기 - 메인 창은 PID 고정(엣지 프로세스만 대상), 팝업은 뜰 때마다 새로 찾음
# ===================================================================

function Find-EmsWindow {
    param([string]$TitleContains)

    $edgeProcessIds = @(Get-Process -Name "msedge" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    if ($edgeProcessIds.Count -eq 0) {
        throw "실행 중인 Microsoft Edge(msedge.exe) 프로세스를 찾지 못했습니다. Edge에서 EMS 페이지를 열어두었는지 확인하세요."
    }

    $root = [System.Windows.Automation.AutomationElement]::RootElement
    $windowCondition = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
        [System.Windows.Automation.ControlType]::Window)
    $candidates = $root.FindAll([System.Windows.Automation.TreeScope]::Children, $windowCondition)

    $matches = @()
    foreach ($w in $candidates) {
        try {
            if (($w.Current.ProcessId -in $edgeProcessIds) -and ($w.Current.Name -like "*$TitleContains*")) {
                $matches += $w
            }
        } catch { }
    }

    if ($matches.Count -eq 0) {
        throw "실행 중인 엣지 창 중에서 제목에 '$TitleContains' 를 포함하는 창을 찾지 못했습니다."
    }
    if ($matches.Count -gt 1) {
        Write-Log "경고: 제목이 일치하는 엣지 창이 $($matches.Count)개 발견됨. 첫 번째 창을 사용합니다."
    }
    return $matches[0]
}

# 팝업(값 선택 창)이 뜰 때까지 폴링 대기 - 메인 창과 별개의 새 창이므로 매번 새로 탐색
function Wait-ForPopupWindow {
    param([string]$TitleContains, [int]$TimeoutSec = 10)

    $edgeProcessIds = @(Get-Process -Name "msedge" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $root = [System.Windows.Automation.AutomationElement]::RootElement
    $windowCondition = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
        [System.Windows.Automation.ControlType]::Window)

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while ($sw.Elapsed.TotalSeconds -lt $TimeoutSec) {
        $candidates = $root.FindAll([System.Windows.Automation.TreeScope]::Children, $windowCondition)
        foreach ($w in $candidates) {
            try {
                if (($w.Current.ProcessId -in $edgeProcessIds) -and ($w.Current.Name -like "*$TitleContains*")) {
                    return $w
                }
            } catch { }
        }
        Start-Sleep -Milliseconds $PollingIntervalMs
    }
    return $null
}

# 팝업 창이 닫혔는지 확인(스테일 요소 접근 시 예외가 나는 것을 이용)
function Wait-ForWindowClosed {
    param([System.Windows.Automation.AutomationElement]$Window, [int]$TimeoutSec = 10)
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while ($sw.Elapsed.TotalSeconds -lt $TimeoutSec) {
        try {
            $null = $Window.Current.ProcessId
        } catch {
            return $true
        }
        Start-Sleep -Milliseconds $PollingIntervalMs
    }
    return $false
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

# "찾아보기" 링크 전용: Name 텍스트에 의존하지 않는다(값이 채워지면 그 안의
# 내용을 반영해 이름 자체가 바뀌는 것으로 보여, 이름 일치/포함 방식 둘 다
# 불안정했음). 대신 ControlType(Hyperlink)만으로 후보를 모은 뒤, 기준 요소
# (알람코드 입력창)와 화면상 위치(X,Y 모두)가 가장 가까운 것을 고른다.
# 화면 갱신 타이밍에 걸리는 경우를 대비해 타임아웃까지 폴링 재시도한다.
function Wait-ForNearestLookupLink {
    param(
        [System.Windows.Automation.AutomationElement]$Parent,
        [string]$ControlTypeName,
        [System.Windows.Automation.AutomationElement]$ReferenceElement,
        [int]$TimeoutSec = 10
    )
    $ct = Get-ControlTypeByName $ControlTypeName
    $ctCondition = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ControlTypeProperty, $ct)
    $refRect = $ReferenceElement.Current.BoundingRectangle
    $refX = $refRect.X + ($refRect.Width / 2)
    $refY = $refRect.Y + ($refRect.Height / 2)

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while ($sw.Elapsed.TotalSeconds -lt $TimeoutSec) {
        $candidates = @($Parent.FindAll([System.Windows.Automation.TreeScope]::Descendants, $ctCondition))

        if ($candidates.Count -gt 0) {
            $best = $null
            $bestDist = [double]::MaxValue
            foreach ($c in $candidates) {
                try {
                    $r = $c.Current.BoundingRectangle
                    $cx = $r.X + ($r.Width / 2)
                    $cy = $r.Y + ($r.Height / 2)
                    $dist = [math]::Sqrt([math]::Pow($cx - $refX, 2) + [math]::Pow($cy - $refY, 2))
                    if ($dist -lt $bestDist) {
                        $bestDist = $dist
                        $best = $c
                    }
                } catch { }
            }
            if ($best) { return $best }
        }
        Start-Sleep -Milliseconds $PollingIntervalMs
    }
    return $null
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

# "생성" 버튼 클릭 후, 특정 요소(조건으로 지정)의 표시값에 $Substring 이
# 포함될 때까지 대기 (표준 에러 설명란(NStandardErrorDesc)에 새 알람명이
# 반영됐는지 확인하는 용도)
function Wait-ForElementTextContaining {
    param(
        [System.Windows.Automation.AutomationElement]$Parent,
        [System.Windows.Automation.Condition]$Condition,
        [string]$Substring,
        [int]$TimeoutSec = 10
    )
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while ($sw.Elapsed.TotalSeconds -lt $TimeoutSec) {
        $el = Find-ElementNow -Parent $Parent -Condition $Condition
        if ($el) {
            try {
                $val = Get-ElementDisplayValue -Element $el
                if ($val -and $val.Contains($Substring)) { return $true }
            } catch { }
        }
        Start-Sleep -Milliseconds $PollingIntervalMs
    }
    return $false
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

# 알람코드 입력 후 자동으로 채워지는 필드(예: 설비 에러명 표시란)에 실제로
# 값이 들어왔는지 확인될 때까지 폴링 대기. 고정 딜레이 대신 사용.
function Wait-ForElementValueNonEmpty {
    param(
        [System.Windows.Automation.AutomationElement]$Parent,
        [System.Windows.Automation.Condition]$Condition,
        [int]$TimeoutSec = 10
    )
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while ($sw.Elapsed.TotalSeconds -lt $TimeoutSec) {
        $el = Find-ElementNow -Parent $Parent -Condition $Condition
        if ($el) {
            try {
                $val = Get-ElementDisplayValue -Element $el
                if ($val -and $val.Trim() -ne "") { return $true }
            } catch { }
        }
        Start-Sleep -Milliseconds $PollingIntervalMs
    }
    return $false
}

# ===================================================================
# 5. 값 입력 / 클릭 / 선택
# ===================================================================

function Set-UiaValue {
    param(
        [System.Windows.Automation.AutomationElement]$Element,
        [string]$Value
    )
    $pattern = $Element.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern)
    $pattern.SetValue($Value)

    Start-Sleep -Milliseconds 150
    $actual = $Element.GetCurrentPropertyValue([System.Windows.Automation.ValuePattern]::ValueProperty)
    if ($actual -ne $Value) {
        throw "값 입력 검증 실패 (입력='$Value', 실제='$actual')"
    }
}

function Invoke-UiaClick {
    param([System.Windows.Automation.AutomationElement]$Element)

    $invokePatternObj = Get-PatternObjectSafe -PatternClassName "InvokePattern"
    if ($invokePatternObj) {
        $invokePattern = $null
        if ($Element.TryGetCurrentPattern($invokePatternObj, [ref]$invokePattern)) {
            Write-Log "[Invoke-UiaClick] InvokePattern.Invoke() 호출 시도"
            $invokePattern.Invoke()
            Write-Log "[Invoke-UiaClick] InvokePattern.Invoke() 호출 완료(예외 없음)"
            return
        } else {
            Write-Log "[Invoke-UiaClick] InvokePattern 타입은 있으나 이 요소는 미지원(TryGetCurrentPattern=false)"
        }
    } else {
        Write-Log "[Invoke-UiaClick] InvokePattern 타입 자체를 찾지 못함"
    }

    # InvokePattern 미지원 요소에 대한 대안: LegacyIAccessiblePattern.DoDefaultAction()
    $legacyPatternObj = Get-PatternObjectSafe -PatternClassName "LegacyIAccessiblePattern"
    if ($legacyPatternObj) {
        $legacyPattern = $null
        if ($Element.TryGetCurrentPattern($legacyPatternObj, [ref]$legacyPattern)) {
            Write-Log "[Invoke-UiaClick] LegacyIAccessiblePattern.DoDefaultAction() 호출 시도"
            $legacyPattern.DoDefaultAction()
            Write-Log "[Invoke-UiaClick] LegacyIAccessiblePattern.DoDefaultAction() 호출 완료(예외 없음)"
            return
        } else {
            Write-Log "[Invoke-UiaClick] LegacyIAccessiblePattern 타입은 있으나 이 요소는 미지원(TryGetCurrentPattern=false)"
        }
    } else {
        Write-Log "[Invoke-UiaClick] LegacyIAccessiblePattern 타입 자체를 찾지 못함"
    }

    throw "이 요소는 InvokePattern과 LegacyIAccessiblePattern을 모두 지원하지 않습니다."
}

# 라디오 버튼 선택 - 원래 동작(SelectionItemPattern.Select())을 우선 시도하고,
# 미지원이면 클릭으로 대체
function Select-UiaRadioButton {
    param([System.Windows.Automation.AutomationElement]$Element)

    $selPatternObj = Get-PatternObjectSafe -PatternClassName "SelectionItemPattern"
    if ($selPatternObj) {
        $sp = $null
        if ($Element.TryGetCurrentPattern($selPatternObj, [ref]$sp)) {
            $sp.Select()
            return
        }
    }
    Invoke-UiaClick -Element $Element
}

# ===================================================================
# 6. 안전 정지(STOP.txt) 확인
# ===================================================================

function Test-StopRequested {
    return (Test-Path -LiteralPath $StopFlagPath)
}

# ===================================================================
# 7. 조회 + 팝업 선택 + 조회 + 연필클릭 -> Update Error Code 화면 진입
#    (수정할 때, 재검증할 때 둘 다 이 루틴을 그대로 반복 사용)
# ===================================================================

function Invoke-SearchAndOpenUpdateScreen {
    param(
        [System.Windows.Automation.AutomationElement]$MainWindow,
        [string]$EquipmentName,
        [string]$AlarmCode
    )

    # 1) 장비명 + 알람코드 입력 (Search Error Code 화면)
    $cond = New-ConditionFromConfig $Config.EquipmentNameBox
    $el = Wait-UIAElement -Parent $MainWindow -Condition $cond -TimeoutSec $TimeoutSec
    if (-not $el) { throw "장비명 입력창(SEmNo)을 찾지 못했습니다." }
    Set-UiaValue -Element $el -Value $EquipmentName

    $cond = New-ConditionFromConfig $Config.AlarmCodeBox
    $alarmCodeEl = Find-ElementNow -Parent $MainWindow -Condition $cond
    if (-not $alarmCodeEl) { throw "알람코드 입력창(SPlcErrCode)을 찾지 못했습니다." }
    Set-UiaValue -Element $alarmCodeEl -Value $AlarmCode

    # 알람코드 입력 시 자동으로 채워지는 설비 에러명(SPlcErrDesc)에 실제 값이
    # 들어올 때까지 대기 (고정 딜레이 대신, 실제 채워졌는지 확인 후 다음 단계 진행)
    $cond = New-ConditionFromConfig $Config.AutoFilledErrDescBox
    $filled = Wait-ForElementValueNonEmpty -Parent $MainWindow -Condition $cond -TimeoutSec $TimeoutSec
    if (-not $filled) { throw "설비 에러명 자동 입력(SPlcErrDesc)이 채워지는 것을 확인하지 못했습니다." }

    # 2) 찾아보기(Hyperlink) 클릭 -> 값 선택 팝업 대기
    #    같은 이름 계열의 "찾아보기" 링크가 화면에 여러 개(다른 필드용) 있고,
    #    정확한 문구도 입력 시점에 따라 조금 달라질 수 있어 "포함" 여부로 찾고,
    #    그 중 알람코드 입력창과 세로 위치가 가장 가까운 것을 찾아 클릭한다.
    $el = Wait-ForNearestLookupLink -Parent $MainWindow -ControlTypeName $Config.LookupLink.ControlType `
        -ReferenceElement $alarmCodeEl -TimeoutSec $TimeoutSec
    if (-not $el) { throw "알람코드 입력창 근처에서 찾아보기 링크(Hyperlink)를 찾지 못했습니다." }
    Write-ElementDebugInfo -Element $el -Label "찾아보기 링크(클릭 대상)"
    Invoke-UiaClick -Element $el

    $popup = Wait-ForPopupWindow -TitleContains $PopupWindowTitleContains -TimeoutSec $TimeoutSec
    if (-not $popup) { throw "값 선택 팝업(Search and Select List of Values)이 뜨지 않았습니다." }

    # 3) 팝업에서 첫 번째 라디오 선택 (정확한 코드로 검색했으므로 항상 첫 번째가 정답)
    $cond = New-ConditionFromConfig $Config.PopupRadioItem
    $radioEl = Wait-UIAElement -Parent $popup -Condition $cond -TimeoutSec $TimeoutSec
    if (-not $radioEl) { throw "팝업에서 선택할 라디오 버튼을 찾지 못했습니다." }
    Select-UiaRadioButton -Element $radioEl

    # 4) 팝업의 Select 버튼으로 확정 -> 팝업이 닫힐 때까지 대기
    $cond = New-ConditionFromConfig $Config.PopupConfirmButton
    $confirmEl = Find-ElementNow -Parent $popup -Condition $cond
    if (-not $confirmEl) { throw "팝업의 Select 확정 버튼을 찾지 못했습니다." }
    Invoke-UiaClick -Element $confirmEl
    Wait-ForWindowClosed -Window $popup -TimeoutSec $TimeoutSec | Out-Null

    # 5) 조회 버튼 클릭
    $cond = New-ConditionFromConfig $Config.SearchButton
    $el = Wait-UIAElement -Parent $MainWindow -Condition $cond -TimeoutSec $TimeoutSec
    if (-not $el) { throw "조회 버튼을 찾지 못했습니다." }
    Invoke-UiaClick -Element $el

    # 6) 연필 모양 아이콘(Hyperlink, InvokePattern 지원 확인됨) 클릭
    #    -> 클릭하면 같은 창 안에서 "Update Error Code" 화면으로 전환됨(새 창 아님)
    $cond = New-ConditionFromConfig $Config.EditIconLink
    $el = Wait-UIAElement -Parent $MainWindow -Condition $cond -TimeoutSec $TimeoutSec
    if (-not $el) { throw "연필 모양 수정 아이콘을 찾지 못했습니다." }
    Invoke-UiaClick -Element $el

    # 7) Update Error Code 화면 전환 대기 (고정 딜레이 대신, 에러명 입력창이
    #    나타나는 것으로 화면 전환 완료를 판단)
    $cond = New-ConditionFromConfig $Config.AlarmNameBox
    $alarmNameEl = Wait-UIAElement -Parent $MainWindow -Condition $cond -TimeoutSec $TimeoutSec
    if (-not $alarmNameEl) { throw "Update Error Code 화면(에러명 입력창)으로 전환되지 않았습니다." }

    return $alarmNameEl
}

# ===================================================================
# 8. Update Error Code 화면에서 값 입력 + 생성 + 검증 + 저장
# ===================================================================

function Set-AlarmNameAndSave {
    param(
        [System.Windows.Automation.AutomationElement]$MainWindow,
        [System.Windows.Automation.AutomationElement]$AlarmNameEl,
        [string]$NewAlarmName
    )

    # 에러명 입력
    Set-UiaValue -Element $AlarmNameEl -Value $NewAlarmName

    # 상태 입력란에도 동일한 값 입력 (실사용 환경에서 확인된 사양)
    $cond = New-ConditionFromConfig $Config.StatusBox
    $statusEl = Find-ElementNow -Parent $MainWindow -Condition $cond
    if (-not $statusEl) { throw "상태 입력창(NStatus)을 찾지 못했습니다." }
    Set-UiaValue -Element $statusEl -Value $NewAlarmName

    # 생성 버튼 클릭
    $cond = New-ConditionFromConfig $Config.GenerateButton
    $genEl = Find-ElementNow -Parent $MainWindow -Condition $cond
    if (-not $genEl) { throw "생성 버튼을 찾지 못했습니다." }
    Invoke-UiaClick -Element $genEl

    # 저장 전 확인: 표준 에러 설명란(NStandardErrorDesc)에 새 알람명이 포함되는지 확인
    $cond = New-ConditionFromConfig $Config.GeneratedTextDisplay
    $found = Wait-ForElementTextContaining -Parent $MainWindow -Condition $cond -Substring $NewAlarmName -TimeoutSec $GenerateVerifyTimeoutSec
    if (-not $found) {
        throw "'생성' 후 표준 에러 설명란(NStandardErrorDesc)에서 '$NewAlarmName' 을 찾지 못했습니다. 저장하지 않고 실패 처리합니다."
    }

    # 저장 버튼 클릭
    $cond = New-ConditionFromConfig $Config.SaveButton
    $saveEl = Find-ElementNow -Parent $MainWindow -Condition $cond
    if (-not $saveEl) { throw "저장 버튼을 찾지 못했습니다." }
    Invoke-UiaClick -Element $saveEl

    # 저장 결과를 알려주는 별도 요소가 없어(재조회로 검증) 최소한의 처리 대기만 둠
    Start-Sleep -Milliseconds 1000
}

# Error List 클릭 -> Search Error Code 화면으로 복귀 (다음 조회를 위한 상태 복원)
function Return-ToSearchScreen {
    param([System.Windows.Automation.AutomationElement]$MainWindow)

    $cond = New-ConditionFromConfig $Config.ErrorListLink
    $el = Find-ElementNow -Parent $MainWindow -Condition $cond
    if (-not $el) { throw "Error List 링크를 찾지 못했습니다." }
    Invoke-UiaClick -Element $el

    $cond = New-ConditionFromConfig $Config.EquipmentNameBox
    $el = Wait-UIAElement -Parent $MainWindow -Condition $cond -TimeoutSec $TimeoutSec
    if (-not $el) { throw "Search Error Code 화면으로 돌아오지 못했습니다." }
}

# ===================================================================
# 9. 한 행 처리: 수정 -> 복귀 -> 재검증(재조회) -> 복귀
# ===================================================================

function Process-Row {
    param(
        [System.Windows.Automation.AutomationElement]$MainWindow,
        [pscustomobject]$Row,
        [int]$Index,
        [int]$Total
    )

    Write-Log ("{0}/{1} 처리 중: {2}, {3}" -f $Index, $Total, $Row.장비명, $Row.알람코드)
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    $result = [pscustomobject]@{
        장비명     = $Row.장비명
        알람코드   = $Row.알람코드
        새_알람명 = $Row.새_알람명
        검증완료   = "N"
        소요시간초 = 0
        처리결과   = "실패"
        실패사유   = ""
    }

    try {
        # 1) 수정 단계
        $alarmNameEl = Invoke-SearchAndOpenUpdateScreen -MainWindow $MainWindow -EquipmentName $Row.장비명 -AlarmCode $Row.알람코드
        Set-AlarmNameAndSave -MainWindow $MainWindow -AlarmNameEl $alarmNameEl -NewAlarmName $Row.새_알람명
        Return-ToSearchScreen -MainWindow $MainWindow

        # 2) 재검증 단계: 처음부터 다시 조회해서 실제 반영된 값을 확인
        $verifyEl = Invoke-SearchAndOpenUpdateScreen -MainWindow $MainWindow -EquipmentName $Row.장비명 -AlarmCode $Row.알람코드
        $displayedName = Get-ElementDisplayValue -Element $verifyEl
        Return-ToSearchScreen -MainWindow $MainWindow

        if ($displayedName -ne $Row.새_알람명) {
            throw "저장 후 재검증 실패: 화면 값='$displayedName', 기대값='$($Row.새_알람명)'"
        }

        $result.검증완료 = "Y"
        $result.처리결과 = "성공"
        Write-Log "성공: 장비명='$($Row.장비명)' 알람코드='$($Row.알람코드)' -> '$($Row.새_알람명)' (재검증 통과)"
    }
    catch {
        $result.처리결과 = "실패"
        $result.실패사유 = $_.Exception.Message
        Write-Log "실패: 장비명='$($Row.장비명)' 알람코드='$($Row.알람코드)' 사유='$($_.Exception.Message)'"

        # 다음 행이 정상적으로 시작할 수 있도록, 실패했더라도 Search Error Code
        # 화면으로 복귀를 한 번 시도(이미 복귀돼 있으면 조용히 무시됨)
        try { Return-ToSearchScreen -MainWindow $MainWindow } catch { }
    }
    finally {
        $stopwatch.Stop()
        $result.소요시간초 = [math]::Round($stopwatch.Elapsed.TotalSeconds, 1)
    }

    return $result
}

# ===================================================================
# 10. 실시간 결과 창 (GUI) - 일시정지 버튼 없이, 진행 상황만 실시간으로 표시
# ===================================================================

function New-ResultsWindow {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "EMS 알람 자동화 - 실시간 처리 현황"
    $form.Width = 950
    $form.Height = 500
    $form.StartPosition = "CenterScreen"

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Dock = "Fill"
    $grid.ReadOnly = $true
    $grid.AllowUserToAddRows = $false
    $grid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill

    [void]$grid.Columns.Add("장비명", "장비명")
    [void]$grid.Columns.Add("알람코드", "알람코드")
    [void]$grid.Columns.Add("새알람명", "새_알람명")
    [void]$grid.Columns.Add("검증완료", "검증완료")
    [void]$grid.Columns.Add("소요시간", "소요시간(초)")
    [void]$grid.Columns.Add("처리결과", "처리결과")
    [void]$grid.Columns.Add("실패사유", "실패사유")

    $form.Controls.Add($grid)
    $form.Tag = $grid
    return $form
}

function Add-ResultToWindow {
    param(
        [System.Windows.Forms.Form]$ResultsForm,
        [pscustomobject]$Row
    )
    $grid = $ResultsForm.Tag
    $rowIndex = $grid.Rows.Add()
    $r = $grid.Rows[$rowIndex]
    $r.Cells["장비명"].Value   = $Row.장비명
    $r.Cells["알람코드"].Value = $Row.알람코드
    $r.Cells["새알람명"].Value = $Row.새_알람명
    $r.Cells["검증완료"].Value = $Row.검증완료
    $r.Cells["소요시간"].Value = $Row.소요시간초
    $r.Cells["처리결과"].Value = $Row.처리결과
    $r.Cells["실패사유"].Value = $Row.실패사유
    $grid.FirstDisplayedScrollingRowIndex = $rowIndex
}

# ===================================================================
# 11. 메인 실행부
# ===================================================================

Write-Log "===== EMS 알람 자동화 시작 ====="

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
Write-Log "대상 창 확보. PID=$($window.Current.ProcessId), 제목='$($window.Current.Name)'"

# 결과 파일은 이번 실행 기준으로 새로 시작 (행마다 즉시 추가 저장됨)
if (Test-Path -LiteralPath $ResultCsvPath) {
    Remove-Item -LiteralPath $ResultCsvPath -Force
}

$resultsForm = New-ResultsWindow
$resultsForm.Show()
[System.Windows.Forms.Application]::DoEvents()

$results = New-Object System.Collections.Generic.List[pscustomobject]
$total = $rows.Count
$index = 0

foreach ($row in $rows) {
    $index++

    if (Test-StopRequested) {
        Write-Log "STOP.txt 감지됨. 현재까지 결과를 저장하고 종료합니다."
        break
    }

    $r = Process-Row -MainWindow $window -Row $row -Index $index -Total $total
    $results.Add($r)

    # 파일에 즉시 추가 저장
    $r | Export-Csv -Path $ResultCsvPath -NoTypeInformation -Encoding UTF8 -Append

    # 실시간 결과 창에도 즉시 반영
    Add-ResultToWindow -ResultsForm $resultsForm -Row $r
    [System.Windows.Forms.Application]::DoEvents()
}

$successCount = ($results | Where-Object { $_.처리결과 -eq "성공" }).Count
$failCount = ($results | Where-Object { $_.처리결과 -eq "실패" }).Count
Write-Log "===== 종료. 성공 $successCount 건 / 실패 $failCount 건. 결과: $ResultCsvPath ====="

$resultsForm.Text = "$($resultsForm.Text) - 완료 (성공 $successCount / 실패 $failCount)"

# 결과 창은 사용자가 직접 닫을 때까지 화면에 유지
[System.Windows.Forms.Application]::Run($resultsForm)
