<#
    고장 목록(KPI) 현상/원인/조치 자동 작성 스크립트
    ---------------------------------------------
    - Windows PowerShell 5.1(powershell.exe) 기준. Update-EmsAlarms.ps1과 같은 원칙.
    - .NET UI Automation(UIA) 만 사용. 추가 모듈/패키지 설치, 관리자 권한 불필요.
    - 좌표클릭 / SendKeys / Tab 이동 사용하지 않음. 전부 UIA 패턴 기반.
    - 대상 프로그램은 C#으로 만들어진 사내 프로그램(WinForms 버튼으로 확인됨)이며,
      Update-EmsAlarms.ps1에서 검증된 도구 함수(요소 찾기/클릭/정지 키/결과창)를
      그대로 재사용함.

    [현재까지 확인된 화면 흐름 - 세부 AutomationId/ControlType은 미확인(TODO)]
    1) 메인 화면에서 "KPI" 버튼 클릭 -> 한 달 치 고장 목록이 일자별로 정리된
       표(그리드)가 나타남. 각 날짜 칸에는 그 날의 고장 총합이 표시됨.
    2) 그 날짜(고장 총합) 칸을 더블클릭 -> 해당 날짜의 고장 목록(표)이 열림.
       이 표의 각 행은 "장비명 / 알람명" 등의 컬럼을 가짐.
       (작성된 고장=검정색 글자, 미작성 고장=노란/갈색 글자로 화면에 구분되지만,
        UIA는 기본적으로 글자 색을 읽는 표준 속성이 없어서 이걸로 판단하지 않음.
        대신 실제로 그 행을 열어서 텍스트박스 안 내용이 비어있는지로 판단함.)
    3) 그 행의 "알람명" 칸을 더블클릭 -> 고장 입력 화면이 열리고, 하나의 큰
       텍스트박스가 보임. 그 박스 안에 "현상/원인/조치"가 라벨로 구분되어 있음
       (정확한 서식은 회사 보안상 미리 알 수 없어 실행 시점에 그 자리에서
        라벨 위치를 찾아 내용을 끼워넣는 방식으로 처리함. Set-FaultTextByLabels 참고).
    4) 저장 버튼 클릭 -> 그 행에서 나가서 다시 들어가 재조회 -> 텍스트박스에
       입력한 현상/원인/조치가 실제로 들어갔는지 검증.

    [아직 회사에서 실제 화면 보고 확인해야 하는 것 - 모두 TODO로 표시됨]
    - 메인 창을 찾을 프로세스명/창 제목
    - KPI 버튼의 AutomationId/Name
    - 일자별 고장 총합이 표시되는 그리드의 구조(DataGridView인지, 어떤 컬럼인지)
    - "더블클릭"이 실제로 UIA InvokePattern.Invoke() 두 번으로 흉내낼 수 있는지,
      아니면 다른 방식(LegacyIAccessible DoDefaultAction 등)이 필요한지
      (WinForms의 그리드 셀 더블클릭 이벤트는 버튼 클릭과 다르게 동작할 수 있어
       반드시 실제 화면에서 검증 필요 - Invoke-UiaDoubleClick 주석 참고)
    - 그 날짜별 고장 목록 표의 컬럼 구성(장비명/알람명이 몇 번째 컬럼인지)
    - 큰 텍스트박스의 AutomationId/ControlType, 저장 버튼의 AutomationId

    [CSV 컬럼] 장비명, 알람코드, 현상, 원인, 조치
    (날짜 컬럼 없음 - 스크립트가 KPI 화면에 보이는 전체 날짜를 다 훑으면서,
     이 CSV에 있는 장비+알람코드 조합을 찾아 미작성 상태면 채워 넣음)
#>

[CmdletBinding()]
param(
    # 테스트 모드: data.csv 전체가 아니라 앞의 $TestModeRows 건만 처리
    [switch]$TestMode,
    [int]$TestModeRows = 2,

    # 정지 키(기본 F8). Update-EmsAlarms.ps1과 동일한 방식.
    #   F7=0x76  F8=0x77  F9=0x78  F10=0x79  F11=0x7A  F12=0x7B  ScrollLock=0x91  Pause=0x13
    [int]$StopKeyCode = 0x77,
    [string]$StopKeyName = "F8"
)

# 파일이 최신 버전인지 헷갈리지 않도록, 실행할 때마다 콘솔/로그에 이 값을 표시함.
$ScriptVersion = "2026-09-18-A (뼈대/스켈레톤 - 실제 화면 확인 전)"

if ($PSVersionTable.PSEdition -ne 'Desktop') {
    Write-Warning "이 스크립트는 Windows PowerShell 5.1(powershell.exe) 기준으로 검증되었습니다. 현재 PSEdition='$($PSVersionTable.PSEdition)' 입니다."
}

# ===================================================================
# 0. 환경설정 - 실제 화면 확인 후 이 구역만 수정하면 됩니다 (전부 TODO).
# ===================================================================

# 대상 프로그램을 찾을 때 쓰는 프로세스명(확장자 .exe 제외)과 창 제목의 일부.
# Get-EmsUiaTree.ps1 실행 시 나오는 창 목록에서 "프로세스=" 뒤에 나오는 이름을
# 그대로 적으면 됩니다.
$AppProcessNameContains  = "TODO"   # 예: "FaultReportApp"
$AppWindowTitleContains  = "TODO"   # 예: 메인 창 제목의 일부

$Config = @{
    # 메인 화면의 "KPI" 버튼
    KpiButton         = @{ AutomationId = "TODO"; Name = "KPI";  ControlType = "Button" }

    # 일자별 고장 총합이 표시되는 그리드(표) 컨테이너. WinForms DataGridView라면
    # ControlType이 "DataGrid" 또는 "Table"로 잡히는 경우가 많습니다.
    DaySummaryGrid    = @{ AutomationId = "TODO"; Name = "TODO"; ControlType = "TODO" }

    # 특정 날짜를 더블클릭해서 들어간 뒤 보이는, 그 날의 고장 목록 그리드
    FaultListGrid     = @{ AutomationId = "TODO"; Name = "TODO"; ControlType = "TODO" }

    # 고장 입력 화면의 큰 텍스트박스(현상/원인/조치가 라벨로 구분되어 있는 곳)
    FaultTextBox      = @{ AutomationId = "TODO"; Name = "TODO"; ControlType = "Edit" }

    # 고장 입력 화면의 저장 버튼
    SaveButton        = @{ AutomationId = "TODO"; Name = "TODO"; ControlType = "Button" }

    # 고장 입력 화면에서 목록으로 되돌아가는 버튼/링크(뒤로가기, 목록, 닫기 등)
    BackToListButton  = @{ AutomationId = "TODO"; Name = "TODO"; ControlType = "TODO" }
}

# 현상/원인/조치 라벨을 텍스트박스 안에서 찾을 때 쓰는 정규식.
# 회사 보안상 정확한 서식을 미리 알 수 없어, 라벨 뒤에 콜론이 있든 없든,
# 같은 줄이든 다음 줄이든 웬만큼 다 걸리도록 널널하게 잡아둔 것입니다.
# 실제 화면 보고 이 부분만 조정하면 됩니다 (내용 자체는 공유 안 해도 됨,
# "라벨 뒤에 콜론이 있다/없다", "같은 줄이다/다음 줄이다" 같은 구조만 알려주면 됨).
$FaultLabelPatterns = @{
    현상 = '현상\s*:?\s*'
    원인 = '원인\s*:?\s*'
    조치 = '조치\s*:?\s*'
}

$TimeoutSec        = 10
$PollingIntervalMs = 300

# ===================================================================
# 경로 설정
# ===================================================================
$ScriptDir     = Split-Path -Parent $MyInvocation.MyCommand.Path
$DataCsvPath   = Join-Path $ScriptDir "fault_data.csv"
$ResultCsvPath = Join-Path $ScriptDir "fault_result.csv"
$StopFlagPath  = Join-Path $ScriptDir "STOP.txt"
$LogPath       = Join-Path $ScriptDir "fault_run.log"

# ===================================================================
# 1. 어셈블리 로드 (Windows 기본 포함, 설치 불필요)
# ===================================================================
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# 정지 키 감지용 (Update-EmsAlarms.ps1과 동일한 방식 - GetAsyncKeyState).
# 창 포커스나 Windows 메시지 큐에 의존하지 않아, 본체가 자동화로 바쁜 동안에도
# 키 입력을 놓치지 않습니다.
Add-Type -Namespace FaultReportAutomation -Name StopKeyNative -MemberDefinition @'
    [System.Runtime.InteropServices.DllImport("user32.dll")]
    public static extern short GetAsyncKeyState(int vKey);
'@

function Write-Log {
    param([string]$Message)
    $line = "[{0:yyyy-MM-dd HH:mm:ss}] {1}" -f (Get-Date), $Message
    Write-Host $line
    Add-Content -Path $LogPath -Value $line
}

# ===================================================================
# 2. 패턴 안전 조회 (Update-EmsAlarms.ps1과 동일)
# ===================================================================

# 패턴 클래스 이름을 문자열로 받아서, 리플렉션으로 "있으면 쓰고 없으면 조용히
# 실패"하게 함. 이 PC의 .NET 환경에 특정 패턴 타입이 없어도 스크립트가 안 죽음.
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

function Get-PatternSafe {
    param(
        [System.Windows.Automation.AutomationElement]$Element,
        [string]$PatternClassName
    )
    try {
        $patternObj = Get-PatternObjectSafe -PatternClassName $PatternClassName
        if (-not $patternObj) { return $null }
        $p = $null
        if ($Element.TryGetCurrentPattern($patternObj, [ref]$p)) { return $p }
        return $null
    } catch {
        return $null
    }
}

# ===================================================================
# 3. 창/요소 찾기 (Update-EmsAlarms.ps1과 동일한 원칙 - 좌표 미사용)
# ===================================================================

function Find-AppWindow {
    param([string]$ProcessNameContains, [string]$TitleContains, [int]$TimeoutSec = 15)

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while ($sw.Elapsed.TotalSeconds -lt $TimeoutSec) {
        $procIds = @(Get-Process -ErrorAction SilentlyContinue |
            Where-Object { $_.ProcessName -like "*$ProcessNameContains*" } |
            Select-Object -ExpandProperty Id)

        if ($procIds.Count -gt 0) {
            $root = [System.Windows.Automation.AutomationElement]::RootElement
            $windowCondition = New-Object System.Windows.Automation.PropertyCondition(
                [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
                [System.Windows.Automation.ControlType]::Window)
            $candidates = @($root.FindAll([System.Windows.Automation.TreeScope]::Children, $windowCondition))

            foreach ($w in $candidates) {
                try {
                    if (($w.Current.ProcessId -in $procIds) -and ($w.Current.Name -like "*$TitleContains*")) {
                        return $w
                    }
                } catch { }
            }
        }
        Start-Sleep -Milliseconds $PollingIntervalMs
    }
    return $null
}

function Test-ElementVisible {
    param([System.Windows.Automation.AutomationElement]$Element)
    try {
        if ($Element.Current.IsOffscreen) { return $false }
    } catch { return $false }
    try {
        $r = $Element.Current.BoundingRectangle
        if ($r.IsEmpty) { return $false }
        if ([double]::IsNaN($r.X) -or [double]::IsNaN($r.Y) -or [double]::IsInfinity($r.X) -or [double]::IsInfinity($r.Y)) { return $false }
        if ($r.Width -le 0 -or $r.Height -le 0) { return $false }
        return $true
    } catch {
        return $false
    }
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
    if ($Item.ClassName -and $Item.ClassName -ne "TODO") {
        $conditions.Add((New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::ClassNameProperty, $Item.ClassName)))
    }
    if ($Item.ControlType -and $Item.ControlType -ne "TODO") {
        $ctField = [System.Windows.Automation.ControlType].GetField($Item.ControlType, [System.Reflection.BindingFlags]::Public -bor [System.Reflection.BindingFlags]::Static)
        if ($ctField) {
            $ctValue = $ctField.GetValue($null)
            $conditions.Add((New-Object System.Windows.Automation.PropertyCondition(
                [System.Windows.Automation.AutomationElement]::ControlTypeProperty, $ctValue)))
        }
    }

    if ($conditions.Count -eq 0) {
        return (New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::AutomationIdProperty, ""))
    }
    if ($conditions.Count -eq 1) { return $conditions[0] }
    return (New-Object System.Windows.Automation.AndCondition($conditions.ToArray()))
}

function Find-ElementNow {
    param(
        [System.Windows.Automation.AutomationElement]$Parent,
        [System.Windows.Automation.Condition]$Condition,
        [System.Windows.Automation.TreeScope]$Scope = [System.Windows.Automation.TreeScope]::Descendants
    )
    return $Parent.FindFirst($Scope, $Condition)
}

function Find-VisibleElement {
    param(
        [System.Windows.Automation.AutomationElement]$Parent,
        [System.Windows.Automation.Condition]$Condition,
        [string]$Label = ""
    )
    $all = @()
    try {
        $all = @($Parent.FindAll([System.Windows.Automation.TreeScope]::Descendants, $Condition))
    } catch { }
    if ($all.Count -eq 0) { return $null }

    $visible = @($all | Where-Object { Test-ElementVisible -Element $_ })
    if ($visible.Count -gt 0) { return $visible[0] }

    Write-Log "[요소찾기:$Label] 후보 $($all.Count)개 있으나 전부 화면에 보이지 않음(IsOffscreen/좌표 문제). 첫 후보로 대체 진행."
    return $all[0]
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

# ===================================================================
# 4. 클릭 / 값 입력 (Update-EmsAlarms.ps1과 동일한 원칙 - 좌표/키 흉내내기 미사용)
# ===================================================================

function Invoke-UiaClick {
    param([System.Windows.Automation.AutomationElement]$Element)

    $invoke = Get-PatternSafe -Element $Element -PatternClassName "InvokePattern"
    if ($invoke) {
        try { $invoke.Invoke(); return $true } catch { Write-Log "[Invoke-UiaClick] Invoke 실패: $($_.Exception.Message)" }
    }
    $legacy = Get-PatternSafe -Element $Element -PatternClassName "LegacyIAccessiblePattern"
    if ($legacy) {
        try { $legacy.DoDefaultAction(); return $true } catch { Write-Log "[Invoke-UiaClick] DoDefaultAction 실패: $($_.Exception.Message)" }
    }
    Write-Log "[Invoke-UiaClick] Invoke/LegacyIAccessible 둘 다 지원 안 함 - 클릭 실패"
    return $false
}

# [확인 필요/가정] WinForms 그리드에서 "더블클릭으로 상세 열기"가 흔히 쓰이는데,
# UIA의 표준 클릭(InvokePattern.Invoke)은 보통 "한 번 클릭"에 대응하는 동작이라
# 더블클릭 이벤트에 연결된 코드가 그대로 실행될지는 실제 화면에서 검증이 필요함.
# 일단은 "짧은 간격으로 두 번 클릭"으로 흉내내되, 실제로 안 먹히면(다음 화면이
# 안 열리면) 다른 방식(LegacyIAccessible DoDefaultAction 등)을 시도하도록
# 되어 있음. 회사에서 실제로 테스트해서 결과를 알려주면 그에 맞게 조정할 것.
function Invoke-UiaDoubleClick {
    param([System.Windows.Automation.AutomationElement]$Element)

    $ok1 = Invoke-UiaClick -Element $Element
    Start-Sleep -Milliseconds 120
    $ok2 = Invoke-UiaClick -Element $Element
    return ($ok1 -or $ok2)
}

function Get-ElementDisplayValue {
    param([System.Windows.Automation.AutomationElement]$Element)
    $value = Get-PatternSafe -Element $Element -PatternClassName "ValuePattern"
    if ($value) {
        try { return $value.Current.Value } catch { }
    }
    try { return $Element.Current.Name } catch { return "" }
}

function Set-UiaValue {
    param([System.Windows.Automation.AutomationElement]$Element, [string]$Value)
    $value = Get-PatternSafe -Element $Element -PatternClassName "ValuePattern"
    if ($value) {
        try { $value.SetValue($Value); return $true } catch { Write-Log "[Set-UiaValue] SetValue 실패: $($_.Exception.Message)" }
    }
    Write-Log "[Set-UiaValue] ValuePattern 미지원 - 값 설정 실패"
    return $false
}

# ===================================================================
# 5. 그리드(표) 순회 - WinForms DataGridView는 보통 GridPattern을 지원하므로
#    가능하면 그걸 쓰고, 안 되면 DataItem 컨트롤타입으로 하위 요소를 훑음.
# ===================================================================

function Get-GridRows {
    param([System.Windows.Automation.AutomationElement]$GridElement)

    $grid = Get-PatternSafe -Element $GridElement -PatternClassName "GridPattern"
    if ($grid) {
        try {
            $rows = @()
            for ($r = 0; $r -lt $grid.Current.RowCount; $r++) {
                $rows += , @($grid.GetItem($r, 0))  # 필요한 컬럼 인덱스는 실제 화면 보고 조정
            }
            return $rows
        } catch { Write-Log "[Get-GridRows] GridPattern 사용 중 오류: $($_.Exception.Message)" }
    }

    # GridPattern이 없으면 DataItem/행으로 보이는 요소를 직접 훑음 (대체 경로)
    $cond = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
        [System.Windows.Automation.ControlType]::DataItem)
    try {
        return @($GridElement.FindAll([System.Windows.Automation.TreeScope]::Descendants, $cond))
    } catch {
        return @()
    }
}

# ===================================================================
# 6. 현상/원인/조치 텍스트 처리 - 실제 서식을 몰라도 되도록, 실행 시점에
#    텍스트박스 안에서 라벨 위치를 찾아 그 뒤에 끼워넣는 방식.
# ===================================================================

# 텍스트박스에 이미 세 가지 항목이 다 채워져 있는지 확인 (미작성 판별용).
# "라벨 뒤에 뭔가 내용이 있다"를 각 라벨마다 확인함.
function Test-FaultAlreadyWritten {
    param([string]$CurrentText)

    if ([string]::IsNullOrWhiteSpace($CurrentText)) { return $false }

    foreach ($key in $FaultLabelPatterns.Keys) {
        $pattern = $FaultLabelPatterns[$key] + '(\S.*)?$'
        $m = [regex]::Match($CurrentText, $FaultLabelPatterns[$key] + '([^\r\n]*)')
        if (-not $m.Success -or [string]::IsNullOrWhiteSpace($m.Groups[1].Value)) {
            return $false
        }
    }
    return $true
}

# 라벨(현상/원인/조치) 뒤에 내용을 끼워넣은 새 텍스트를 만들어 반환.
# 라벨이 하나라도 안 보이면 $null을 반환(포맷이 예상과 다르다는 뜻이므로
# 호출부에서 실패 처리하도록 함 - 여기서 임의로 추측해서 덮어쓰지 않음).
function Set-FaultTextByLabels {
    param(
        [string]$CurrentText,
        [string]$Symptom,   # 현상
        [string]$Cause,     # 원인
        [string]$Action     # 조치
    )

    $values = @{ 현상 = $Symptom; 원인 = $Cause; 조치 = $Action }
    $result = $CurrentText
    foreach ($key in $FaultLabelPatterns.Keys) {
        $regex = [regex]::new($FaultLabelPatterns[$key] + '([^\r\n]*)')
        $m = $regex.Match($result)
        if (-not $m.Success) {
            Write-Log "[Set-FaultTextByLabels] '$key' 라벨을 텍스트박스에서 찾지 못함 - 서식이 예상과 다를 수 있음"
            return $null
        }
        $replacement = $m.Value.Substring(0, $m.Value.Length - $m.Groups[1].Value.Length) + $values[$key]
        $result = $result.Substring(0, $m.Index) + $replacement + $result.Substring($m.Index + $m.Length)
    }
    return $result
}

# ===================================================================
# 7. 정지 키 / STOP.txt (Update-EmsAlarms.ps1과 동일)
# ===================================================================

$script:StopState = [hashtable]::Synchronized(@{ KeyPressed = $false; Exit = $false })

function Start-StopKeyWatcher {
    param([int]$VirtualKeyCode)

    $rs = [runspacefactory]::CreateRunspace()
    $rs.Open()
    $rs.SessionStateProxy.SetVariable("StopState", $script:StopState)
    $rs.SessionStateProxy.SetVariable("VkCode", $VirtualKeyCode)

    $ps = [powershell]::Create()
    $ps.Runspace = $rs
    [void]$ps.AddScript({
        [void][FaultReportAutomation.StopKeyNative]::GetAsyncKeyState($VkCode)
        while (-not $StopState.Exit) {
            if (([FaultReportAutomation.StopKeyNative]::GetAsyncKeyState($VkCode) -band 0x8001) -ne 0) {
                $StopState.KeyPressed = $true
                break
            }
            Start-Sleep -Milliseconds 40
        }
    })
    $ps.BeginInvoke() | Out-Null
    return [pscustomobject]@{ PowerShell = $ps; Runspace = $rs }
}

function Stop-StopKeyWatcher {
    param($Watcher)
    if (-not $Watcher) { return }
    $script:StopState.Exit = $true
    try { $Watcher.PowerShell.Stop() } catch { }
    try { $Watcher.PowerShell.Dispose() } catch { }
    try { $Watcher.Runspace.Close() } catch { }
    try { $Watcher.Runspace.Dispose() } catch { }
}

function Test-StopRequested {
    if ($script:StopState.KeyPressed) { return $true }
    return (Test-Path -LiteralPath $StopFlagPath)
}

# ===================================================================
# 8. 실시간 결과 창 (Update-EmsAlarms.ps1과 동일한 GUI 패턴)
# ===================================================================

function New-ResultsWindow {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "고장 자동작성 결과"
    $form.Width = 900
    $form.Height = 500
    $form.StartPosition = "CenterScreen"

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Dock = "Fill"
    $grid.ReadOnly = $true
    $grid.AllowUserToAddRows = $false
    $grid.Columns.Add("장비명", "장비명") | Out-Null
    $grid.Columns.Add("알람코드", "알람코드") | Out-Null
    $grid.Columns.Add("처리결과", "처리결과") | Out-Null
    $grid.Columns.Add("검증완료", "검증완료") | Out-Null
    $grid.Columns.Add("실패사유", "실패사유") | Out-Null
    $form.Controls.Add($grid)
    $form.Tag = $grid
    return $form
}

function Add-ResultToWindow {
    param($ResultsForm, $Row)
    $grid = $ResultsForm.Tag
    $rowIndex = $grid.Rows.Add()
    $r = $grid.Rows[$rowIndex]
    $r.Cells["장비명"].Value   = $Row.장비명
    $r.Cells["알람코드"].Value = $Row.알람코드
    $r.Cells["처리결과"].Value = $Row.처리결과
    $r.Cells["검증완료"].Value = $Row.검증완료
    $r.Cells["실패사유"].Value = $Row.실패사유
    $grid.FirstDisplayedScrollingRowIndex = $rowIndex
}

# ===================================================================
# 9. 한 건 처리 - 뼈대만 있음. 실제 화면 구조 확인 후 채워야 함(TODO 다수).
# ===================================================================

function Process-FaultRow {
    param(
        [System.Windows.Automation.AutomationElement]$MainWindow,
        [System.Windows.Automation.AutomationElement]$FaultRowElement,
        [pscustomobject]$CsvRow
    )

    $result = [pscustomobject]@{
        장비명   = $CsvRow.장비명
        알람코드 = $CsvRow.알람코드
        처리결과 = "실패"
        검증완료 = "N"
        실패사유 = ""
    }

    try {
        # TODO: FaultRowElement 안에서 "알람명" 칸을 찾아 더블클릭 -> 입력 화면 진입
        # $alarmCell = ...
        # Invoke-UiaDoubleClick -Element $alarmCell

        $cond = New-ConditionFromConfig $Config.FaultTextBox
        $textBoxEl = Wait-UIAElement -Parent $MainWindow -Condition $cond -TimeoutSec $TimeoutSec
        if (-not $textBoxEl) { throw "고장 입력 텍스트박스를 찾지 못했습니다." }

        $currentText = Get-ElementDisplayValue -Element $textBoxEl
        if (Test-FaultAlreadyWritten -CurrentText $currentText) {
            $result.처리결과 = "건너뜀(이미작성됨)"
            $result.검증완료 = "N/A"
            return $result
        }

        $newText = Set-FaultTextByLabels -CurrentText $currentText `
            -Symptom $CsvRow.현상 -Cause $CsvRow.원인 -Action $CsvRow.조치
        if (-not $newText) { throw "텍스트박스 서식(라벨)이 예상과 달라 안전하게 채우지 못했습니다." }

        Set-UiaValue -Element $textBoxEl -Value $newText

        $cond = New-ConditionFromConfig $Config.SaveButton
        $saveEl = Find-ElementNow -Parent $MainWindow -Condition $cond
        if (-not $saveEl) { throw "저장 버튼을 찾지 못했습니다." }
        Invoke-UiaClick -Element $saveEl

        Start-Sleep -Milliseconds 1000

        # TODO: 재조회(검증) - 목록으로 돌아가서 같은 행을 다시 열어 값이
        # 실제로 반영됐는지 확인하는 로직. Update-EmsAlarms.ps1의 재검증
        # 패턴(Process-Row 참고)과 동일한 방식으로 구현하면 됨.

        $result.처리결과 = "성공(검증 로직 미구현)"
    } catch {
        $result.실패사유 = $_.Exception.Message
        Write-Log "[Process-FaultRow] 실패: $($_.Exception.Message)"
    }

    return $result
}

# ===================================================================
# 10. 메인 실행부
# ===================================================================

Write-Log "===== 고장 자동작성 시작 (스크립트 버전: $ScriptVersion) ====="
Write-Log "[안내] 이 스크립트는 아직 뼈대 단계입니다. `$Config` 안의 TODO 항목들을"
Write-Log "       실제 화면 확인 후 채워야 정상 동작합니다."

if (Test-Path -LiteralPath $StopFlagPath) {
    Remove-Item -LiteralPath $StopFlagPath -Force
    Write-Log "이전 실행에서 남은 STOP.txt 를 삭제했습니다."
}

if (-not (Test-Path -LiteralPath $DataCsvPath)) {
    Write-Log "오류: fault_data.csv 를 찾을 수 없습니다 ($DataCsvPath). 스크립트를 종료합니다."
    exit 1
}

$rows = Import-Csv -Path $DataCsvPath -Encoding UTF8
if ($TestMode) {
    Write-Log "테스트 모드: 앞의 $TestModeRows 건만 처리합니다."
    $rows = $rows | Select-Object -First $TestModeRows
}

Write-Log "대상 프로그램 창을 찾는 중 (프로세스명에 '$AppProcessNameContains', 제목에 '$AppWindowTitleContains' 포함)..."
$window = Find-AppWindow -ProcessNameContains $AppProcessNameContains -TitleContains $AppWindowTitleContains
if (-not $window) {
    Write-Log "오류: 대상 창을 찾지 못했습니다. `$AppProcessNameContains / `$AppWindowTitleContains 값을 확인하세요."
    exit 1
}
Write-Log "대상 창 확보. PID=$($window.Current.ProcessId), 제목='$($window.Current.Name)'"

$resultsForm = New-ResultsWindow
$resultsForm.Show()
[System.Windows.Forms.Application]::DoEvents()

$stopKeyWatcher = Start-StopKeyWatcher -VirtualKeyCode $StopKeyCode
Write-Log "정지 방법: $StopKeyName 키 누르기 또는 STOP.txt 파일 생성."

if (Test-Path -LiteralPath $ResultCsvPath) {
    Remove-Item -LiteralPath $ResultCsvPath -Force
}

# TODO: KPI 버튼 클릭 -> 일자별 그리드 대기 -> 각 날짜 칸을 더블클릭해서
# 그 날의 고장 목록을 열고, CSV의 장비명+알람코드와 일치하는 행을 찾아
# Process-FaultRow 호출. 아래는 자리만 잡아둔 것.

$cond = New-ConditionFromConfig $Config.KpiButton
$kpiEl = Wait-UIAElement -Parent $window -Condition $cond -TimeoutSec $TimeoutSec
if (-not $kpiEl) {
    Write-Log "오류: KPI 버튼을 찾지 못했습니다. `$Config.KpiButton 의 AutomationId/Name을 확인하세요."
} else {
    Invoke-UiaClick -Element $kpiEl
    Write-Log "KPI 버튼 클릭 완료. (이후 날짜별 그리드 탐색 로직은 미구현 - TODO)"
}

Stop-StopKeyWatcher -Watcher $stopKeyWatcher

Write-Log "===== 종료 (뼈대 단계이므로 실제 처리는 수행되지 않았습니다) ====="
$resultsForm.Text = "$($resultsForm.Text) - 뼈대 단계 (실제 화면 정보 반영 필요)"

[System.Windows.Forms.Application]::Run($resultsForm)
