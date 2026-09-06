<#
    EMS 화면 전체 UIA 요소 트리 덤프 도구
    ------------------------------------
    Accessibility Insights 같은 별도 프로그램 설치 없이, Windows 기본 PowerShell +
    .NET UI Automation 만으로 특정 창(예: EMS가 열린 Edge 창) 내부의 모든 요소를
    한 번에 트리 구조 그대로 텍스트 파일로 뽑아주는 1회성 진단 도구입니다.
    (자동화 실행용이 아니라, 사람이 나중에 파일을 열어서 Ctrl+F로 찾아보기 위한 도구)

    Name/AutomationId가 비어있는 요소(이미지 버튼 등)도 절대 건너뛰지 않고 그대로
    포함합니다. 오류가 나는 요소가 있어도 전체 스크립트는 멈추지 않고 계속 진행합니다.

    사용법
    ------
    1) EMS 페이지를 Edge에서 열어둔 상태로 실행:
         .\Get-EmsUiaTree.ps1
       → 현재 열려있는 엣지(msedge.exe) 창 목록(번호/프로세스명/제목/PID)만 출력됩니다.
         (다른 프로그램 창은 이 자동화와 무관하므로 목록에 나오지 않습니다)
       → 그 중 EMS가 열린 Edge 창의 번호를 입력하세요.

    2) 결과는 스크립트와 같은 폴더에 uia_tree_YYYYMMDD_HHMMSS.txt 로 저장됩니다.

    3) 클릭/입력 가능한 요소만 빠르게 보고 싶다면:
         .\Get-EmsUiaTree.ps1 -OnlyInteractable

    4) Name/AutomationId/ClassName(=F12의 class 속성) 중 아무 곳에나 특정
       문자열이 포함된 요소만 보고 싶다면:
         .\Get-EmsUiaTree.ps1 -Filter "알람"

    5) 화면이 너무 복잡해서 얕은 깊이까지만 빠르게 훑어보고 싶다면:
         .\Get-EmsUiaTree.ps1 -MaxDepth 8

    실행 전 준비사항: EMS 페이지가 Edge에서 열려 있어야 합니다. 실행 정책 때문에
    막히면 먼저 `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass` 를
    실행한 뒤 다시 시도하세요(관리자 권한 불필요, 현재 PowerShell 창에만 적용).
#>

[CmdletBinding()]
param(
    [int]$MaxDepth = 0,          # 0 = 무제한. 화면이 복잡할 때만 작은 값으로 제한.
    [switch]$OnlyInteractable,   # InvokePattern 또는 ValuePattern 지원 요소만 출력
    [string]$Filter,             # Name, AutomationId, ClassName(=F12의 class 속성) 중 어디든 이 문자열이 포함된 요소만 출력
    [int]$ProgressInterval = 200 # 몇 개 요소마다 진행 상황을 출력할지
)

if ($PSVersionTable.PSEdition -ne 'Desktop') {
    Write-Warning "이 스크립트는 Windows PowerShell 5.1(powershell.exe) 기준으로 검증되었습니다. 현재 PSEdition='$($PSVersionTable.PSEdition)' 입니다."
}

Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$OutPath = Join-Path $ScriptDir "uia_tree_$timestamp.txt"

# ===================================================================
# 1. 열려있는 모든 최상위 창 목록을 번호 매겨 보여주고, 사람이 직접 선택
#    (제목/프로세스명만으로 자동 판단하지 않음 - 엣지 창이 여러 개일 수 있으므로)
# ===================================================================

$edgeProcessIds = @(Get-Process -Name "msedge" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
if ($edgeProcessIds.Count -eq 0) {
    Write-Host "실행 중인 Microsoft Edge(msedge.exe) 프로세스를 찾지 못했습니다. Edge를 먼저 열어주세요."
    return
}

$root = [System.Windows.Automation.AutomationElement]::RootElement
$windowCondition = New-Object System.Windows.Automation.PropertyCondition(
    [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
    [System.Windows.Automation.ControlType]::Window)
$allTopLevelWindows = @($root.FindAll([System.Windows.Automation.TreeScope]::Children, $windowCondition))

# 엣지(msedge.exe) 창만 남김 - 다른 프로그램은 이 자동화와 무관하므로 목록에서 제외
$topLevelWindows = @($allTopLevelWindows | Where-Object {
    try { $_.Current.ProcessId -in $edgeProcessIds } catch { $false }
})

if ($topLevelWindows.Count -eq 0) {
    Write-Host "Edge 프로세스는 실행 중이지만 UIA로 보이는 엣지 창이 없습니다. 창이 최소화되어 있지 않은지 확인해주세요."
    return
}

$windowList = @()
for ($i = 0; $i -lt $topLevelWindows.Count; $i++) {
    $w = $topLevelWindows[$i]
    $procName = "?"
    try {
        $procId = $w.Current.ProcessId
        $proc = Get-Process -Id $procId -ErrorAction SilentlyContinue
        if ($proc) { $procName = $proc.ProcessName }
    } catch { }
    $title = ""
    try { $title = $w.Current.Name } catch { }

    $windowList += [pscustomobject]@{
        Index       = $i
        ProcessName = $procName
        Pid_        = $procId
        Title       = $title
        Element     = $w
    }
}

Write-Host "현재 열려있는 엣지 창 목록:"
foreach ($item in $windowList) {
    Write-Host ("  [{0}] 프로세스={1,-12} PID={2,-8} 제목='{3}'" -f $item.Index, $item.ProcessName, $item.Pid_, $item.Title)
}

$selectedIndex = -1
while ($selectedIndex -lt 0 -or $selectedIndex -ge $windowList.Count) {
    $inputStr = Read-Host "요소를 뽑을 대상 창의 번호를 입력하세요"
    if (-not [int]::TryParse($inputStr, [ref]$selectedIndex)) {
        Write-Host "숫자를 입력해주세요."
        $selectedIndex = -1
        continue
    }
    if ($selectedIndex -lt 0 -or $selectedIndex -ge $windowList.Count) {
        Write-Host "목록에 있는 번호를 입력해주세요."
    }
}

$targetWindow = $windowList[$selectedIndex].Element
Write-Host "선택된 창: '$($windowList[$selectedIndex].Title)' (PID=$($windowList[$selectedIndex].Pid_))"
Write-Host "요소 트리를 뽑는 중... (화면 크기에 따라 몇 초~수 분 걸릴 수 있습니다)"

# ===================================================================
# 2. 트리 전체 순회 (RawView 기준 - 레이아웃용 요소까지 전부 포함)
# ===================================================================

$walker = [System.Windows.Automation.TreeWalker]::RawViewWalker
$lines = New-Object System.Collections.Generic.List[string]
$script:visitedCount = 0
$script:writtenCount = 0

function Test-PatternSupported {
    param([System.Windows.Automation.AutomationElement]$Element, $PatternObj)
    try {
        $p = $null
        return [bool]$Element.TryGetCurrentPattern($PatternObj, [ref]$p)
    } catch {
        return $false
    }
}

# 요소 1개의 정보를 뽑아 파이프 구분 한 줄로 만든다.
# 오류가 나도 예외를 던지지 않고, 문제가 난 필드만 "오류"로 표시한다.
function Build-ElementLine {
    param([System.Windows.Automation.AutomationElement]$Element, [int]$Depth)

    $ct = "오류"; $name = ""; $autoId = ""; $className = ""; $helpText = ""
    $isEnabled = ""; $rectStr = ""; $hasInvoke = $false; $hasValue = $false

    try { $ct = ($Element.Current.ControlType.ProgrammaticName -replace '^ControlType\.', '') } catch { $ct = "오류" }
    try { $name = $Element.Current.Name } catch { $name = "" }
    try { $autoId = $Element.Current.AutomationId } catch { $autoId = "" }
    try { $className = $Element.Current.ClassName } catch { $className = "" }
    try { $helpText = $Element.Current.HelpText } catch { $helpText = "" }
    try { $isEnabled = $Element.Current.IsEnabled } catch { $isEnabled = "" }
    try {
        $r = $Element.Current.BoundingRectangle
        $rectStr = "{0:F0},{1:F0},{2:F0},{3:F0}" -f $r.X, $r.Y, $r.Width, $r.Height
    } catch { $rectStr = "" }

    $hasInvoke = Test-PatternSupported -Element $Element -PatternObj ([System.Windows.Automation.InvokePattern]::Pattern)
    $hasValue  = Test-PatternSupported -Element $Element -PatternObj ([System.Windows.Automation.ValuePattern]::Pattern)

    # 파이프(|)가 이름/텍스트 값 안에 섞여 컬럼이 밀리지 않도록 치환
    $safe = { param($s) if ($null -eq $s) { "" } else { ([string]$s) -replace '\|', '/' -replace "`r`n|`n", ' ' } }
    $name = & $safe $name
    $className = & $safe $className
    $helpText = & $safe $helpText

    $indent = "  " * $Depth
    $line = "{0}[{1}]|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}" -f `
        $indent, $Depth, $ct, $name, $autoId, $className, $helpText, $isEnabled, $rectStr, `
        $(if ($hasInvoke) { "Y" } else { "N" }), $(if ($hasValue) { "Y" } else { "N" })

    return [pscustomobject]@{
        Line = $line
        Name = $name
        AutomationId = $autoId
        ClassName = $className
        HasInvoke = $hasInvoke
        HasValue = $hasValue
    }
}

function Walk-Tree {
    param([System.Windows.Automation.AutomationElement]$Element, [int]$Depth)

    $script:visitedCount++
    if (($script:visitedCount % $ProgressInterval) -eq 0) {
        Write-Host "탐색 중... 현재까지 $($script:visitedCount)개 요소 처리"
    }

    try {
        $info = Build-ElementLine -Element $Element -Depth $Depth

        $include = $true
        if ($OnlyInteractable) { $include = ($info.HasInvoke -or $info.HasValue) }
        if ($include -and $Filter) {
            $include = ($info.Name -like "*$Filter*") -or ($info.AutomationId -like "*$Filter*") -or ($info.ClassName -like "*$Filter*")
        }
        if ($include) {
            $lines.Add($info.Line)
            $script:writtenCount++
        }
    } catch {
        $indent = "  " * $Depth
        $lines.Add("$indent[$Depth]|오류|오류|오류|오류|오류|오류|오류|오류|오류 (예외: $($_.Exception.Message))")
        $script:writtenCount++
    }

    if ($MaxDepth -gt 0 -and $Depth -ge $MaxDepth) { return }

    $child = $null
    try { $child = $walker.GetFirstChild($Element) } catch { $child = $null }
    while ($child) {
        Walk-Tree -Element $child -Depth ($Depth + 1)
        $next = $null
        try { $next = $walker.GetNextSibling($child) } catch { $next = $null }
        $child = $next
    }
}

$header = "[Depth]|ControlType|Name|AutomationId|ClassName|HelpText|IsEnabled|BoundingRect(X,Y,W,H)|InvokePattern|ValuePattern"
$lines.Add($header)

Walk-Tree -Element $targetWindow -Depth 0

$lines | Out-File -FilePath $OutPath -Encoding UTF8

Write-Host "완료: 전체 $($script:visitedCount)개 요소 중 $($script:writtenCount)개를 '$OutPath' 에 저장했습니다."
Write-Host "이 파일을 메모장/엑셀로 열어서(파이프 | 기준으로 열 구분), 필요한 입력창/버튼의"
Write-Host "Name / AutomationId / BoundingRect / InvokePattern / ValuePattern 값을 확인하세요."
