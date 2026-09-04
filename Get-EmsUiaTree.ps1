<#
    EMS 화면 요소 정보 뽑기 도구 (Accessibility Insights / UIATreeInspector 대체용)
    ------------------------------------------------------------------
    Accessibility Insights 같은 별도 프로그램을 설치할 수 없는 환경을 위한 도구입니다.
    Windows 기본 PowerShell + .NET UI Automation 만으로, 화면에 있는 요소들의
    Name / AutomationId / ControlType 을 뽑아 텍스트 파일로 저장합니다.
    이렇게 뽑은 값을 Update-EmsAlarms.ps1 상단의 $Config 에 그대로 채워 넣으면 됩니다.

    사용법
    ------
    1) EMS 페이지를 Edge에서 열어둔 상태에서 아래처럼 실행 (창 제목을 모르면 그냥 실행):
         .\Get-EmsUiaTree.ps1
       → 현재 열려있는 모든 창의 제목 목록을 보여줍니다. 그중 EMS 창 제목을 확인하세요.

    2) 창 제목 일부를 알면, 그 창의 요소 트리를 파일로 저장:
         .\Get-EmsUiaTree.ps1 -TitleContains "EMS"
       → 스크립트와 같은 폴더에 ems_uia_tree.txt 파일이 생성됩니다.

    3) 입력창/버튼/텍스트뿐 아니라 전체 요소를 다 보고 싶다면:
         .\Get-EmsUiaTree.ps1 -TitleContains "EMS" -All
#>

[CmdletBinding()]
param(
    [string]$TitleContains,

    # 기본은 자동화에 쓸만한 컨트롤(입력창/버튼/텍스트/표 항목)만 출력. 전체를 보려면 -All 사용.
    [switch]$All,

    [int]$MaxDepth = 20,
    [int]$MaxNodes = 4000
)

if ($PSVersionTable.PSEdition -ne 'Desktop') {
    Write-Warning "이 스크립트는 Windows PowerShell 5.1(powershell.exe) 기준으로 검증되었습니다. 현재 PSEdition='$($PSVersionTable.PSEdition)' 입니다."
}

Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$OutPath = Join-Path $ScriptDir "ems_uia_tree.txt"

$InterestingControlTypes = @('Edit', 'Button', 'Text', 'DataItem', 'ListItem', 'Document', 'ComboBox', 'CheckBox')

# ===================================================================
# 1. 대상 창 찾기 - 실행 중인 엣지(msedge.exe) 창만 대상으로 함
#    (제목이 없으면 엣지 창 후보 목록만 보여주고 종료)
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
$allTopLevelWindows = $root.FindAll([System.Windows.Automation.TreeScope]::Children, $windowCondition)

# 전체 창 중 프로세스가 msedge 인 것만 남김 (다른 프로그램 창이 섞이지 않도록)
$allWindows = @($allTopLevelWindows | Where-Object {
    try { $_.Current.ProcessId -in $edgeProcessIds } catch { $false }
})

if ($allWindows.Count -eq 0) {
    Write-Host "Edge 프로세스는 실행 중이지만 UIA로 보이는 엣지 창이 없습니다. 창이 최소화되어 있지 않은지 확인해주세요."
    return
}

if (-not $TitleContains) {
    Write-Host "현재 열려있는 엣지 창 목록 (이 중 EMS 창 제목의 일부를 -TitleContains 로 넘겨주세요):"
    foreach ($w in $allWindows) {
        try {
            Write-Host ("  PID={0,-8} 제목='{1}'" -f $w.Current.ProcessId, $w.Current.Name)
        } catch { }
    }
    return
}

$targetWindow = $null
foreach ($w in $allWindows) {
    try {
        if ($w.Current.Name -like "*$TitleContains*") { $targetWindow = $w; break }
    } catch { }
}

if (-not $targetWindow) {
    Write-Host "엣지 창 중에서 제목에 '$TitleContains' 를 포함하는 창을 찾지 못했습니다. 아래 목록을 참고하세요:"
    foreach ($w in $allWindows) {
        try {
            Write-Host ("  PID={0,-8} 제목='{1}'" -f $w.Current.ProcessId, $w.Current.Name)
        } catch { }
    }
    return
}

Write-Host "대상 창: '$($targetWindow.Current.Name)' (PID=$($targetWindow.Current.ProcessId))"
Write-Host "요소 트리를 뽑는 중... (화면 크기에 따라 몇 초~수십 초 걸릴 수 있습니다)"

# ===================================================================
# 2. 트리 순회 후 파일로 저장
# ===================================================================

$lines = New-Object System.Collections.Generic.List[string]
$nodeCount = 0
$walker = [System.Windows.Automation.TreeWalker]::ControlViewWalker

function Format-NodeLine {
    param($Element, [int]$Depth)
    $ct = $null
    $name = $null
    $autoId = $null
    $className = $null
    try { $ct = $Element.Current.ControlType.ProgrammaticName -replace '^ControlType\.', '' } catch { $ct = "?" }
    try { $name = $Element.Current.Name } catch { $name = "" }
    try { $autoId = $Element.Current.AutomationId } catch { $autoId = "" }
    try { $className = $Element.Current.ClassName } catch { $className = "" }

    $indent = "  " * $Depth
    return "${indent}[$ct] Name='$name' AutomationId='$autoId' ClassName='$className'"
}

function Walk-Tree {
    param($Element, [int]$Depth)

    if ($script:nodeCount -ge $MaxNodes -or $Depth -gt $MaxDepth) { return }

    $ctName = $null
    try { $ctName = $Element.Current.ControlType.ProgrammaticName -replace '^ControlType\.', '' } catch { }

    $show = $All -or ($ctName -in $InterestingControlTypes)
    if ($show) {
        $lines.Add((Format-NodeLine -Element $Element -Depth $Depth))
        $script:nodeCount++
    }

    $child = $walker.GetFirstChild($Element)
    while ($child) {
        Walk-Tree -Element $child -Depth ($Depth + 1)
        $child = $walker.GetNextSibling($child)
    }
}

Walk-Tree -Element $targetWindow -Depth 0

$lines | Out-File -FilePath $OutPath -Encoding UTF8

Write-Host "완료: $nodeCount 개 요소를 '$OutPath' 에 저장했습니다."
Write-Host "이 파일을 열어서, 원하는 입력창/버튼에 해당하는 줄의 Name / AutomationId 값을"
Write-Host "Update-EmsAlarms.ps1 의 `$Config 에 그대로 옮겨 적으면 됩니다."
if ($nodeCount -ge $MaxNodes) {
    Write-Host "주의: 요소 수가 너무 많아 $MaxNodes 개에서 잘렸습니다. 필요하면 -MaxNodes 값을 늘려서 다시 실행하세요."
}
