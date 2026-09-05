<#
    마우스 위치 요소 즉석 확인 도구
    ------------------------------
    전체 트리를 다 뽑아놓고 원하는 요소를 찾는 게 힘들 때 쓰는 보조 도구입니다.
    마우스 커서를 확인하고 싶은 요소(버튼/입력창/이미지 등) 위에 놓고 Enter를 누르면,
    바로 그 요소의 정보를 보여줍니다. 그 요소를 감싸는 "틀"(부모 요소)들도 몇 단계
    위까지 같이 보여주므로, 이미지 버튼처럼 이름이 없는 요소도 부모 쪽에 식별 정보가
    있는지 바로 확인할 수 있습니다.

    사용법
    ------
    .\Get-ElementAtCursor.ps1
      1) 콘솔에 안내가 뜨면, 확인하고 싶은 요소 위에 마우스를 갖다 놓으세요.
      2) (마우스는 그대로 두고) 콘솔 창으로 포커스를 옮겨서 Enter를 누르세요.
         → 그 순간 마우스 아래에 있던 요소 + 부모 몇 단계 정보가 출력됩니다.
      3) 다른 요소도 계속 확인하려면 1~2번을 반복하면 됩니다.
      4) 끝내려면 q 입력 후 Enter.

    -AncestorLevels 옵션으로 부모를 몇 단계까지 같이 보여줄지 조절할 수 있습니다
    (기본 5단계).
#>

[CmdletBinding()]
param(
    [int]$AncestorLevels = 5
)

if ($PSVersionTable.PSEdition -ne 'Desktop') {
    Write-Warning "이 스크립트는 Windows PowerShell 5.1(powershell.exe) 기준으로 검증되었습니다. 현재 PSEdition='$($PSVersionTable.PSEdition)' 입니다."
}

Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

function Test-PatternSupported {
    param([System.Windows.Automation.AutomationElement]$Element, $PatternObj)
    try {
        $p = $null
        return [bool]$Element.TryGetCurrentPattern($PatternObj, [ref]$p)
    } catch {
        return $false
    }
}

function Show-ElementInfo {
    param(
        [System.Windows.Automation.AutomationElement]$Element,
        [string]$Prefix
    )

    $ct = "오류"; $name = ""; $autoId = ""; $className = ""; $helpText = ""
    $isEnabled = ""; $rectStr = ""; $procName = "?"

    try { $ct = ($Element.Current.ControlType.ProgrammaticName -replace '^ControlType\.', '') } catch { }
    try { $name = $Element.Current.Name } catch { }
    try { $autoId = $Element.Current.AutomationId } catch { }
    try { $className = $Element.Current.ClassName } catch { }
    try { $helpText = $Element.Current.HelpText } catch { }
    try { $isEnabled = $Element.Current.IsEnabled } catch { }
    try {
        $r = $Element.Current.BoundingRectangle
        $rectStr = "{0:F0},{1:F0},{2:F0},{3:F0}" -f $r.X, $r.Y, $r.Width, $r.Height
    } catch { }
    try {
        $proc = Get-Process -Id $Element.Current.ProcessId -ErrorAction SilentlyContinue
        if ($proc) { $procName = $proc.ProcessName }
    } catch { }

    $hasInvoke = Test-PatternSupported -Element $Element -PatternObj ([System.Windows.Automation.InvokePattern]::Pattern)
    $hasValue  = Test-PatternSupported -Element $Element -PatternObj ([System.Windows.Automation.ValuePattern]::Pattern)
    $hasLegacy = Test-PatternSupported -Element $Element -PatternObj ([System.Windows.Automation.LegacyIAccessiblePattern]::Pattern)

    Write-Host "${Prefix}[$ct] Name='$name' AutomationId='$autoId'"
    Write-Host "${Prefix}    ClassName='$className' HelpText='$helpText' IsEnabled=$isEnabled"
    Write-Host "${Prefix}    BoundingRect(X,Y,W,H)=$rectStr Process=$procName"
    Write-Host "${Prefix}    InvokePattern=$hasInvoke  ValuePattern=$hasValue  LegacyIAccessible(DoDefaultAction)=$hasLegacy"
}

$walker = [System.Windows.Automation.TreeWalker]::RawViewWalker

Write-Host "===================================================================="
Write-Host " 확인하고 싶은 요소 위에 마우스를 놓고, 이 콘솔 창에서 Enter를 누르세요."
Write-Host " 끝내려면 q 입력 후 Enter."
Write-Host "===================================================================="

while ($true) {
    $cmd = Read-Host "`n[Enter]=현재 마우스 위치 요소 확인   [q]=종료"
    if ($cmd -eq 'q' -or $cmd -eq 'Q') { break }

    $pos = [System.Windows.Forms.Cursor]::Position
    $point = New-Object System.Windows.Point($pos.X, $pos.Y)

    $element = $null
    try {
        $element = [System.Windows.Automation.AutomationElement]::FromPoint($point)
    } catch {
        Write-Host "요소를 찾는 중 오류: $($_.Exception.Message)"
        continue
    }
    if (-not $element) {
        Write-Host "그 위치에서 요소를 찾지 못했습니다."
        continue
    }

    # 마우스 아래 요소부터 부모(틀) 방향으로 $AncestorLevels 단계까지 수집
    $chain = New-Object System.Collections.Generic.List[object]
    $current = $element
    for ($i = 0; $i -le $AncestorLevels; $i++) {
        if (-not $current) { break }
        $chain.Add($current)
        try { $current = $walker.GetParent($current) } catch { $current = $null }
    }
    $chain.Reverse()  # 맨 바깥쪽 조상부터 -> 실제 마우스 아래 요소 순서로 출력

    Write-Host "`n--- 마우스 위치(X=$($pos.X), Y=$($pos.Y)) 요소 및 부모 체인 (바깥→안쪽) ---"
    for ($i = 0; $i -lt $chain.Count; $i++) {
        $isTarget = ($i -eq $chain.Count - 1)
        $marker = if ($isTarget) { ">> " } else { "   " }
        $indent = "  " * $i
        Show-ElementInfo -Element $chain[$i] -Prefix ($indent + $marker)
    }
    if ($isTarget -and $chain.Count -gt 0) {
        Write-Host "(>> 표시가 실제로 마우스 아래에 있던 요소입니다. 그 위 줄들은 그 요소를 감싸는 틀/부모입니다)"
    }
}

Write-Host "종료합니다."
