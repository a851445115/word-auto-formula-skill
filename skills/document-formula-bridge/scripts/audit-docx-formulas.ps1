[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$InputPath,

    [switch]$AsJson
)

$ErrorActionPreference = "Stop"

function Resolve-AbsoluteExistingPath {
    param([Parameter(Mandatory = $true)][string]$Path)

    $resolved = Resolve-Path -LiteralPath $Path -ErrorAction Stop
    return [System.IO.Path]::GetFullPath($resolved.Path)
}

function Release-ComObject {
    param($ComObject)

    if ($null -ne $ComObject) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject)
    }
}

function New-WordApplication {
    $word = New-Object -ComObject Word.Application
    $word.DisplayAlerts = 0
    $word.Visible = $false

    if ($word.Name -notlike "*Microsoft Word*" -or $word.Path -match "WPS|Kingsoft") {
        $name = $word.Name
        $path = $word.Path
        try {
            $word.Quit()
        } catch {
        }
        Release-ComObject $word
        throw "Word.Application resolved to '$name' at '$path'. Repair Word COM registration before auditing MathType output."
    }

    return $word
}

function Test-MathTypeInlineShape {
    param($InlineShape)

    try {
        return $InlineShape.OLEFormat.ProgID -eq "Equation.DSMT4"
    } catch {
        return $false
    }
}

function Test-MathTypeShape {
    param($Shape)

    try {
        return $Shape.OLEFormat.ProgID -eq "Equation.DSMT4"
    } catch {
        return $false
    }
}

function Get-MathTypeObjectCount {
    param([Parameter(Mandatory = $true)]$Document)

    $count = 0

    for ($i = 1; $i -le $Document.InlineShapes.Count; $i++) {
        if (Test-MathTypeInlineShape -InlineShape $Document.InlineShapes.Item($i)) {
            $count++
        }
    }

    for ($i = 1; $i -le $Document.Shapes.Count; $i++) {
        if (Test-MathTypeShape -Shape $Document.Shapes.Item($i)) {
            $count++
        }
    }

    return $count
}

function Get-LatexMatches {
    param(
        [Parameter(Mandatory = $true)][string]$Text,
        [Parameter(Mandatory = $true)][string]$Pattern
    )

    return [System.Text.RegularExpressions.Regex]::Matches(
        $Text,
        $Pattern,
        [System.Text.RegularExpressions.RegexOptions]::Singleline
    )
}

function Get-MatchSamples {
    param(
        $Matches,
        [Parameter(Mandatory = $true)][string]$GroupName
    )

    if ($null -eq $Matches) {
        return @()
    }

    $allMatches = @($Matches) | Where-Object { $null -ne $_ }
    $samples = New-Object System.Collections.Generic.List[string]
    if ($allMatches.Count -eq 0) {
        return @()
    }
    $limit = [Math]::Min($allMatches.Count, 5)

    for ($i = 0; $i -lt $limit; $i++) {
        $value = $allMatches[$i].Groups[$GroupName].Value
        $value = ($value -replace "\s+", " ").Trim()
        if ($value.Length -gt 80) {
            $value = $value.Substring(0, 80) + "..."
        }
        $samples.Add($value)
    }

    return @($samples)
}

$resolvedInputPath = Resolve-AbsoluteExistingPath -Path $InputPath
$word = $null
$document = $null

try {
    $word = New-WordApplication
    $document = $word.Documents.Open($resolvedInputPath)
    $text = $document.Range().Text
    $inlineMatches = Get-LatexMatches -Text $text -Pattern '(?<!\$)\$(?!\$)(?<content>[^\r$]+?)(?<!\$)\$(?!\$)'
    $displayMatches = Get-LatexMatches -Text $text -Pattern '\$\$(?<content>[^\r]+?)\$\$'

    $result = [pscustomobject]@{
        InputPath             = $resolvedInputPath
        WordName              = $word.Name
        WordVersion           = $word.Version
        WordPath              = $word.Path
        MathTypeObjects       = Get-MathTypeObjectCount -Document $document
        RemainingInlineLatex  = $inlineMatches.Count
        RemainingDisplayLatex = $displayMatches.Count
        InlineSamples         = @(Get-MatchSamples -Matches $inlineMatches -GroupName "content")
        DisplaySamples        = @(Get-MatchSamples -Matches $displayMatches -GroupName "content")
    }

    if ($AsJson) {
        $result | ConvertTo-Json -Depth 4
    } else {
        $result
    }
} finally {
    if ($null -ne $document) {
        try {
            $document.Close(0)
        } catch {
        }
        Release-ComObject $document
    }

    if ($null -ne $word) {
        try {
            $word.Quit()
        } catch {
        }
        Release-ComObject $word
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
