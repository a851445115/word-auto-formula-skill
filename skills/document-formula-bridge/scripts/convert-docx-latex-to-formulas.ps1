[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$InputPath,

    [string]$OutputPath,

    [switch]$InPlace,

    [switch]$Visible,

    [string]$MathTypeTemplatePath,

    [ValidateRange(0, 10)]
    [int]$MaxResidualPasses = 2,

    [switch]$AggressiveCleanup,

    [switch]$PreflightOnly
)

$ErrorActionPreference = "Stop"

function Resolve-AbsoluteExistingPath {
    param([Parameter(Mandatory = $true)][string]$Path)

    $resolved = Resolve-Path -LiteralPath $Path -ErrorAction Stop
    return [System.IO.Path]::GetFullPath($resolved.Path)
}

function Resolve-AbsolutePath {
    param([Parameter(Mandatory = $true)][string]$Path)

    if ([System.IO.Path]::IsPathRooted($Path)) {
        return [System.IO.Path]::GetFullPath($Path)
    }

    return [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $Path))
}

function Get-DefaultOutputPath {
    param([Parameter(Mandatory = $true)][string]$ResolvedInputPath)

    $directory = [System.IO.Path]::GetDirectoryName($ResolvedInputPath)
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($ResolvedInputPath)
    $extension = [System.IO.Path]::GetExtension($ResolvedInputPath)
    return Join-Path $directory ("{0}_mathtype{1}" -f $baseName, $extension)
}

function Get-MathTypeTemplatePath {
    param([string]$PreferredPath)

    $candidates = @()
    if ($PreferredPath) {
        $candidates += (Resolve-AbsolutePath -Path $PreferredPath)
    }

    $candidates += @(
        "C:\Program Files (x86)\MathType\Office Support\32\MathType Commands 2016.dotm",
        "C:\Program Files (x86)\MathType\Office Support\64\MathType Commands 2016.dotm",
        "C:\Program Files\MathType\Office Support\32\MathType Commands 2016.dotm",
        "C:\Program Files\MathType\Office Support\64\MathType Commands 2016.dotm"
    )

    foreach ($candidate in $candidates | Select-Object -Unique) {
        if (Test-Path -LiteralPath $candidate) {
            return $candidate
        }
    }

    throw "MathType template not found. Pass -MathTypeTemplatePath explicitly or install MathType Office Support."
}

function Release-ComObject {
    param($ComObject)

    if ($null -ne $ComObject) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject)
    }
}

function New-WordApplication {
    param([switch]$ShowWindow)

    $word = New-Object -ComObject Word.Application
    $word.DisplayAlerts = 0
    $word.Visible = [bool]$ShowWindow

    if ($word.Name -notlike "*Microsoft Word*" -or $word.Path -match "WPS|Kingsoft") {
        $name = $word.Name
        $path = $word.Path
        try {
            $word.Quit()
        } catch {
        }
        Release-ComObject $word
        throw "Word.Application resolved to '$name' at '$path'. Repair Word COM registration before running MathType automation."
    }

    return $word
}

function Ensure-MathTypeAddIn {
    param(
        [Parameter(Mandatory = $true)]$Word,
        [Parameter(Mandatory = $true)][string]$TemplatePath
    )

    for ($i = 1; $i -le $Word.AddIns.Count; $i++) {
        $addIn = $Word.AddIns.Item($i)
        if ($addIn.FullName -eq $TemplatePath -or $addIn.Name -eq [System.IO.Path]::GetFileName($TemplatePath)) {
            $addIn.Installed = $true
            return $addIn
        }
    }

    return $Word.AddIns.Add($TemplatePath, $true)
}

function Invoke-MathTypeTexToggle {
    param([Parameter(Mandatory = $true)]$Word)

    try {
        $null = $Word.GetType().InvokeMember(
            "Run",
            [System.Reflection.BindingFlags]::InvokeMethod,
            $null,
            $Word,
            @("MTCommand_OnTexToggle", $null)
        )
    } catch {
        $message = $_.Exception.Message
        if ($_.Exception.InnerException) {
            $message = $_.Exception.InnerException.Message
        }
        throw "Failed to run MTCommand_OnTexToggle. $message"
    }
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

function Get-DocumentAudit {
    param([Parameter(Mandatory = $true)]$Document)

    $text = $Document.Range().Text
    $inlineMatches = Get-LatexMatches -Text $text -Pattern '(?<!\$)\$(?!\$)(?<content>[^\r$]+?)(?<!\$)\$(?!\$)'
    $displayMatches = Get-LatexMatches -Text $text -Pattern '\$\$(?<content>[^\r]+?)\$\$'

    return [pscustomobject]@{
        MathTypeObjects       = Get-MathTypeObjectCount -Document $Document
        RemainingInlineLatex  = $inlineMatches.Count
        RemainingDisplayLatex = $displayMatches.Count
        InlineSamples         = @(Get-MatchSamples -Matches $inlineMatches -GroupName "content")
        DisplaySamples        = @(Get-MatchSamples -Matches $displayMatches -GroupName "content")
    }
}

function Remove-DollarMarkersAroundMathTypeObjects {
    param([Parameter(Mandatory = $true)]$Document)

    $positions = New-Object System.Collections.Generic.HashSet[int]
    $documentEnd = $Document.Range().End

    for ($i = 1; $i -le $Document.InlineShapes.Count; $i++) {
        $inlineShape = $Document.InlineShapes.Item($i)
        if (-not (Test-MathTypeInlineShape -InlineShape $inlineShape)) {
            continue
        }

        $range = $inlineShape.Range
        if ($range.Start -gt 0) {
            $leftRange = $Document.Range($range.Start - 1, $range.Start)
            if ($leftRange.Text -eq '$') {
                [void]$positions.Add($range.Start - 1)
            }
        }

        if ($range.End -lt $documentEnd) {
            $rightRange = $Document.Range($range.End, $range.End + 1)
            if ($rightRange.Text -eq '$') {
                [void]$positions.Add($range.End)
            }
        }
    }

    $removed = 0
    foreach ($position in ($positions | Sort-Object -Descending)) {
        $Document.Range($position, $position + 1).Delete()
        $removed++
    }

    return $removed
}

function Remove-DelimiterOnlyParagraphs {
    param([Parameter(Mandatory = $true)]$Document)

    $removed = 0

    for ($i = $Document.Paragraphs.Count; $i -ge 1; $i--) {
        $paragraph = $Document.Paragraphs.Item($i)
        $text = ($paragraph.Range.Text -replace "[`r`a]", "").Trim()
        if ($text -eq '$' -or $text -eq '$$') {
            $paragraph.Range.Delete()
            $removed++
        }
    }

    return $removed
}

function Invoke-DisplayLatexPass {
    param(
        [Parameter(Mandatory = $true)]$Word,
        [Parameter(Mandatory = $true)]$Document
    )

    $matches = Get-LatexMatches -Text $Document.Range().Text -Pattern '\$\$(?<content>[^\r]+?)\$\$'
    $attempts = 0

    for ($i = $matches.Count - 1; $i -ge 0; $i--) {
        $match = $matches.Item($i)
        $content = ($match.Groups["content"].Value -replace "\s+", " ").Trim()
        if ([string]::IsNullOrWhiteSpace($content)) {
            continue
        }

        $range = $Document.Range($match.Index, $match.Index + $match.Length)
        $range.Select()
        Invoke-MathTypeTexToggle -Word $Word
        $attempts++
    }

    return $attempts
}

function Get-MathTypeSafeInlineTex {
    param([Parameter(Mandatory = $true)][string]$Content)

    $trimmed = ($Content -replace "\s+", " ").Trim()
    if ([string]::IsNullOrWhiteSpace($trimmed)) {
        return $null
    }

    if ($trimmed -match '^[0-9A-Za-z .,+\-=/()]+$' -and $trimmed -notmatch '[\\{}_^]') {
        return '\mathrm{' + $trimmed + '}'
    }

    return $trimmed
}

function Replace-RangeWithGeneratedMathTypeObject {
    param(
        [Parameter(Mandatory = $true)]$Word,
        [Parameter(Mandatory = $true)]$Document,
        [Parameter(Mandatory = $true)][int]$Start,
        [Parameter(Mandatory = $true)][int]$End,
        [Parameter(Mandatory = $true)][string]$Tex
    )

    $tempDocument = $null
    try {
        $tempDocument = $Word.Documents.Add()
        $tempDocument.Range().Text = '$' + $Tex + '$'
        $tempDocument.Range().Select()
        Invoke-MathTypeTexToggle -Word $Word

        $sourceRange = $null
        for ($i = 1; $i -le $tempDocument.InlineShapes.Count; $i++) {
            $inlineShape = $tempDocument.InlineShapes.Item($i)
            if (Test-MathTypeInlineShape -InlineShape $inlineShape) {
                $sourceRange = $inlineShape.Range
                break
            }
        }

        if ($null -eq $sourceRange) {
            return $false
        }

        $targetRange = $Document.Range($Start, $End)
        $targetRange.FormattedText = $sourceRange.FormattedText
        return $true
    } finally {
        if ($null -ne $tempDocument) {
            try {
                $tempDocument.Close(0)
            } catch {
            }
            Release-ComObject $tempDocument
        }
    }
}

function Invoke-InlineFallbackPass {
    param(
        [Parameter(Mandatory = $true)]$Word,
        [Parameter(Mandatory = $true)]$Document
    )

    $matches = Get-LatexMatches -Text $Document.Range().Text -Pattern '(?<!\$)\$(?!\$)(?<content>[^\r$]+?)(?<!\$)\$(?!\$)'
    $converted = 0

    for ($i = $matches.Count - 1; $i -ge 0; $i--) {
        $match = $matches.Item($i)
        $content = $match.Groups["content"].Value

        if ($content.Contains("`r") -or $content.Contains("`n") -or $content.Length -gt 200) {
            continue
        }

        $tex = Get-MathTypeSafeInlineTex -Content $content
        if ([string]::IsNullOrWhiteSpace($tex)) {
            continue
        }

        if (Replace-RangeWithGeneratedMathTypeObject -Word $Word -Document $Document -Start $match.Index -End ($match.Index + $match.Length) -Tex $tex) {
            $converted++
        }
    }

    return $converted
}

$resolvedInputPath = Resolve-AbsoluteExistingPath -Path $InputPath
$resolvedOutputPath = if ($OutputPath) {
    Resolve-AbsolutePath -Path $OutputPath
} else {
    Get-DefaultOutputPath -ResolvedInputPath $resolvedInputPath
}

if (-not $InPlace) {
    if ($resolvedInputPath -eq $resolvedOutputPath) {
        throw "OutputPath resolves to the input file. Use -InPlace explicitly or choose a different output path."
    }
}

$templatePath = Get-MathTypeTemplatePath -PreferredPath $MathTypeTemplatePath
$word = $null
$document = $null

try {
    $word = New-WordApplication -ShowWindow:$Visible
    $null = Ensure-MathTypeAddIn -Word $word -TemplatePath $templatePath

    if ($PreflightOnly) {
        [pscustomobject]@{
            InputPath            = $resolvedInputPath
            WordName             = $word.Name
            WordVersion          = $word.Version
            WordPath             = $word.Path
            MathTypeTemplatePath = $templatePath
            Ready                = $true
        }
        return
    }

    if (-not $InPlace) {
        $outputDirectory = [System.IO.Path]::GetDirectoryName($resolvedOutputPath)
        if ($outputDirectory) {
            [void][System.IO.Directory]::CreateDirectory($outputDirectory)
        }
        Copy-Item -LiteralPath $resolvedInputPath -Destination $resolvedOutputPath -Force
    } else {
        $resolvedOutputPath = $resolvedInputPath
    }

    Write-Host ("Word: {0} {1} at {2}" -f $word.Name, $word.Version, $word.Path)
    Write-Host ("MathType template: {0}" -f $templatePath)
    Write-Host ("Working document: {0}" -f $resolvedOutputPath)

    $document = $word.Documents.Open($resolvedOutputPath)
    $document.Activate()

    $beforeAudit = Get-DocumentAudit -Document $document
    Write-Host ("Initial audit: {0} MathType objects, {1} inline LaTeX, {2} display LaTeX" -f $beforeAudit.MathTypeObjects, $beforeAudit.RemainingInlineLatex, $beforeAudit.RemainingDisplayLatex)

    Write-Host "Running whole-document MathType toggle"
    $word.Selection.WholeStory()
    Invoke-MathTypeTexToggle -Word $word

    $removedDollarMarkers = 0
    $removedDelimiterParagraphs = 0
    $displayAttempts = 0
    $inlineFallbackConversions = 0

    if ($AggressiveCleanup) {
        Write-Host "Removing stray dollar markers around MathType objects"
        $removedDollarMarkers = Remove-DollarMarkersAroundMathTypeObjects -Document $document
        Write-Host ("Removed dollar markers: {0}" -f $removedDollarMarkers)

        Write-Host "Removing delimiter-only paragraphs"
        $removedDelimiterParagraphs = Remove-DelimiterOnlyParagraphs -Document $document
        Write-Host ("Removed delimiter-only paragraphs: {0}" -f $removedDelimiterParagraphs)

        for ($pass = 1; $pass -le $MaxResidualPasses; $pass++) {
            Write-Host ("Residual pass {0}: display retry" -f $pass)
            $displayPass = Invoke-DisplayLatexPass -Word $word -Document $document
            Write-Host ("Residual pass {0}: inline fallback" -f $pass)
            $inlinePass = Invoke-InlineFallbackPass -Word $word -Document $document

            $displayAttempts += $displayPass
            $inlineFallbackConversions += $inlinePass

            if (($displayPass + $inlinePass) -eq 0) {
                break
            }

            $removedDollarMarkers += Remove-DollarMarkersAroundMathTypeObjects -Document $document
            $removedDelimiterParagraphs += Remove-DelimiterOnlyParagraphs -Document $document
            Write-Host ("Residual pass {0}: display attempts={1}, inline fallback conversions={2}" -f $pass, $displayPass, $inlinePass)
        }
    }

    $document.Save()
    $finalAudit = Get-DocumentAudit -Document $document

    [pscustomobject]@{
        InputPath                  = $resolvedInputPath
        OutputPath                 = $resolvedOutputPath
        WordName                   = $word.Name
        WordVersion                = $word.Version
        WordPath                   = $word.Path
        MathTypeTemplatePath       = $templatePath
        InitialMathTypeObjects     = $beforeAudit.MathTypeObjects
        FinalMathTypeObjects       = $finalAudit.MathTypeObjects
        AddedMathTypeObjects       = $finalAudit.MathTypeObjects - $beforeAudit.MathTypeObjects
        RemainingInlineLatex       = $finalAudit.RemainingInlineLatex
        RemainingDisplayLatex      = $finalAudit.RemainingDisplayLatex
        DollarMarkersRemoved       = $removedDollarMarkers
        DelimiterParagraphsRemoved = $removedDelimiterParagraphs
        ResidualDisplayAttempts    = $displayAttempts
        ResidualInlineConversions  = $inlineFallbackConversions
        InlineSamples              = $finalAudit.InlineSamples
        DisplaySamples             = $finalAudit.DisplaySamples
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
