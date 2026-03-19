[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SourcePath,

    [Parameter(Mandatory = $true)]
    [string]$DestinationPath,

    [switch]$Visible
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
        throw "Word.Application resolved to '$name' at '$path'. Repair Word COM registration before converting MathType formulas to TeX."
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

$resolvedSourcePath = Resolve-AbsoluteExistingPath -Path $SourcePath
$resolvedDestinationPath = Resolve-AbsolutePath -Path $DestinationPath
$destinationDirectory = [System.IO.Path]::GetDirectoryName($resolvedDestinationPath)

if ($destinationDirectory) {
    [void][System.IO.Directory]::CreateDirectory($destinationDirectory)
}

Copy-Item -LiteralPath $resolvedSourcePath -Destination $resolvedDestinationPath -Force

$word = $null
$document = $null

try {
    $word = New-WordApplication -ShowWindow:$Visible
    $document = $word.Documents.Open($resolvedDestinationPath, $false, $false)
    $document.Activate()

    $initialInlineShapes = $document.InlineShapes.Count
    $converted = 0
    $skipped = 0
    $failed = 0

    for ($i = $initialInlineShapes; $i -ge 1; $i--) {
        $shape = $document.InlineShapes.Item($i)

        if (-not (Test-MathTypeInlineShape -InlineShape $shape)) {
            $skipped++
            continue
        }

        try {
            $shape.Range.Select()
            $word.Run("MathTypeCommands.UILib.MTCommand_TeXToggle")
            $converted++
        } catch {
            $failed++
        }
    }

    $document.Save()

    [pscustomobject]@{
        source                  = $resolvedSourcePath
        destination             = $resolvedDestinationPath
        initial_inline_shapes   = $initialInlineShapes
        remaining_inline_shapes = $document.InlineShapes.Count
        converted               = $converted
        skipped                 = $skipped
        failed                  = $failed
    } | ConvertTo-Json -Compress
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
