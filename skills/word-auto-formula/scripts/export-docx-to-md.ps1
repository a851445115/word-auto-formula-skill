[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,

    [Parameter(Mandatory = $true)]
    [ValidateSet("formula-preserved", "latex-raw")]
    [string]$Mode,

    [string]$OutputDirectory,

    [switch]$KeepTemp,

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

function New-TempDirectory {
    $path = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString("N"))
    [void][System.IO.Directory]::CreateDirectory($path)
    return $path
}

function Expand-DocxPackage {
    param(
        [Parameter(Mandatory = $true)][string]$SourcePath,
        [Parameter(Mandatory = $true)][string]$DestinationDirectory
    )

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::ExtractToDirectory($SourcePath, $DestinationDirectory)
}

$resolvedInputPath = Resolve-AbsoluteExistingPath -Path $InputPath
$resolvedOutputDirectory = if ($OutputDirectory) {
    Resolve-AbsolutePath -Path $OutputDirectory
} else {
    [System.IO.Path]::GetDirectoryName($resolvedInputPath)
}

[void][System.IO.Directory]::CreateDirectory($resolvedOutputDirectory)

$scriptDirectory = Split-Path -Parent $PSCommandPath
$extractorScript = Join-Path $scriptDirectory "extract-docx-formula-preserved.py"
$assetConverterScript = Join-Path $scriptDirectory "convert-docx-assets-to-png.ps1"
$latexConverterScript = Join-Path $scriptDirectory "convert-docx-mathtype-to-latex.ps1"

$baseName = [System.IO.Path]::GetFileNameWithoutExtension($resolvedInputPath)
$workingDocxPath = $resolvedInputPath
$latexSummary = $null

if ($Mode -eq "formula-preserved") {
    $outputMarkdownPath = Join-Path $resolvedOutputDirectory ("{0}_formula-preserved.md" -f $baseName)
    $assetsDirectory = Join-Path $resolvedOutputDirectory ("{0}_formula-preserved_assets" -f $baseName)
} else {
    $workingDocxPath = Join-Path $resolvedOutputDirectory ("{0}_latex.docx" -f $baseName)
    $outputMarkdownPath = Join-Path $resolvedOutputDirectory ("{0}_latex_raw.md" -f $baseName)
    $assetsDirectory = Join-Path $resolvedOutputDirectory ("{0}_latex_raw_assets" -f $baseName)

    $latexCommand = @(
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-File", $latexConverterScript,
        "-SourcePath", $resolvedInputPath,
        "-DestinationPath", $workingDocxPath
    )

    if ($Visible) {
        $latexCommand += "-Visible"
    }

    $latexJson = & powershell @latexCommand
    if ($LASTEXITCODE -ne 0) {
        throw "convert-docx-mathtype-to-latex.ps1 failed."
    }

    $latexSummary = $latexJson | ConvertFrom-Json
}

$assetsRelativePath = [System.IO.Path]::GetFileName($assetsDirectory)
$tempDirectory = New-TempDirectory

try {
    Expand-DocxPackage -SourcePath $workingDocxPath -DestinationDirectory $tempDirectory

    $pythonArgs = @(
        $extractorScript,
        "--source-dir", $tempDirectory,
        "--output-md", $outputMarkdownPath,
        "--assets-dir", $assetsDirectory,
        "--assets-rel", $assetsRelativePath,
        "--converter-script", $assetConverterScript
    )

    $extractorJson = & python @pythonArgs
    if ($LASTEXITCODE -ne 0) {
        throw "extract-docx-formula-preserved.py failed."
    }
    $extractorSummary = $extractorJson | ConvertFrom-Json

    [pscustomobject]@{
        mode                = $Mode
        input_path          = $resolvedInputPath
        working_docx_path   = $workingDocxPath
        output_markdown     = $outputMarkdownPath
        assets_directory    = $assetsDirectory
        temp_directory      = if ($KeepTemp) { $tempDirectory } else { $null }
        latex_conversion    = $latexSummary
        markdown_extraction = $extractorSummary
    } | ConvertTo-Json -Depth 6
} finally {
    if (-not $KeepTemp -and (Test-Path -LiteralPath $tempDirectory)) {
        Remove-Item -LiteralPath $tempDirectory -Recurse -Force
    }
}
