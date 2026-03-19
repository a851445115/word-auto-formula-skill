param(
    [Parameter(Mandatory = $true)]
    [string]$ManifestPath
)

$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Drawing

$entries = Get-Content -Raw -Encoding UTF8 -LiteralPath $ManifestPath | ConvertFrom-Json

foreach ($entry in $entries) {
    $src = [string]$entry.src
    $dst = [string]$entry.dst

    $dstDir = Split-Path -Parent $dst
    if (-not (Test-Path -LiteralPath $dstDir)) {
        New-Item -ItemType Directory -Force -Path $dstDir | Out-Null
    }

    $image = $null
    $bitmap = $null
    $graphics = $null

    try {
        $image = [System.Drawing.Image]::FromFile($src)
        $bitmap = New-Object System.Drawing.Bitmap $image.Width, $image.Height
        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
        $graphics.Clear([System.Drawing.Color]::White)
        $graphics.DrawImage($image, 0, 0, $image.Width, $image.Height)
        $bitmap.Save($dst, [System.Drawing.Imaging.ImageFormat]::Png)
    } finally {
        if ($graphics) { $graphics.Dispose() }
        if ($bitmap) { $bitmap.Dispose() }
        if ($image) { $image.Dispose() }
    }
}
