param(
    [string]$ManifestPath = "tmp\pdfs\render_manifest.json",
    [string]$OutputDir = "tmp\pdfs\ocr"
)

$ErrorActionPreference = "Stop"
$Root = (Get-Location).Path
$ManifestFullPath = (Resolve-Path $ManifestPath).Path

function Resolve-ProjectPath($PathValue) {
    if ([System.IO.Path]::IsPathRooted($PathValue)) {
        return $PathValue
    }
    return Join-Path $Root $PathValue
}

$OutputFullPath = Resolve-ProjectPath $OutputDir
New-Item -ItemType Directory -Force $OutputFullPath | Out-Null

Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
$null = [Windows.Storage.FileAccessMode, Windows.Storage, ContentType = WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
$null = [Windows.Graphics.Imaging.SoftwareBitmap, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
$null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime]

$asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
    $_.Name -eq "AsTask" -and
    $_.GetParameters().Count -eq 1 -and
    $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
})[0]

function Await-WinRt($Operation, $ResultType) {
    $asTask = $asTaskGeneric.MakeGenericMethod($ResultType)
    $netTask = $asTask.Invoke($null, @($Operation))
    $netTask.Wait() | Out-Null
    $netTask.Result
}

function Invoke-Ocr($ImagePath) {
    $file = Await-WinRt ([Windows.Storage.StorageFile]::GetFileFromPathAsync($ImagePath)) ([Windows.Storage.StorageFile])
    $stream = Await-WinRt ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
    try {
        $decoder = Await-WinRt ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
        $bitmap = Await-WinRt ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
        $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
        if ($null -eq $engine) {
            throw "Windows OCR engine is unavailable."
        }

        $result = Await-WinRt ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])
        $lines = @()
        foreach ($line in $result.Lines) {
            $words = @()
            foreach ($word in $line.Words) {
                $r = $word.BoundingRect
                $words += [pscustomobject]@{
                    text = $word.Text
                    x = [math]::Round($r.X, 2)
                    y = [math]::Round($r.Y, 2)
                    w = [math]::Round($r.Width, 2)
                    h = [math]::Round($r.Height, 2)
                }
            }

            if ($words.Count -gt 0) {
                $minX = ($words | Measure-Object x -Minimum).Minimum
                $minY = ($words | Measure-Object y -Minimum).Minimum
                $maxX = ($words | ForEach-Object { $_.x + $_.w } | Measure-Object -Maximum).Maximum
                $maxY = ($words | ForEach-Object { $_.y + $_.h } | Measure-Object -Maximum).Maximum
                $lines += [pscustomobject]@{
                    text = ($words.text -join " ")
                    x = [math]::Round($minX, 2)
                    y = [math]::Round($minY, 2)
                    w = [math]::Round($maxX - $minX, 2)
                    h = [math]::Round($maxY - $minY, 2)
                    words = $words
                }
            }
        }

        return [pscustomobject]@{
            text = $result.Text
            lines = $lines
        }
    }
    finally {
        if ($null -ne $stream) {
            $stream.Dispose()
        }
    }
}

$items = Get-Content -Raw $ManifestFullPath | ConvertFrom-Json
$count = 0
foreach ($item in $items) {
    $count += 1
    $imagePath = Resolve-ProjectPath $item.image
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($item.image)
    $outPath = Join-Path $OutputFullPath "$baseName.json"
    Write-Host "OCR $count/$($items.Count): $($item.image)"

    $ocr = Invoke-Ocr $imagePath
    $payload = [pscustomobject]@{
        pdf = $item.pdf
        page = $item.page
        image = $item.image
        width = $item.width
        height = $item.height
        text = $ocr.text
        lines = $ocr.lines
    }
    $payload | ConvertTo-Json -Depth 8 | Set-Content -Encoding UTF8 $outPath
}

Write-Host "OCR complete: $($items.Count) pages."
