param(
    [Parameter(Mandatory=$true)][string]$Dir,
    [Parameter(Mandatory=$true)][string]$Out,
    [double]$Scale = 2.0
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Media.Ocr.OcrEngine,Windows.Foundation,ContentType=WindowsRuntime]
$null = [Windows.Storage.StorageFile,Windows.Storage,ContentType=WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapDecoder,Windows.Graphics.Imaging,ContentType=WindowsRuntime]
$null = [Windows.Storage.Streams.IRandomAccessStream,Windows.Storage.Streams,ContentType=WindowsRuntime]

$asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
    $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and
    $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]

function Await($WinRtTask, $ResultType) {
    $asTask = $asTaskGeneric.MakeGenericMethod($ResultType)
    $netTask = $asTask.Invoke($null, @($WinRtTask))
    $netTask.Wait(-1) | Out-Null
    $netTask.Result
}

$null = [Windows.Globalization.Language,Windows.Globalization,ContentType=WindowsRuntime]
$engine = $null
try {
    $lang = New-Object Windows.Globalization.Language 'zh-Hans-CN'
    $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($lang)
} catch { }
if ($engine -eq $null) { $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages() }
if ($engine -eq $null) { throw 'no OCR engine' }

Add-Type -AssemblyName System.Drawing

$results = @{}
$files = Get-ChildItem -Path $Dir -Filter *.png | Sort-Object Name
foreach ($f in $files) {
    $src = [System.Drawing.Bitmap]::FromFile($f.FullName)
    $w = [int]($src.Width * $Scale)
    $h = [int]($src.Height * $Scale)
    $big = New-Object System.Drawing.Bitmap($w, $h)
    $g = [System.Drawing.Graphics]::FromImage($big)
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.DrawImage($src, 0, 0, $w, $h)
    $g.Dispose()
    $tmp = Join-Path $env:TEMP ("ocr_" + $f.BaseName + ".png")
    $big.Save($tmp, [System.Drawing.Imaging.ImageFormat]::Png)
    $big.Dispose(); $src.Dispose()

    $file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($tmp)) ([Windows.Storage.StorageFile])
    $stream = Await ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
    $decoder = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
    $bitmap = Await ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
    $res = Await ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])

    $lines = @()
    foreach ($ln in $res.Lines) {
        $words = @()
        foreach ($wd in $ln.Words) {
            $r = $wd.BoundingRect
            $words += [ordered]@{
                t = $wd.Text
                x = [math]::Round($r.X / $Scale, 1)
                y = [math]::Round($r.Y / $Scale, 1)
                w = [math]::Round($r.Width / $Scale, 1)
                h = [math]::Round($r.Height / $Scale, 1)
            }
        }
        $lines += ,@($words)
    }
    $results[$f.Name] = [ordered]@{ text = $res.Text; lines = $lines }
    $bitmap.Dispose(); $stream.Dispose()
    Remove-Item $tmp -Force
}
$json = $results | ConvertTo-Json -Depth 8 -Compress
[System.IO.File]::WriteAllText($Out, $json, (New-Object System.Text.UTF8Encoding($false)))
Write-Output ("ocr done: " + $files.Count + " images -> " + $Out)
