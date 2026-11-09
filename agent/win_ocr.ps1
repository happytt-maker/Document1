# win_ocr.ps1 -- Windows built-in OCR (Windows.Media.Ocr)
#
# NOTE: keep this file pure ASCII. Windows PowerShell 5.1 reads .ps1 files as
# ANSI (GBK on a Chinese system) unless they carry a BOM, so any non-ASCII
# character inside this file will corrupt the parser. All Chinese text comes
# from the OCR result at runtime, never from this script.
#
# Caller: agent/ocr.py (runs it as a subprocess, reads the JSON output file)
#
# Usage:
#   powershell -NoProfile -ExecutionPolicy Bypass -File win_ocr.ps1 -Path img.png -Out result.json

param(
    [Parameter(Mandatory = $true)][string]$Path,
    [Parameter(Mandatory = $true)][string]$Out,
    [string]$Lang = "zh-Hans-CN"
)

$ErrorActionPreference = "Stop"

try {
    Add-Type -AssemblyName System.Runtime.WindowsRuntime | Out-Null

    $null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime]
    $null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Foundation, ContentType = WindowsRuntime]
    $null = [Windows.Storage.StorageFile, Windows.Foundation, ContentType = WindowsRuntime]
    $null = [Windows.Storage.FileAccessMode, Windows.Foundation, ContentType = WindowsRuntime]

    # WinRT IAsyncOperation -> .NET Task (PS 5.1 has no await)
    $asTask = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
            $_.Name -eq 'AsTask' -and
            $_.GetParameters().Count -eq 1 -and
            $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
        })
    $asTaskGeneric = @($asTask)[0]

    function Await($op, [Type]$type) {
        $m = $asTaskGeneric.MakeGenericMethod($type)
        $task = $m.Invoke($null, @($op))
        $task.Wait(-1) | Out-Null
        $task.Result
    }

    $file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($Path)) ([Windows.Storage.StorageFile])
    $stream = Await ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
    $decoder = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
    $bmp = Await ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])

    # Pick recognizer language: exact match -> any zh-Hans -> user profile
    $engine = $null
    $avail = [Windows.Media.Ocr.OcrEngine]::AvailableRecognizerLanguages
    $want = $avail | Where-Object { $_.LanguageTag -eq $Lang } | Select-Object -First 1
    if ($null -eq $want) {
        $want = $avail | Where-Object { $_.LanguageTag -like "zh-Hans*" } | Select-Object -First 1
    }
    if ($null -eq $want) {
        $want = $avail | Select-Object -First 1
    }
    if ($null -ne $want) {
        $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($want)
    }
    if ($null -eq $engine) {
        $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
    }
    if ($null -eq $engine) {
        throw "no OCR language available (install the Simplified Chinese language pack)"
    }

    $res = Await ($engine.RecognizeAsync($bmp)) ([Windows.Media.Ocr.OcrResult])

    # Emit every word (not Line.Text): Windows OCR joins CJK words with spaces,
    # which would turn "设备验收单" into "设 备 验 收 单". The Python side rebuilds
    # each line from word boxes, using the real horizontal gaps.
    $lines = @()
    foreach ($ln in $res.Lines) {
        $words = @()
        foreach ($wd in $ln.Words) {
            $r = $wd.BoundingRect
            $words += [pscustomobject]@{
                t = [string]$wd.Text
                x = [math]::Round($r.X, 1)
                y = [math]::Round($r.Y, 1)
                w = [math]::Round($r.Width, 1)
                h = [math]::Round($r.Height, 1)
            }
        }
        if ($words.Count -gt 0) {
            $lines += [pscustomobject]@{ words = $words }
        }
    }

    $payload = [pscustomobject]@{
        ok     = $true
        engine = "windows-ocr"
        lang   = [string]$engine.RecognizerLanguage.LanguageTag
        width  = [int]$decoder.PixelWidth
        height = [int]$decoder.PixelHeight
        lines  = $lines
    }
    $payload | ConvertTo-Json -Depth 6 -Compress | Set-Content -Path $Out -Encoding UTF8
}
catch {
    ([pscustomobject]@{ ok = $false; error = [string]$_.Exception.Message }) |
        ConvertTo-Json -Compress | Set-Content -Path $Out -Encoding UTF8
}
