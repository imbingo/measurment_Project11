$script = Join-Path $PSScriptRoot "start_metrology_v2_4_ocr.ps1"
if (-not (Test-Path -LiteralPath $script)) {
    throw "Missing compatible OCR startup script: $script"
}
& $script @args
exit $LASTEXITCODE
