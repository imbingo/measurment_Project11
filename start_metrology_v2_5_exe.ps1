$script = Join-Path $PSScriptRoot "start_metrology_v2_4_exe.ps1"
if (-not (Test-Path -LiteralPath $script)) {
    throw "Missing compatible startup script: $script"
}
& $script @args
exit $LASTEXITCODE
