$ErrorActionPreference = "Stop"
Set-Location -LiteralPath $PSScriptRoot

python .\metrology_data_platform_v2_6.py @args
