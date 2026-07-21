param(
  [Parameter(Mandatory=$true)][string]$Input,
  [Parameter(Mandatory=$true)][string]$Output
)
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Push-Location $root
try { & "$root\.venv\Scripts\python.exe" -m src.preprocessor --input $Input --output $Output --config "$root\config.json" }
finally { Pop-Location }
