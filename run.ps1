param(
  [Parameter(Mandatory=$true)][string]$Input,
  [Parameter(Mandatory=$true)][string]$Output
)
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
& "$root\.venv\Scripts\python.exe" "$root\src\preprocessor.py" --input $Input --output $Output
