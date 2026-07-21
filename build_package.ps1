$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$package = Join-Path $root 'dist\fastgpt-file-preprocessor'
$zip = Join-Path $root 'dist\fastgpt-file-preprocessor.zip'
Remove-Item -Recurse -Force $package -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force $package | Out-Null
Copy-Item "$root\src", "$root\requirements.txt", "$root\run.ps1", "$root\README.md", "$root\ARCHITECTURE.md", "$root\config.json", "$root\config.example.json" -Destination $package -Recurse
Get-ChildItem -LiteralPath "$package\src" -Directory -Recurse -Filter '__pycache__' | Remove-Item -Recurse -Force
Remove-Item $zip -Force -ErrorAction SilentlyContinue
Compress-Archive -Path $package -DestinationPath $zip
Write-Host "部署包：$zip"
