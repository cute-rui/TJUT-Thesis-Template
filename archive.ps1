$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$dir = Join-Path $PSScriptRoot "archive\v_$timestamp"
New-Item -ItemType Directory -Path $dir -Force | Out-Null
Copy-Item (Join-Path $PSScriptRoot "generate_thesis.js") (Join-Path $dir "generate_thesis.js")
Write-Host "已归档至: $dir"
