param()
$ErrorActionPreference = 'Stop'
if (-not [Environment]::Is64BitOperatingSystem) { throw 'This package requires Windows x64.' }
$taskPython = Join-Path $PSScriptRoot 'skills\customer-express-issue-query-playwright\runtime\python\python.exe'
if (-not (Test-Path -LiteralPath $taskPython)) { throw 'Bundled Python is missing; extract the complete package first.' }
& $taskPython -B (Join-Path $PSScriptRoot 'verify.py')
if ($LASTEXITCODE -ne 0) { throw 'Portable package self-check failed.' }
