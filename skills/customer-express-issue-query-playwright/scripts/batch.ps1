param(
    [Parameter(Mandatory=$true)][string]$WorkDir,
    [string]$Endpoint = 'http://127.0.0.1:9222',
    [ValidateRange(1,60)][int]$BatchSize = 25,
    [ValidateRange(120,60000)][int]$DelayMs = 150,
    [ValidateRange(0,100000)][int]$MaxBatches = 0,
    [ValidateRange(0,10000)][int]$PageIndex = 0,
    [ValidateRange(1000,300000)][int]$RequestTimeoutMs = 30000,
    [switch]$RecheckMissing
)
$ErrorActionPreference = 'Stop'
$taskPython = Join-Path $PSScriptRoot '..\runtime\python\python.exe'
if (-not (Test-Path -LiteralPath $taskPython)) { throw 'Extract the complete portable package: bundled Python is missing.' }
$taskArgs = @('-B', (Join-Path $PSScriptRoot 'cdp_query.py'), 'query', '--workdir', $WorkDir, '--endpoint', $Endpoint,
    '--batch-size', $BatchSize, '--delay-ms', $DelayMs, '--max-batches', $MaxBatches, '--request-timeout-ms', $RequestTimeoutMs)
if ($PageIndex -gt 0) { $taskArgs += @('--page-index', $PageIndex) }
if ($RecheckMissing) { $taskArgs += '--recheck-missing' }
& $taskPython @taskArgs
exit $LASTEXITCODE
