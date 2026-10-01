[CmdletBinding(DefaultParameterSetName = 'Workspace')]
param(
    [Parameter(ParameterSetName = 'Workspace')][string]$Workspace,
    [Parameter(Mandatory = $true, ParameterSetName = 'SkillsDir')][string]$SkillsDir
)

$ErrorActionPreference = 'Stop'
$skillName = 'customer-express-issue-query-playwright'
$source = Join-Path $PSScriptRoot ('skills\' + $skillName)
if (-not (Test-Path -LiteralPath (Join-Path $source 'SKILL.md') -PathType Leaf)) {
    throw 'The packaged skill is incomplete: SKILL.md was not found.'
}
if (-not (Test-Path -LiteralPath (Join-Path $source 'runtime\python\python.exe') -PathType Leaf)) {
    throw 'The bundled Python runtime was not found. Extract the complete portable package before installing.'
}
if ($PSCmdlet.ParameterSetName -eq 'Workspace') {
    if (-not $Workspace) { $Workspace = Read-Host 'Enter the existing DSH workspace directory' }
    if (-not (Test-Path -LiteralPath $Workspace -PathType Container)) {
        throw 'Workspace must be an existing directory. For Codex or another tool, use -SkillsDir instead.'
    }
    $workspacePath = (Resolve-Path -LiteralPath $Workspace).ProviderPath
    $targetRoot = Join-Path $workspacePath '.dsh\skills'
}
else {
    if ([string]::IsNullOrWhiteSpace($SkillsDir)) { throw 'SkillsDir must not be empty.' }
    $targetRoot = [IO.Path]::GetFullPath([Environment]::ExpandEnvironmentVariables($SkillsDir))
}
$targetRoot = [IO.Path]::GetFullPath($targetRoot).TrimEnd('\', '/')
$target = Join-Path $targetRoot $skillName
if (Test-Path -LiteralPath $target) {
    throw ('The destination already exists; installation never overwrites or merges skills: ' + $target)
}
if (Test-Path -LiteralPath $targetRoot -PathType Leaf) { throw 'The skills destination is a file, not a directory.' }
New-Item -ItemType Directory -Force -Path $targetRoot | Out-Null
$stage = Join-Path $targetRoot ('.' + $skillName + '.install-' + [Guid]::NewGuid().ToString('N'))
try {
    # Copy the one skill in full, including its private portable runtime.
    Copy-Item -LiteralPath $source -Destination $stage -Recurse -ErrorAction Stop
    # Atomic rename fails if the destination appeared meanwhile; it never merges directories.
    [IO.Directory]::Move($stage, $target)
}
catch {
    if (Test-Path -LiteralPath $stage -PathType Container) {
        $resolvedStage = [IO.Path]::GetFullPath((Resolve-Path -LiteralPath $stage).ProviderPath)
        if ($resolvedStage.StartsWith($targetRoot + '\', [StringComparison]::OrdinalIgnoreCase) -and [IO.Path]::GetFileName($resolvedStage) -eq [IO.Path]::GetFileName($stage)) {
            Remove-Item -LiteralPath $resolvedStage -Recurse -Force
        }
    }
    throw
}
Write-Output ('Installed one skill with its bundled runtime: ' + $target)
Write-Output 'Reload the skill list in your AI tool. Start the dedicated CDP browser and log in to JST before querying.'
