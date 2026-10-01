[CmdletBinding()]
param(
    [ValidateSet('Auto', 'Edge', 'Chrome')][string]$Browser = 'Auto',
    [string]$BrowserPath,
    [ValidateRange(1024, 65535)][int]$Port = 9222,
    [string]$UserDataDir,
    [string]$StartUrl = 'about:blank',
    [ValidateRange(3, 120)][int]$StartupTimeoutSeconds = 30
)

$ErrorActionPreference = 'Stop'

function Get-FullDirectoryPath([string]$Value) {
    return [IO.Path]::GetFullPath([Environment]::ExpandEnvironmentVariables($Value)).TrimEnd('\', '/')
}

function Get-CommandOption([string]$CommandLine, [string]$Name) {
    if (-not $CommandLine) { return $null }
    $escaped = [regex]::Escape($Name)
    $pattern = '(?i)(?:^|\s)(?:"--' + $escaped + '=([^"]*)"|--' + $escaped + '=(?:"([^"]*)"|([^\s"]+))|--' + $escaped + '\s+(?:"([^"]*)"|([^\s"]+)))(?=\s|$)'
    $match = [regex]::Match($CommandLine, $pattern)
    if (-not $match.Success) { return $null }
    foreach ($index in 1..5) {
        if ($match.Groups[$index].Success) { return $match.Groups[$index].Value }
    }
    return $null
}

function Get-CdpVersion([string]$Endpoint) {
    $response = $null
    $reader = $null
    try {
        $request = [Net.HttpWebRequest]::Create($Endpoint + '/json/version')
        $request.Proxy = $null
        $request.Timeout = 1500
        $request.ReadWriteTimeout = 1500
        $request.AllowAutoRedirect = $false
        $response = $request.GetResponse()
        $reader = New-Object IO.StreamReader($response.GetResponseStream())
        $buffer = New-Object char[] 65537
        $count = $reader.ReadBlock($buffer, 0, $buffer.Length)
        if ($count -gt 65536) { return $null }
        $version = (-join $buffer[0..($count - 1)]) | ConvertFrom-Json
        if (-not $version.Browser -or -not $version.webSocketDebuggerUrl) { return $null }
        $socketUri = [Uri]$version.webSocketDebuggerUrl
        if ($socketUri.Scheme -ne 'ws' -or $socketUri.Host -notin @('127.0.0.1', 'localhost', '[::1]', '::1') -or $socketUri.Port -ne $Port) { return $null }
        return $version
    }
    catch { return $null }
    finally {
        if ($null -ne $reader) { $reader.Dispose() }
        if ($null -ne $response) { $response.Dispose() }
    }
}

function Test-PortAvailable([int]$Number) {
    $listener = New-Object Net.Sockets.TcpListener([Net.IPAddress]::Loopback, $Number)
    try {
        $listener.ExclusiveAddressUse = $true
        $listener.Start()
        return $true
    }
    catch { return $false }
    finally { $listener.Stop() }
}

if (-not $env:LOCALAPPDATA) { throw 'LOCALAPPDATA is unavailable. Run this launcher on Windows under your normal user account.' }
if (-not $UserDataDir) {
    $UserDataDir = Join-Path $env:LOCALAPPDATA 'CustomerExpressIssueQueryPlaywright\browser-profile'
}
$profilePath = Get-FullDirectoryPath $UserDataDir
if ($profilePath.Contains('"')) { throw 'UserDataDir must not contain quotation marks.' }
if ([IO.Path]::GetPathRoot($profilePath).TrimEnd('\', '/') -eq $profilePath) {
    throw 'UserDataDir must be a dedicated browser profile folder, not a drive root.'
}
$protectedProfiles = @(
    'Microsoft\Edge\User Data', 'Microsoft\Edge Beta\User Data',
    'Microsoft\Edge Dev\User Data', 'Microsoft\Edge SxS\User Data',
    'Google\Chrome\User Data', 'Google\Chrome Beta\User Data',
    'Google\Chrome Dev\User Data', 'Google\Chrome SxS\User Data',
    'Chromium\User Data'
)
foreach ($relativePath in $protectedProfiles) {
    $protectedPath = Get-FullDirectoryPath (Join-Path $env:LOCALAPPDATA $relativePath)
    if ($profilePath.Equals($protectedPath, [StringComparison]::OrdinalIgnoreCase) -or $profilePath.StartsWith($protectedPath + '\', [StringComparison]::OrdinalIgnoreCase)) {
        throw 'Refusing to use a normal Edge/Chrome profile. Choose a dedicated UserDataDir for this skill.'
    }
}
if (Test-Path -LiteralPath $profilePath -PathType Leaf) { throw 'UserDataDir points to a file; choose a dedicated directory.' }
$startUri = $null
if (-not [Uri]::TryCreate($StartUrl, [UriKind]::Absolute, [ref]$startUri) -or $StartUrl.Contains('"') -or $StartUrl -match '[\r\n]' -or ($StartUrl -ne 'about:blank' -and $startUri.Scheme -notin @('http', 'https'))) {
    throw 'StartUrl must be about:blank or an absolute HTTP/HTTPS URL.'
}

$processNames = @('msedge.exe', 'chrome.exe', 'chromium.exe')
if ($BrowserPath) {
    $BrowserPath = [IO.Path]::GetFullPath([Environment]::ExpandEnvironmentVariables($BrowserPath))
    if (-not (Test-Path -LiteralPath $BrowserPath -PathType Leaf) -or [IO.Path]::GetExtension($BrowserPath) -ne '.exe') {
        throw 'BrowserPath must point to an existing browser .exe file.'
    }
    $processNames += [IO.Path]::GetFileName($BrowserPath)
}
try {
    $processes = @(Get-CimInstance -ClassName Win32_Process -ErrorAction Stop | Where-Object { $_.Name -in $processNames })
}
catch { throw 'Cannot inspect running browser processes safely. Check Windows CIM access and try again.' }
$matchingProcess = $null
foreach ($browserProcess in $processes) {
    $runningProfile = Get-CommandOption $browserProcess.CommandLine 'user-data-dir'
    if (-not $runningProfile -or -not [IO.Path]::IsPathRooted($runningProfile)) { continue }
    $runningProfilePath = Get-FullDirectoryPath $runningProfile
    if (-not $profilePath.Equals($runningProfilePath, [StringComparison]::OrdinalIgnoreCase)) { continue }
    # Renderer/utility children may inherit the profile argument. The browser parent owns CDP.
    if ($null -ne (Get-CommandOption $browserProcess.CommandLine 'type')) { continue }
    $runningPort = Get-CommandOption $browserProcess.CommandLine 'remote-debugging-port'
    if ($runningPort -ne [string]$Port) {
        throw 'This profile is already open on another debugging port or without CDP. Reuse its original endpoint, or choose another dedicated UserDataDir; no browser was closed.'
    }
    $runningAddress = Get-CommandOption $browserProcess.CommandLine 'remote-debugging-address'
    if ($runningAddress -and $runningAddress -notin @('127.0.0.1', 'localhost', '::1')) {
        throw 'The existing profile was started with a non-loopback debugging address. Use a dedicated profile with loopback CDP.'
    }
    $matchingProcess = $browserProcess
}

$endpoint = 'http://127.0.0.1:' + $Port
$version = Get-CdpVersion $endpoint
if ($version) {
    if (-not $matchingProcess) {
        throw 'The requested port serves CDP for a different or unverifiable browser profile. Choose another Port; no browser was changed.'
    }
    Write-Output ('Reusing the dedicated browser. CDP endpoint: ' + $endpoint)
    Write-Output ('Profile: ' + $profilePath)
    return
}
if (-not $matchingProcess -and -not (Test-PortAvailable $Port)) {
    throw 'The requested port is occupied and is not a reusable CDP browser. Choose another Port.'
}

if (-not $matchingProcess) {
    if (-not $BrowserPath) {
        $browserOrder = if ($Browser -eq 'Auto') { @('Edge', 'Chrome') } else { @($Browser) }
        $installRoots = @(${env:ProgramFiles(x86)}, $env:ProgramFiles, $env:LOCALAPPDATA) | Where-Object { $_ }
        foreach ($browserName in $browserOrder) {
            $relativeExe = if ($browserName -eq 'Edge') { 'Microsoft\Edge\Application\msedge.exe' } else { 'Google\Chrome\Application\chrome.exe' }
            foreach ($installRoot in $installRoots) {
                $candidate = Join-Path $installRoot $relativeExe
                if (Test-Path -LiteralPath $candidate -PathType Leaf) { $BrowserPath = $candidate; break }
            }
            if ($BrowserPath) { break }
        }
    }
    if (-not $BrowserPath) { throw 'No supported browser was found. Install Edge/Chrome or provide -BrowserPath.' }
    New-Item -ItemType Directory -Force -Path $profilePath | Out-Null
    $browserArguments = @(
        '--remote-debugging-address=127.0.0.1',
        ('--remote-debugging-port=' + $Port),
        ('--user-data-dir="' + $profilePath + '"'),
        '--no-first-run', '--no-default-browser-check', '--new-window',
        ('"' + $StartUrl + '"')
    )
    # Do not stop any existing browser. A separate profile keeps ordinary sessions isolated.
    Start-Process -FilePath $BrowserPath -ArgumentList $browserArguments -WindowStyle Hidden | Out-Null
}

$timer = [Diagnostics.Stopwatch]::StartNew()
do {
    $version = Get-CdpVersion $endpoint
    if ($version) {
        Write-Output ('Browser ready. CDP endpoint: ' + $endpoint)
        Write-Output ('Profile: ' + $profilePath)
        Write-Output 'Log in to JST in this dedicated browser, then run the query using this CDP endpoint.'
        return
    }
    Start-Sleep -Milliseconds 300
} while ($timer.Elapsed.TotalSeconds -lt $StartupTimeoutSeconds)
throw ('Browser CDP did not become ready within ' + $StartupTimeoutSeconds + ' seconds. Inspect the dedicated browser, profile lock, and requested port; no browser was closed.')
