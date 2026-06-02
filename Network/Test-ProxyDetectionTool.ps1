<#
.SYNOPSIS
    Detects classic proxy configuration sources and reports origin attribution.

.DESCRIPTION
    Collects proxy signals from WinINet (user and machine-default), WinHTTP,
    policy-based registry keys (GPO), WPAD DNS and PAC reachability.
    Produces console output, TXT log and JSON report with confidence scoring
    and source attribution.

.NOTES
    Designed for troubleshooting and auditing. It provides high practical
    coverage for classic proxy paths, but cannot guarantee 100% detection in
    environments using opaque network agents.
#>

[CmdletBinding()]
param(
    [string]$LogPath = ".\Test-ProxyDetectionTool.log",
    [string]$JsonPath = ".\Test-ProxyDetectionTool.json",
    [int]$TimeoutSec = 6,
    [switch]$RawEvidence
)

$ErrorActionPreference = "SilentlyContinue"

function Get-TimestampedPath {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Stamp
    )

    $directory = Split-Path -Path $Path -Parent
    if ([string]::IsNullOrWhiteSpace($directory)) {
        $directory = "."
    }

    $fileName = Split-Path -Path $Path -Leaf
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
    $extension = [System.IO.Path]::GetExtension($fileName)

    if ([string]::IsNullOrWhiteSpace($baseName)) {
        $baseName = "output"
    }

    $timestampedName = if ([string]::IsNullOrWhiteSpace($extension)) {
        "{0}-{1}" -f $baseName, $Stamp
    }
    else {
        "{0}-{1}{2}" -f $baseName, $Stamp, $extension
    }

    return Join-Path -Path $directory -ChildPath $timestampedName
}

function Write-Log {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [ValidateSet("INFO", "WARN", "ERROR", "DEBUG")][string]$Level = "INFO"
    )

    $timestamp = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss.fff")
    Add-Content -Path $LogPath -Value "[$timestamp] [$Level] $Message" -ErrorAction SilentlyContinue
}

function New-SourceEvidence {
    param(
        [string]$SourceType,
        [string]$Context,
        [string]$Signal,
        [string]$EffectiveValue,
        [string]$SourcePathOrCommand,
        [int]$Confidence,
        [bool]$IsEffective
    )

    [pscustomobject]@{
        SourceType          = $SourceType
        Context             = $Context
        Signal              = $Signal
        EffectiveValue      = $EffectiveValue
        SourcePathOrCommand = $SourcePathOrCommand
        Confidence          = $Confidence
        IsEffective         = $IsEffective
    }
}

function Get-RegistryProperties {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        return $null
    }

    return Get-ItemProperty -LiteralPath $Path -ErrorAction SilentlyContinue
}

function Get-WinInetContext {
    param(
        [Parameter(Mandatory = $true)][string]$RegistryPath,
        [Parameter(Mandatory = $true)][string]$ContextName
    )

    $p = Get-RegistryProperties -Path $RegistryPath
    if ($null -eq $p) {
        return [pscustomobject]@{
            Context       = $ContextName
            Path          = $RegistryPath
            Exists        = $false
            ProxyEnable   = $null
            ProxyServer   = $null
            ProxyOverride = $null
            AutoConfigURL = $null
            AutoDetect    = $null
        }
    }

    [pscustomobject]@{
        Context       = $ContextName
        Path          = $RegistryPath
        Exists        = $true
        ProxyEnable   = $p.ProxyEnable
        ProxyServer   = $p.ProxyServer
        ProxyOverride = $p.ProxyOverride
        AutoConfigURL = $p.AutoConfigURL
        AutoDetect    = $p.AutoDetect
    }
}

function Get-PolicyProxyContext {
    param(
        [Parameter(Mandatory = $true)][string]$RegistryPath,
        [Parameter(Mandatory = $true)][string]$ContextName
    )

    $p = Get-RegistryProperties -Path $RegistryPath
    if ($null -eq $p) {
        return [pscustomobject]@{
            Context              = $ContextName
            Path                 = $RegistryPath
            Exists               = $false
            ProxyEnable          = $null
            ProxyServer          = $null
            ProxyOverride        = $null
            AutoConfigURL        = $null
            AutoDetect           = $null
            ProxySettingsPerUser = $null
        }
    }

    [pscustomobject]@{
        Context              = $ContextName
        Path                 = $RegistryPath
        Exists               = $true
        ProxyEnable          = $p.ProxyEnable
        ProxyServer          = $p.ProxyServer
        ProxyOverride        = $p.ProxyOverride
        AutoConfigURL        = $p.AutoConfigURL
        AutoDetect           = $p.AutoDetect
        ProxySettingsPerUser = $p.ProxySettingsPerUser
    }
}

function Get-WinHTTP {
    $showProxyOutput = (netsh winhttp show proxy) | Out-String
    $showAdvOutput = (netsh winhttp show advproxy) | Out-String

    $proxyServer = $null
    $bypass = $null
    $autoConfigUrl = $null
    $autoDetect = $null

    if ($showProxyOutput -match "Direct access \(no proxy server\)") {
        $proxyServer = "NoProxy"
    }
    if ($showProxyOutput -match "Proxy Server\(s\)\s*:\s*(.+)") {
        $proxyServer = $Matches[1].Trim()
    }
    if ($showProxyOutput -match "Bypass List\s*:\s*(.+)") {
        $bypass = $Matches[1].Trim()
    }

    if ($showAdvOutput -match "AutoDetect\s*:\s*(.+)") {
        $autoDetect = $Matches[1].Trim()
    }
    if ($showAdvOutput -match "AutoConfigUrl\s*:\s*(.+)") {
        $autoConfigUrl = $Matches[1].Trim()
    }

    [pscustomobject]@{
        ProxyServer      = $proxyServer
        ProxyBypass      = $bypass
        AutoDetect       = $autoDetect
        AutoConfigURL    = $autoConfigUrl
        RawShowProxy     = $showProxyOutput.Trim()
        RawShowAdvProxy  = $showAdvOutput.Trim()
        SourceCommand    = "netsh winhttp show proxy | netsh winhttp show advproxy"
    }
}

function Get-DnsSuffixes {
    $suffixes = @()
    $cimItems = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration
    foreach ($item in $cimItems) {
        if ($item.DNSDomainSuffixSearchOrder) {
            $suffixes += $item.DNSDomainSuffixSearchOrder
        }
        if ($item.DNSDomain) {
            $suffixes += $item.DNSDomain
        }
    }

    return $suffixes | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique
}

function Test-WPADDns {
    param([string[]]$Suffixes)

    $results = @()
    foreach ($suffix in $Suffixes) {
        $wpadHost = "wpad.$suffix"
        $resolved = $false
        $addresses = @()
        $errorMessage = $null

        try {
            $dns = Resolve-DnsName -Name $wpadHost -ErrorAction Stop
            $addresses = $dns | Where-Object { $_.IPAddress } | Select-Object -ExpandProperty IPAddress
            if ($addresses.Count -gt 0) {
                $resolved = $true
            }
        }
        catch {
            $errorMessage = $_.Exception.Message
        }

        $results += [pscustomobject]@{
            Hostname     = $wpadHost
            Resolved     = $resolved
            IPAddresses  = @($addresses)
            ErrorMessage = $errorMessage
        }
    }

    return $results
}

function Test-PacUrl {
    param(
        [Parameter(Mandatory = $true)][string]$Url,
        [Parameter(Mandatory = $true)][int]$TimeoutSeconds
    )

    $start = Get-Date
    try {
        $resp = Invoke-WebRequest -Uri $Url -Method Get -TimeoutSec $TimeoutSeconds -UseBasicParsing -ErrorAction Stop
        $elapsed = [int]((Get-Date) - $start).TotalMilliseconds
        $content = [string]$resp.Content
        $hasPacFunction = $content -match "FindProxyForURL"

        return [pscustomobject]@{
            Url              = $Url
            Reachable        = $true
            HttpStatus       = [int]$resp.StatusCode
            LatencyMs        = $elapsed
            ContainsPacLogic = $hasPacFunction
            ErrorMessage     = $null
        }
    }
    catch {
        $elapsed = [int]((Get-Date) - $start).TotalMilliseconds
        return [pscustomobject]@{
            Url              = $Url
            Reachable        = $false
            HttpStatus       = $null
            LatencyMs        = $elapsed
            ContainsPacLogic = $false
            ErrorMessage     = $_.Exception.Message
        }
    }
}

function Get-ExecutionContextName {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)

    if ($identity.Name -match "NT AUTHORITY\\SYSTEM") {
        return "SYSTEM"
    }

    if ($principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        return "UserAdmin"
    }

    return "User"
}

function Add-SourceIfPresent {
    param(
        [ref]$Sources,
        [string]$SourceType,
        [string]$Context,
        [string]$Signal,
        [object]$Value,
        [string]$PathOrCommand,
        [int]$Confidence,
        [bool]$IsEffective
    )

    if ($null -eq $Value) {
        return
    }

    if ($Value -is [string] -and [string]::IsNullOrWhiteSpace($Value)) {
        return
    }

    # Disabled flags should not be counted as detected proxy evidence.
    if ($Signal -in @("ProxyEnable", "AutoDetect")) {
        $intValue = $null
        if ([int]::TryParse([string]$Value, [ref]$intValue) -and $intValue -eq 0) {
            return
        }
    }

    $Sources.Value += New-SourceEvidence -SourceType $SourceType -Context $Context -Signal $Signal -EffectiveValue ([string]$Value) -SourcePathOrCommand $PathOrCommand -Confidence $Confidence -IsEffective $IsEffective
}

function Get-SummaryResult {
    param(
        [object[]]$Sources,
        [object[]]$PacResults,
        [object[]]$WpadDnsResults
    )

    $maxConfidence = 0
    if ($Sources.Count -gt 0) {
        $maxConfidence = ($Sources | Measure-Object -Property Confidence -Maximum).Maximum
    }

    $wpadResolved = ($WpadDnsResults | Where-Object { $_.Resolved }).Count -gt 0
    $pacValid = ($PacResults | Where-Object { $_.Reachable -and $_.ContainsPacLogic }).Count -gt 0

    if ($pacValid -and $maxConfidence -lt 92) {
        $maxConfidence = 92
    }
    elseif ($wpadResolved -and $maxConfidence -lt 65) {
        $maxConfidence = 65
    }

    $status = "Not detected"
    if ($maxConfidence -ge 80) {
        $status = "Detected"
    }
    elseif ($maxConfidence -ge 45) {
        $status = "Configured (verify)"
    }

    $primary = $null
    if ($Sources.Count -gt 0) {
        $primary = ($Sources | Sort-Object -Property Confidence -Descending | Select-Object -First 1)
    }

    return [pscustomobject]@{
        Status               = $status
        Detected             = ($status -ne "Not detected")
        OverallConfidence    = $maxConfidence
        PrimaryDetectionPath = if ($null -ne $primary) { "$($primary.SourceType): $($primary.Signal)" } else { "None" }
        SourceCount          = $Sources.Count
        WpadDnsResolved      = $wpadResolved
        PacValidated         = $pacValid
    }
}

# Initialize output files
$runStamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
$LogPath = Get-TimestampedPath -Path $LogPath -Stamp $runStamp
$JsonPath = Get-TimestampedPath -Path $JsonPath -Stamp $runStamp

$logDir = Split-Path -Path $LogPath -Parent
if (-not [string]::IsNullOrWhiteSpace($logDir) -and -not (Test-Path -LiteralPath $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}

$jsonDir = Split-Path -Path $JsonPath -Parent
if (-not [string]::IsNullOrWhiteSpace($jsonDir) -and -not (Test-Path -LiteralPath $jsonDir)) {
    New-Item -ItemType Directory -Path $jsonDir -Force | Out-Null
}

Remove-Item -LiteralPath $LogPath -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath $JsonPath -Force -ErrorAction SilentlyContinue
New-Item -ItemType File -Path $LogPath -Force | Out-Null

Write-Log -Message "Proxy detection script started"
$runContext = Get-ExecutionContextName
Write-Log -Message "ExecutionContext: $runContext"

$sources = @()
$notes = @()

# Collect local and policy settings
$winInetUserPath = "Registry::HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings"
$winInetMachineDefaultPath = "Registry::HKEY_USERS\.DEFAULT\Software\Microsoft\Windows\CurrentVersion\Internet Settings"
$policyUserPath = "Registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings"
$policyMachinePath = "Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings"
$winHttpSettingsPath = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\Connections"

$winInetUser = Get-WinInetContext -RegistryPath $winInetUserPath -ContextName "User"
$winInetMachineDefault = Get-WinInetContext -RegistryPath $winInetMachineDefaultPath -ContextName "MachineDefault"
$policyUser = Get-PolicyProxyContext -RegistryPath $policyUserPath -ContextName "UserPolicy"
$policyMachine = Get-PolicyProxyContext -RegistryPath $policyMachinePath -ContextName "MachinePolicy"
$winHttp = Get-WinHTTP

# Attribution precedence: policy values are considered effective over local for the same context.
$hasUserPolicyProxy = ($policyUser.Exists -and (($null -ne $policyUser.ProxyEnable) -or $policyUser.ProxyServer -or $policyUser.AutoConfigURL))
$hasMachinePolicyProxy = ($policyMachine.Exists -and (($null -ne $policyMachine.ProxyEnable) -or $policyMachine.ProxyServer -or $policyMachine.AutoConfigURL))

Add-SourceIfPresent -Sources ([ref]$sources) -SourceType "RegistryLocal" -Context "User" -Signal "ProxyEnable" -Value $winInetUser.ProxyEnable -PathOrCommand $winInetUser.Path -Confidence 72 -IsEffective (-not $hasUserPolicyProxy)
Add-SourceIfPresent -Sources ([ref]$sources) -SourceType "RegistryLocal" -Context "User" -Signal "ProxyServer" -Value $winInetUser.ProxyServer -PathOrCommand $winInetUser.Path -Confidence 78 -IsEffective (-not $hasUserPolicyProxy)
Add-SourceIfPresent -Sources ([ref]$sources) -SourceType "RegistryLocal" -Context "User" -Signal "ProxyOverride" -Value $winInetUser.ProxyOverride -PathOrCommand $winInetUser.Path -Confidence 58 -IsEffective (-not $hasUserPolicyProxy)
Add-SourceIfPresent -Sources ([ref]$sources) -SourceType "RegistryLocal" -Context "User" -Signal "AutoConfigURL" -Value $winInetUser.AutoConfigURL -PathOrCommand $winInetUser.Path -Confidence 82 -IsEffective (-not $hasUserPolicyProxy)
Add-SourceIfPresent -Sources ([ref]$sources) -SourceType "RegistryLocal" -Context "User" -Signal "AutoDetect" -Value $winInetUser.AutoDetect -PathOrCommand $winInetUser.Path -Confidence 62 -IsEffective (-not $hasUserPolicyProxy)

Add-SourceIfPresent -Sources ([ref]$sources) -SourceType "RegistryLocal" -Context "MachineDefault" -Signal "ProxyEnable" -Value $winInetMachineDefault.ProxyEnable -PathOrCommand $winInetMachineDefault.Path -Confidence 68 -IsEffective (-not $hasMachinePolicyProxy)
Add-SourceIfPresent -Sources ([ref]$sources) -SourceType "RegistryLocal" -Context "MachineDefault" -Signal "ProxyServer" -Value $winInetMachineDefault.ProxyServer -PathOrCommand $winInetMachineDefault.Path -Confidence 74 -IsEffective (-not $hasMachinePolicyProxy)
Add-SourceIfPresent -Sources ([ref]$sources) -SourceType "RegistryLocal" -Context "MachineDefault" -Signal "AutoConfigURL" -Value $winInetMachineDefault.AutoConfigURL -PathOrCommand $winInetMachineDefault.Path -Confidence 80 -IsEffective (-not $hasMachinePolicyProxy)
Add-SourceIfPresent -Sources ([ref]$sources) -SourceType "RegistryLocal" -Context "MachineDefault" -Signal "AutoDetect" -Value $winInetMachineDefault.AutoDetect -PathOrCommand $winInetMachineDefault.Path -Confidence 60 -IsEffective (-not $hasMachinePolicyProxy)

Add-SourceIfPresent -Sources ([ref]$sources) -SourceType "GPO" -Context "UserPolicy" -Signal "ProxyEnable" -Value $policyUser.ProxyEnable -PathOrCommand $policyUser.Path -Confidence 86 -IsEffective $true
Add-SourceIfPresent -Sources ([ref]$sources) -SourceType "GPO" -Context "UserPolicy" -Signal "ProxyServer" -Value $policyUser.ProxyServer -PathOrCommand $policyUser.Path -Confidence 90 -IsEffective $true
Add-SourceIfPresent -Sources ([ref]$sources) -SourceType "GPO" -Context "UserPolicy" -Signal "ProxyOverride" -Value $policyUser.ProxyOverride -PathOrCommand $policyUser.Path -Confidence 82 -IsEffective $true
Add-SourceIfPresent -Sources ([ref]$sources) -SourceType "GPO" -Context "UserPolicy" -Signal "AutoConfigURL" -Value $policyUser.AutoConfigURL -PathOrCommand $policyUser.Path -Confidence 92 -IsEffective $true
Add-SourceIfPresent -Sources ([ref]$sources) -SourceType "GPO" -Context "UserPolicy" -Signal "AutoDetect" -Value $policyUser.AutoDetect -PathOrCommand $policyUser.Path -Confidence 84 -IsEffective $true
Add-SourceIfPresent -Sources ([ref]$sources) -SourceType "GPO" -Context "MachinePolicy" -Signal "ProxySettingsPerUser" -Value $policyMachine.ProxySettingsPerUser -PathOrCommand $policyMachine.Path -Confidence 88 -IsEffective $true
Add-SourceIfPresent -Sources ([ref]$sources) -SourceType "GPO" -Context "MachinePolicy" -Signal "ProxyEnable" -Value $policyMachine.ProxyEnable -PathOrCommand $policyMachine.Path -Confidence 86 -IsEffective $true
Add-SourceIfPresent -Sources ([ref]$sources) -SourceType "GPO" -Context "MachinePolicy" -Signal "ProxyServer" -Value $policyMachine.ProxyServer -PathOrCommand $policyMachine.Path -Confidence 90 -IsEffective $true
Add-SourceIfPresent -Sources ([ref]$sources) -SourceType "GPO" -Context "MachinePolicy" -Signal "AutoConfigURL" -Value $policyMachine.AutoConfigURL -PathOrCommand $policyMachine.Path -Confidence 92 -IsEffective $true
Add-SourceIfPresent -Sources ([ref]$sources) -SourceType "GPO" -Context "MachinePolicy" -Signal "AutoDetect" -Value $policyMachine.AutoDetect -PathOrCommand $policyMachine.Path -Confidence 84 -IsEffective $true

if ($winHttp.ProxyServer -and $winHttp.ProxyServer -ne "NoProxy") {
    Add-SourceIfPresent -Sources ([ref]$sources) -SourceType "WinHTTP" -Context "SYSTEM" -Signal "ProxyServer" -Value $winHttp.ProxyServer -PathOrCommand $winHttp.SourceCommand -Confidence 90 -IsEffective $true
}
if ($winHttp.ProxyBypass) {
    Add-SourceIfPresent -Sources ([ref]$sources) -SourceType "WinHTTP" -Context "SYSTEM" -Signal "BypassList" -Value $winHttp.ProxyBypass -PathOrCommand $winHttp.SourceCommand -Confidence 74 -IsEffective $true
}
if ($winHttp.AutoConfigURL) {
    Add-SourceIfPresent -Sources ([ref]$sources) -SourceType "WinHTTP" -Context "SYSTEM" -Signal "AutoConfigURL" -Value $winHttp.AutoConfigURL -PathOrCommand $winHttp.SourceCommand -Confidence 88 -IsEffective $true
}
if ($winHttp.AutoDetect) {
    Add-SourceIfPresent -Sources ([ref]$sources) -SourceType "WinHTTP" -Context "SYSTEM" -Signal "AutoDetect" -Value $winHttp.AutoDetect -PathOrCommand $winHttp.SourceCommand -Confidence 76 -IsEffective $true
}

$winHttpConnections = Get-RegistryProperties -Path $winHttpSettingsPath
if ($null -ne $winHttpConnections -and $winHttpConnections.WinHttpSettings) {
    $notes += "WinHttpSettings binary value exists in machine registry."
}

# WPAD DNS
$suffixes = Get-DnsSuffixes
$wpadDnsResults = Test-WPADDns -Suffixes $suffixes
foreach ($item in $wpadDnsResults) {
    if ($item.Resolved) {
        $sources += New-SourceEvidence -SourceType "WPAD" -Context "Network" -Signal "DNSResolved" -EffectiveValue ($item.Hostname + " -> " + (($item.IPAddresses -join ","))) -SourcePathOrCommand "Resolve-DnsName $($item.Hostname)" -Confidence 65 -IsEffective $false
    }
}

# PAC validation candidates
$pacCandidates = @()
if ($winInetUser.AutoConfigURL) { $pacCandidates += $winInetUser.AutoConfigURL }
if ($winInetMachineDefault.AutoConfigURL) { $pacCandidates += $winInetMachineDefault.AutoConfigURL }
if ($policyUser.AutoConfigURL) { $pacCandidates += $policyUser.AutoConfigURL }
if ($policyMachine.AutoConfigURL) { $pacCandidates += $policyMachine.AutoConfigURL }
if ($winHttp.AutoConfigURL) { $pacCandidates += $winHttp.AutoConfigURL }
foreach ($resolvedHost in ($wpadDnsResults | Where-Object { $_.Resolved })) {
    $pacCandidates += ("http://{0}/wpad.dat" -f $resolvedHost.Hostname)
}
$pacCandidates = $pacCandidates | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique

$pacResults = @()
foreach ($candidate in $pacCandidates) {
    $pacResult = Test-PacUrl -Url $candidate -TimeoutSeconds $TimeoutSec
    $pacResults += $pacResult

    if ($pacResult.Reachable -and $pacResult.ContainsPacLogic) {
        $sources += New-SourceEvidence -SourceType "PAC" -Context "Network" -Signal "PACValidated" -EffectiveValue $candidate -SourcePathOrCommand "Invoke-WebRequest $candidate" -Confidence 95 -IsEffective $true
    }
    elseif ($pacResult.Reachable) {
        $sources += New-SourceEvidence -SourceType "PAC" -Context "Network" -Signal "PACReachableNoLogic" -EffectiveValue $candidate -SourcePathOrCommand "Invoke-WebRequest $candidate" -Confidence 55 -IsEffective $false
    }
}

if ($sources.Count -eq 0) {
    $notes += "No explicit proxy signals were detected from registry, policy, WinHTTP, WPAD or PAC."
}
if ($hasUserPolicyProxy -or $hasMachinePolicyProxy) {
    $notes += "Policy-based proxy settings were found and treated as effective in precedence logic."
}
if (($wpadDnsResults | Where-Object { $_.Resolved }).Count -gt 0 -and ($pacResults | Where-Object { $_.Reachable -and $_.ContainsPacLogic }).Count -eq 0) {
    $notes += "WPAD DNS was found but no PAC URL was validated."
}
if ($runContext -ne "SYSTEM") {
    $notes += "Script is not running as SYSTEM. Some machine-context behavior may differ."
}

$summary = Get-SummaryResult -Sources $sources -PacResults $pacResults -WpadDnsResults $wpadDnsResults

$effectiveConfig = [pscustomobject]@{
    UserEffectiveProxy = if ($hasUserPolicyProxy) { $policyUser.ProxyServer } else { $winInetUser.ProxyServer }
    MachineEffectiveProxy = if ($hasMachinePolicyProxy) { $policyMachine.ProxyServer } else { $winInetMachineDefault.ProxyServer }
    SystemWinHTTPProxy = $winHttp.ProxyServer
    UserEffectiveAutoConfigURL = if ($hasUserPolicyProxy) { $policyUser.AutoConfigURL } else { $winInetUser.AutoConfigURL }
    MachineEffectiveAutoConfigURL = if ($hasMachinePolicyProxy) { $policyMachine.AutoConfigURL } else { $winInetMachineDefault.AutoConfigURL }
    SystemWinHTTPAutoConfigURL = $winHttp.AutoConfigURL
}

$report = [ordered]@{
    TimestampUtc      = (Get-Date).ToUniversalTime().ToString("o")
    ComputerName      = $env:COMPUTERNAME
    ExecutionContext  = $runContext
    Summary           = $summary
    EffectiveConfig   = $effectiveConfig
    Sources           = $sources
    WinHTTP           = $winHttp
    WinINetUser       = $winInetUser
    WinINetMachine    = $winInetMachineDefault
    PolicyUser        = $policyUser
    PolicyMachine     = $policyMachine
    WPAD_DNS          = $wpadDnsResults
    PAC               = $pacResults
    Notes             = $notes
}

if ($RawEvidence) {
    $report["RawEvidence"] = [ordered]@{
        WinHttpShowProxyText = $winHttp.RawShowProxy
        WinHttpShowAdvText   = $winHttp.RawShowAdvProxy
    }
}

# Write JSON output
$report | ConvertTo-Json -Depth 8 | Set-Content -Path $JsonPath -Encoding UTF8

# Log summary and source lines to TXT
Write-Log -Message "Status: $($summary.Status)"
Write-Log -Message "OverallConfidence: $($summary.OverallConfidence)"
Write-Log -Message "PrimaryDetectionPath: $($summary.PrimaryDetectionPath)"
foreach ($s in $sources) {
    Write-Log -Message ("SourceType={0}; Context={1}; Signal={2}; Value={3}; Origin={4}; Effective={5}; Confidence={6}" -f $s.SourceType, $s.Context, $s.Signal, $s.EffectiveValue, $s.SourcePathOrCommand, $s.IsEffective, $s.Confidence)
}

# Console output
Write-Host ""
Write-Host "==============================================="
Write-Host "   Proxy Detection With Source Attribution"
Write-Host "==============================================="
Write-Host "Status              : $($summary.Status)"
Write-Host "Detected            : $($summary.Detected)"
Write-Host "Overall Confidence  : $($summary.OverallConfidence)"
Write-Host "Primary Path        : $($summary.PrimaryDetectionPath)"
Write-Host "Sources Found       : $($summary.SourceCount)"
Write-Host "WPAD DNS Resolved   : $($summary.WpadDnsResolved)"
Write-Host "PAC Validated       : $($summary.PacValidated)"
Write-Host ""
Write-Host "Top Sources (up to 10):"
$sources | Sort-Object Confidence -Descending | Select-Object -First 10 | ForEach-Object {
    Write-Host (" - [{0}] {1} | {2} | {3} | C={4}" -f $_.SourceType, $_.Context, $_.Signal, $_.EffectiveValue, $_.Confidence)
}
Write-Host ""
Write-Host "TXT Log             : $LogPath"
Write-Host "JSON Report         : $JsonPath"
Write-Host ""

Write-Log -Message "Proxy detection script completed"
