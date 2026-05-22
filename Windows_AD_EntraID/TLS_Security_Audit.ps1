<#
.SYNOPSIS
    Security audit script for TLS configuration and cipher suite settings
.DESCRIPTION
    This script checks TLS protocol versions and cipher suite configurations 
    on a Windows system to identify potential security vulnerabilities.
.NOTES
    Author: Security Audit Script
    Version: 2.0
    
.EXAMPLE
    .\TLS_Security_Audit.ps1
    Run the security audit script to check TLS configurations
    
.INPUTS
    None - Script reads directly from system registry and TLS configuration
    
.OUTPUTS
    PSCustomObject array with structured security audit results
    - Protocol/Component name
    - Status (Healthy, Warning, Critical, Missing)
    - Severity level
    - Detailed message
    - Security recommendations

.COMPONENT
    Windows Registry, SCHANNEL, TLS Configuration

.FUNCTIONALITY
    Performs security audit of TLS protocol versions and identifies 
    potentially insecure cipher suites that should be disabled.
#>

#Requires -Version 5.1

[CmdletBinding()]
param()

#region Helper Functions
function Test-IsAdministrator {
    <#
    .SYNOPSIS
        Checks if the script is running with administrator privileges
    #>
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Test-TLSProtocol {
    <#
    .SYNOPSIS
        Tests a specific TLS/SSL protocol configuration
    .PARAMETER Protocol
        Protocol name (e.g., "TLS 1.0", "SSL 2.0")
    .PARAMETER Type
        Type: "Server" or "Client"
    .PARAMETER ShouldBeEnabled
        Whether this protocol should be enabled (security best practice)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Protocol,
        
        [Parameter(Mandatory)]
        [ValidateSet("Server", "Client")]
        [string]$Type,
        
        [Parameter(Mandatory)]
        [bool]$ShouldBeEnabled
    )
    
    $regPath = "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\$Protocol\$Type"
    $componentName = "$Protocol $Type"
    
    try {
        $regValue = Get-ItemProperty -Path $regPath -Name Enabled -ErrorAction Stop
        $isEnabled = $regValue.Enabled -eq 1
        
        if ($ShouldBeEnabled) {
            # Protocol should be enabled (TLS 1.2, TLS 1.3)
            if (-not $isEnabled) {
                return [PSCustomObject]@{
                    Component = $componentName
                    Status = "Warning"
                    Severity = "Medium"
                    Message = "$componentName is disabled"
                    Recommendation = "Enable $componentName for secure communications"
                }
            } else {
                return [PSCustomObject]@{
                    Component = $componentName
                    Status = "Healthy"
                    Severity = "None"
                    Message = "$componentName is properly enabled"
                    Recommendation = $null
                }
            }
        } else {
            # Protocol should be disabled (TLS 1.0, TLS 1.1, SSL 2.0, SSL 3.0)
            if ($isEnabled) {
                $severity = if ($Protocol -like "SSL*") { "Critical" } else { "High" }
                return [PSCustomObject]@{
                    Component = $componentName
                    Status = "Critical"
                    Severity = $severity
                    Message = "$componentName is enabled (INSECURE)"
                    Recommendation = "Disable $componentName immediately - it is deprecated and insecure"
                }
            } else {
                return [PSCustomObject]@{
                    Component = $componentName
                    Status = "Healthy"
                    Severity = "None"
                    Message = "$componentName is properly disabled"
                    Recommendation = $null
                }
            }
        }
    } catch [System.Management.Automation.ItemNotFoundException] {
        # Registry key doesn't exist
        if ($ShouldBeEnabled) {
            return [PSCustomObject]@{
                Component = $componentName
                Status = "Warning"
                Severity = "Medium"
                Message = "$componentName registry key missing"
                Recommendation = "Explicitly configure $componentName (enable for secure protocols)"
            }
        } else {
            return [PSCustomObject]@{
                Component = $componentName
                Status = "Warning"
                Severity = "Low"
                Message = "$componentName registry key missing"
                Recommendation = "Explicitly disable $componentName for better security posture"
            }
        }
    } catch {
        return [PSCustomObject]@{
            Component = $componentName
            Status = "Error"
            Severity = "Unknown"
            Message = "Error reading registry: $($_.Exception.Message)"
            Recommendation = "Verify registry permissions and system configuration"
        }
    }
}
#endregion

#region Pre-flight Checks
# Check for administrator privileges
if (-not (Test-IsAdministrator)) {
    Write-Warning "This script should be run with administrator privileges for complete results."
    Write-Warning "Some registry keys may not be accessible without elevation."
}

# Check PowerShell version
if ($PSVersionTable.PSVersion.Major -lt 5) {
    Write-Warning "PowerShell 5.1 or higher is recommended for best results."
}
#endregion

# region TLS Protocol Version Checks
# Initialize results array
$results = @() 

# Test deprecated protocols (should be disabled)
Write-Verbose "Checking deprecated SSL/TLS protocols..."
$results += Test-TLSProtocol -Protocol "SSL 2.0" -Type "Server" -ShouldBeEnabled $false
$results += Test-TLSProtocol -Protocol "SSL 2.0" -Type "Client" -ShouldBeEnabled $false
$results += Test-TLSProtocol -Protocol "SSL 3.0" -Type "Server" -ShouldBeEnabled $false
$results += Test-TLSProtocol -Protocol "SSL 3.0" -Type "Client" -ShouldBeEnabled $false
$results += Test-TLSProtocol -Protocol "TLS 1.0" -Type "Server" -ShouldBeEnabled $false
$results += Test-TLSProtocol -Protocol "TLS 1.0" -Type "Client" -ShouldBeEnabled $false
$results += Test-TLSProtocol -Protocol "TLS 1.1" -Type "Server" -ShouldBeEnabled $false
$results += Test-TLSProtocol -Protocol "TLS 1.1" -Type "Client" -ShouldBeEnabled $false

# Test modern protocols (should be enabled)
Write-Verbose "Checking modern TLS protocols..."
$results += Test-TLSProtocol -Protocol "TLS 1.2" -Type "Server" -ShouldBeEnabled $true
$results += Test-TLSProtocol -Protocol "TLS 1.2" -Type "Client" -ShouldBeEnabled $true
$results += Test-TLSProtocol -Protocol "TLS 1.3" -Type "Server" -ShouldBeEnabled $true
$results += Test-TLSProtocol -Protocol "TLS 1.3" -Type "Client" -ShouldBeEnabled $true
# endregion

# region Cipher Suite Security Check
Write-Verbose "Checking cipher suite configuration..."

# Define list of deprecated/insecure cipher suites that should be disabled
# Based on current security best practices (2026):
# - All RSA key exchange (no forward secrecy)
# - All CBC mode ciphers (vulnerable to padding oracle attacks)
# - All 3DES, RC4, DES, NULL, MD5 ciphers
# - DHE_DSS ciphers (DSS is weaker than RSA)
# Modern secure ciphers should use: ECDHE + AES-GCM or ChaCha20-Poly1305

$oldCipherSuites = @(
    # RSA key exchange ciphers (no forward secrecy)
    "TLS_RSA_WITH_AES_256_GCM_SHA384"
    "TLS_RSA_WITH_AES_128_GCM_SHA256"
    "TLS_RSA_WITH_AES_256_CBC_SHA256"
    "TLS_RSA_WITH_AES_128_CBC_SHA256"
    "TLS_RSA_WITH_AES_256_CBC_SHA"
    "TLS_RSA_WITH_AES_128_CBC_SHA"
    "TLS_RSA_WITH_3DES_EDE_CBC_SHA"
    
    # DHE_RSA with CBC mode (vulnerable to attacks)
    "TLS_DHE_RSA_WITH_AES_256_CBC_SHA256"
    "TLS_DHE_RSA_WITH_AES_128_CBC_SHA256"
    "TLS_DHE_RSA_WITH_AES_256_CBC_SHA"
    "TLS_DHE_RSA_WITH_AES_128_CBC_SHA"
    
    # DHE_DSS ciphers (DSS weaker than RSA)
    "TLS_DHE_DSS_WITH_AES_256_CBC_SHA256"
    "TLS_DHE_DSS_WITH_AES_128_CBC_SHA256"
    "TLS_DHE_DSS_WITH_AES_256_CBC_SHA"
    "TLS_DHE_DSS_WITH_AES_128_CBC_SHA"
    "TLS_DHE_DSS_WITH_3DES_EDE_CBC_SHA"
    
    # ECDHE with CBC mode (should prefer GCM)
    "TLS_ECDHE_RSA_WITH_AES_256_CBC_SHA384"
    "TLS_ECDHE_RSA_WITH_AES_128_CBC_SHA256"
    "TLS_ECDHE_RSA_WITH_AES_256_CBC_SHA"
    "TLS_ECDHE_RSA_WITH_AES_128_CBC_SHA"
    "TLS_ECDHE_ECDSA_WITH_AES_256_CBC_SHA384"
    "TLS_ECDHE_ECDSA_WITH_AES_128_CBC_SHA256"
    "TLS_ECDHE_ECDSA_WITH_AES_256_CBC_SHA"
    "TLS_ECDHE_ECDSA_WITH_AES_128_CBC_SHA"
    
    # RC4 ciphers (completely broken)
    "TLS_RSA_WITH_RC4_128_SHA"
    "TLS_RSA_WITH_RC4_128_MD5"
    "TLS_ECDHE_RSA_WITH_RC4_128_SHA"
    "TLS_ECDHE_ECDSA_WITH_RC4_128_SHA"
    
    # NULL encryption ciphers (no encryption!)
    "TLS_RSA_WITH_NULL_SHA256"
    "TLS_RSA_WITH_NULL_SHA"
    "TLS_ECDHE_RSA_WITH_NULL_SHA"
    "TLS_ECDHE_ECDSA_WITH_NULL_SHA"
    
    # PSK ciphers (pre-shared key, uncommon in enterprise)
    "TLS_PSK_WITH_AES_256_GCM_SHA384"
    "TLS_PSK_WITH_AES_128_GCM_SHA256"
    "TLS_PSK_WITH_AES_256_CBC_SHA384"
    "TLS_PSK_WITH_AES_128_CBC_SHA256"
    "TLS_PSK_WITH_NULL_SHA384"
    "TLS_PSK_WITH_NULL_SHA256"
    
    # Additional weak ciphers
    "TLS_RSA_EXPORT_WITH_RC4_40_MD5"
    "TLS_RSA_EXPORT_WITH_DES40_CBC_SHA"
    "TLS_RSA_WITH_DES_CBC_SHA"
    "TLS_DHE_DSS_WITH_DES_CBC_SHA"
)

# Recommended secure cipher suites for comparison (TLS 1.2 and 1.3):
# TLS 1.3 (Best - if supported by Windows Server 2022+ / Windows 11+):
#   - TLS_AES_256_GCM_SHA384
#   - TLS_AES_128_GCM_SHA256
#   - TLS_CHACHA20_POLY1305_SHA256
# 
# TLS 1.2 (Secure - for Windows Server 2012 R2+ / Windows 8.1+):
#   - TLS_ECDHE_RSA_WITH_AES_256_GCM_SHA384
#   - TLS_ECDHE_RSA_WITH_AES_128_GCM_SHA256
#   - TLS_ECDHE_ECDSA_WITH_AES_256_GCM_SHA384
#   - TLS_ECDHE_ECDSA_WITH_AES_128_GCM_SHA256
#
# Key requirements for secure ciphers:
#   - Forward secrecy (ECDHE or DHE key exchange)
#   - AEAD mode (GCM, ChaCha20-Poly1305, not CBC)
#   - Strong encryption (AES-128 or higher, not 3DES/RC4)
#   - Strong hash (SHA-256 or higher, not SHA-1 or MD5)


# Check if Get-TlsCipherSuite cmdlet is available
$cipherSuiteCmdletAvailable = Get-Command Get-TlsCipherSuite -ErrorAction SilentlyContinue

if ($cipherSuiteCmdletAvailable) {
    # Initialize array to store enabled old cipher suites
    $suitesFound = @()
    
    # Check each cipher suite to see if it's currently enabled
    foreach ($suite in $oldCipherSuites) {
        try {
            $foundSuite = Get-TlsCipherSuite -Name $suite -ErrorAction SilentlyContinue
            if ($foundSuite) {
                $suitesFound += $foundSuite
            }
        } catch {
            $errorMsg = $_.Exception.Message
            Write-Verbose "Error checking cipher suite ${suite}: $errorMsg"
        }
    }
    
    # Evaluate cipher suite results
    if ($suitesFound.Count -eq 0) {
        $results += [PSCustomObject]@{
            Component = "Cipher Suites"
            Status = "Healthy"
            Severity = "None"
            Message = "No deprecated cipher suites found enabled"
            Recommendation = $null
        }
    } else {
        $suiteNames = ($suitesFound | Select-Object -ExpandProperty Name) -join ", "
        $results += [PSCustomObject]@{
            Component = "Cipher Suites"
            Status = "Critical"
            Severity = "High"
            Message = "Found $($suitesFound.Count) deprecated cipher suite(s) enabled: $suiteNames"
            Recommendation = "Disable deprecated cipher suites and use only modern secure cipher suites"
        }
    }
} else {
    $results += [PSCustomObject]@{
        Component = "Cipher Suites"
        Status = "Warning"
        Severity = "Low"
        Message = "Get-TlsCipherSuite cmdlet not available (requires Windows Server 2012 R2+ or Windows 8.1+)"
        Recommendation = "Upgrade to a newer Windows version or check cipher suites manually via registry"
    }
}
# endregion

# region Generate Summary Statistics
$criticalCount = ($results | Where-Object { $_.Status -eq "Critical" }).Count
$warningCount = ($results | Where-Object { $_.Status -eq "Warning" }).Count
$healthyCount = ($results | Where-Object { $_.Status -eq "Healthy" }).Count
$errorCount = ($results | Where-Object { $_.Status -eq "Error" }).Count

$summary = [PSCustomObject]@{
    TotalChecks = $results.Count
    Healthy = $healthyCount
    Warnings = $warningCount
    Critical = $criticalCount
    Errors = $errorCount
    OverallStatus = if ($criticalCount -gt 0) { "Critical" } 
                    elseif ($errorCount -gt 0) { "Error" }
                    elseif ($warningCount -gt 0) { "Warning" } 
                    else { "Healthy" }
}
# endregion

# region Output Results
# Display header
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "   TLS SECURITY AUDIT RESULTS" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

# Display summary
Write-Host "SUMMARY:" -ForegroundColor White
Write-Host "  Total Checks: $($summary.TotalChecks)" -ForegroundColor White
Write-Host "  Healthy: $($summary.Healthy)" -ForegroundColor Green
Write-Host "  Warnings: $($summary.Warnings)" -ForegroundColor Yellow
Write-Host "  Critical: $($summary.Critical)" -ForegroundColor Red
Write-Host "  Errors: $($summary.Errors)" -ForegroundColor Magenta
Write-Host "  Overall Status: $($summary.OverallStatus)`n" -ForegroundColor $(
    switch ($summary.OverallStatus) {
        "Healthy" { "Green" }
        "Warning" { "Yellow" }
        "Critical" { "Red" }
        "Error" { "Magenta" }
    }
)

# Display detailed results
Write-Host "DETAILED RESULTS:" -ForegroundColor White
Write-Host "----------------------------------------`n" -ForegroundColor Gray

foreach ($result in $results) {
    # Color based on status
    $color = switch ($result.Status) {
        "Healthy" { "Green" }
        "Warning" { "Yellow" }
        "Critical" { "Red" }
        "Error" { "Magenta" }
        default { "White" }
    }
    
    Write-Host "[$($result.Status)]" -ForegroundColor $color -NoNewline
    Write-Host " $($result.Component)" -ForegroundColor White
    Write-Host "  $($result.Message)" -ForegroundColor Gray
    
    if ($result.Recommendation) {
        Write-Host "  → Recommendation: $($result.Recommendation)" -ForegroundColor Cyan
    }
    Write-Host ""
}

# Display actionable recommendations
$actionableRecommendations = $results | Where-Object { $null -ne $_.Recommendation }
if ($actionableRecommendations.Count -gt 0) {
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "   ACTION ITEMS ($($actionableRecommendations.Count))" -ForegroundColor Cyan
    Write-Host "========================================`n" -ForegroundColor Cyan
    
    $priority = @{
        "Critical" = 1
        "High" = 2
        "Medium" = 3
        "Low" = 4
    }
    
    $sortedRecommendations = $actionableRecommendations | Sort-Object { $priority[$_.Severity] }
    
    $counter = 1
    foreach ($rec in $sortedRecommendations) {
        Write-Host "$counter. " -ForegroundColor White -NoNewline
        Write-Host "[$($rec.Severity)]" -ForegroundColor Yellow -NoNewline
        Write-Host " $($rec.Component): " -ForegroundColor White -NoNewline
        Write-Host "$($rec.Recommendation)" -ForegroundColor Cyan
        $counter++
    }
    Write-Host ""
}

Write-Host "========================================`n" -ForegroundColor Cyan

# Return the results object for further processing/automation
return $results
# endregion