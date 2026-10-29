# -------------------------------------------------------
# Generate Excel Compliance Report for iOS CIS Settings
# -------------------------------------------------------

Import-Module Microsoft.Graph.deviceManagement -ErrorAction Stop

# =========================
# Module install/import (minimal footprint)
# =========================
$requiredModules = @(
    'TUN.CredentialManager',                                  # for Get-StoredCredential | TUN version needed for PowerShell
    'Microsoft.Graph.Beta.DeviceManagement',                  # for Get-MgBetaDeviceManagementConfigurationPolicy and Get-MgBetaDeviceManagementConfigurationPolicySetting
    'Microsoft.Graph.Beta.Devices',                           # for Get-MgBetaDeviceAppManagementVppToken
    'ImportExcel'                                             # for Export-Excel
)

foreach ($m in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $m)) {
        Install-Module $m -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module $m -ErrorAction Stop
}

# =========================
# TEST TENANT APP (Service Principal via certificate thumbprint saved on server)
# =========================
# App Name: EIS - Intune Monthly Reports and Notifications | App-Auth with certificate (non-interactive, no token caching)
$TenantID     = "4ec3f009-66b1-45dd-a319-82e49e19cd5b"     # Test Tenant
$ClientID     = "dac6f085-fa17-4b91-a1f0-ccc3d2a3ca65"     # Service Principal (app) client ID

# =========================
# Helper: Get-PlainTextSecret | Retrieve stored credential as plain text string (for app auth scenarios where secret is stored in Windows Credential Manager)
# =========================
function Get-PlainTextSecret {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Target)
    # Disambiguate the module to avoid calling the wrong version
    $cred = $null
    if (Get-Module -ListAvailable -Name 'TUN.CredentialManager') {
        $cred = TUN.CredentialManager\Get-StoredCredential -Target $Target
    } elseif (Get-Module -ListAvailable -Name 'CredentialManager') {
        $cred = CredentialManager\Get-StoredCredential -Target $Target
    } else {
        throw "No CredentialManager module found. Install TUN.CredentialManager or import the legacy module."
    }

    if (-not $cred) { throw "Credential not found for target '$Target'." }

    $pwd = $cred.Password

    switch ($pwd.GetType().FullName) {
        'System.Security.SecureString' {
            # Decrypt securestring to plaintext in memory (does not write to console or disk)
            return [Runtime.InteropServices.Marshal]::PtrToStringAuto(
                [Runtime.InteropServices.Marshal]::SecureStringToBSTR($pwd)
            )
        }
        'System.String' {
            return $pwd
        }
        default {
            throw "Unexpected password type: $($pwd.GetType().FullName)"
        }
    }
}

# ---------------------------
# Helper: Retrieve Certificate Info from Credential Manager
# ---------------------------

# Get plaintext from Credential Manager (handles SecureString or String correctly)
$thumbRaw = Get-PlainTextSecret -Target 'TEST-IntuneMonthlyReportsCertThumbprint'

# Remove anything that's not hex, and enforce uppercase
$CertThumbprint = ($thumbRaw -replace '[^A-Fa-f0-9]', '').ToUpper()

if ($CertThumbprint.Length -ne 40) {
    throw "Thumbprint must be exactly 40 hex characters. Got $($CertThumbprint.Length)."
}

# Prefer passing the certificate object (works whether it's in CurrentUser or LocalMachine)
$cert = Get-ChildItem Cert:\CurrentUser\My, Cert:\LocalMachine\My -Recurse |
        Where-Object Thumbprint -eq $CertThumbprint |
        Select-Object -First 1

if (-not $cert) { throw "Certificate $CertThumbprint not found in CurrentUser or LocalMachine 'My' stores." }

# =========================
# Connect to Graph (App-only) using the certificate object
# =========================
Connect-MgGraph -TenantId $TenantID -ClientId $ClientID -Certificate $cert -NoWelcome



# ==========================
# EXPECTED CIS VALUES TABLE
# ==========================
$Expected = @{
    "Allow Open From Managed To Unmanaged"                     = "False"
    "Force AirDrop Unmanaged"                                  = "True"
    "Require Managed Pasteboard"                               = "True"
    "Allow UI App Installation"                                = "False"
    "Allow Files Network Drive Access"                         = "False"

    "Allow Assistant While Locked (Deprecated)"                = "False"
    "Safari Force Fraud Warning"                               = "True"

    "Force Encrypted Backup"                                   = "True"
    "Allow Managed Apps Cloud Sync"                            = "False"
    "Allow Enterprise Book Backup"                             = "False"
    "Allow Enterprise Book Metadata Sync"                      = "False"
    "Block iCloud Photos sync"                                 = "Yes"
    "Allow Cloud Photo Library"                                = "False"
    "Allow Photo Stream (Deprecated) / Allow Shared Stream"    = "False"
    "Allow Activity Continuation"                              = "False"
    "Allow Cloud Backup"                                       = "False"
    "Allow Cloud Document Sync"                                = "False"
    "Allow Cloud Keychain Sync"                                = "False"

    "Force Watch Wrist Detection"                              = "True"
    "Force AirPlay Outgoing Requests Pairing Password"         = "True"
    "Allow Auto Unlock"                                        = "False"
    "Allow AirPrint iBeacon Discovery"                         = "False"
    "Allow Files USB Drive Access"                             = "False"

    "Allow Diagnostic Submission"                              = "False"
    "Allow Screen Shot"                                        = "False"
    "Allow Untrusted TLS Prompt"                               = "False"
    "Force Limit Ad Tracking"                                  = "True"
    "Allow Enterprise App Trust"                               = "False"
    "Allow Apple Personalized Advertising"                     = "False"
    "Allow Erase Content And Settings"                         = "False"
    "Allow Device Name Modification"                           = "False"
    "Allow UI Configuration Profile Installation"              = "False"
    "Activation Lock Allowed While Supervised"                 = "True"
    "Force Automatic Date And Time"                            = "True"
    "Allow VPN Creation"                                       = "False"

    "Allow Lock Screen Control Center"                         = "False"
    "Allow Lock Screen Notifications View"                     = "False"
    "Allow Lock Screen Today View"                             = "False"
    "Block Wallet notifications in lock screen"                = "Yes"

    "Require Passcode on Device"                               = "True"
    "Allow Simple Passcode"                                    = "False"
    "Require Alphanumeric Passcode"                            = "Alphanumeric"
    "Min Length"                                               = "6"
    "Max Grace Period"                                         = "0"
    "Max Inactivity"                                           = "2"
    "Allow Fingerprint For Unlock"                             = "False"
    "Allow Password Proximity Requests"                        = "False"
    "Allow Password Sharing"                                   = "False"
    "Force Authentication Before Auto Fill"                    = "True"

    "Allow Voice Dialing (Deprecated)"                         = "False"

    "Lock Screen Footnote"                                     = "Configured"
    "Asset Tag Information"                                    = "Configured"
}


# ---------------------------------------------
# Helpers for beta settings resolution
# ---------------------------------------------
function Get-SettingDisplayName {
    param($Setting)
    if ($Setting.AdditionalProperties.displayName) {
        return $Setting.AdditionalProperties.displayName
    } elseif ($Setting.SettingDefinitionIdDisplayName) {
        return $Setting.SettingDefinitionIdDisplayName
    } elseif ($Setting.SettingInstance -and $Setting.SettingInstance.DisplayName) {
        return $Setting.SettingInstance.DisplayName
    } elseif ($Setting.SettingInstance -and $Setting.SettingInstance.SettingDefinitionIdDisplayName) {
        return $Setting.SettingInstance.SettingDefinitionIdDisplayName
    }
    return $null
}

function Get-NormalizedValueFromJson {
    param($ValueJson)
    $actual = $ValueJson
    if ($actual -is [string]) {
        try { $actual = $actual | ConvertFrom-Json } catch { return $ValueJson }
    }
    if ($actual -is [string]) { return $actual }
    foreach ($key in "value","simpleValue","StringValue","BooleanValue","IntValue","EnumValue") {
        if ($actual.$key) { return $actual.$key }
    }
    return $actual
}

function Get-SettingActualValue {
    param($Setting)
    # Try typed value first (beta setting instance)
    if ($Setting.SettingInstance -and $Setting.SettingInstance.SimpleSettingValue -and $Setting.SettingInstance.SimpleSettingValue.Value) {
        return $Setting.SettingInstance.SimpleSettingValue.Value
    }
    # Fall back to ValueJson normalization
    return Get-NormalizedValueFromJson -ValueJson $Setting.ValueJson
}

function Compare-Values {
    param($Actual, [string]$ExpectedValue)
    $a = $Actual.ToString().Trim().ToLower()
    $e = $ExpectedValue.ToString().Trim().ToLower()
    return ($e.Length -gt 0 -and $a -eq $e)
}

# Optional: alias normalizer (helps when Expected keys differ slightly)
$Aliases = @{
    "Locked Screen Experience > Block Wallet notifications in lock screen" = "Block Wallet notifications in lock screen"
    "Allow Shared Stream"                                                  = "Allow Photo Stream (Deprecated) / Allow Shared Stream"
    "Allow Photo Stream (Deprecated)"                                      = "Allow Photo Stream (Deprecated) / Allow Shared Stream"
}
function Resolve-ExpectedKey {
    param([string]$DisplayName)
    if ([string]::IsNullOrWhiteSpace($DisplayName)) { return $null }
    if ($Expected.ContainsKey($DisplayName)) { return $DisplayName }
    if ($Aliases.ContainsKey($DisplayName))  { return $Aliases[$DisplayName] }
    $trimName = ($DisplayName -replace '^Locked Screen Experience\s*>\s*','').Trim()
    if ($Expected.ContainsKey($trimName)) { return $trimName }
    return $null
}

# ---------------------------------------------
# Settings loop (ALWAYS adds rows to $Results)
# ---------------------------------------------
$RunTimestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

foreach ($Policy in $Policies) {
    $Settings = Get-MgBetaDeviceManagementConfigurationPolicySetting `
        -DeviceManagementConfigurationPolicyId $Policy.Id -All

    foreach ($Setting in $Settings) {

        # Resolve display name & actual value
        $Name   = Get-SettingDisplayName -Setting $Setting
        if ([string]::IsNullOrWhiteSpace($Name)) {
            # Can't evaluate without a name; still record a row with Unknown compliance
            $Actual = Get-SettingActualValue -Setting $Setting
            $Results += [pscustomobject]@{
                PolicyName     = $Policy.Name
                PolicyId       = $Policy.Id
                Setting        = "[Unnamed Setting: $($Setting.SettingDefinitionId)]"
                ActualValue    = $Actual
                ExpectedValue  = ""
                Compliance     = "Unknown"
                TenantId       = $TenantID
                RunTimestamp   = $RunTimestamp
            }
            continue
        }

        $Actual = Get-SettingActualValue -Setting $Setting

        # Try to find a matching expected value; otherwise leave blank & mark Unknown
        $ekey          = Resolve-ExpectedKey -DisplayName $Name
        $hasExpected   = ($ekey -ne $null) -and $Expected.ContainsKey($ekey)
        $ExpectedValue = if ($hasExpected) { $Expected[$ekey] } else { "" }
        $Compliance    = "Unknown"

        if ($hasExpected) {
            $Compliance = if (Compare-Values -Actual $Actual -ExpectedValue $ExpectedValue) { "PASS" } else { "FAIL" }
        }

        # ALWAYS add a row to $Results (Option C)
        $Results += [pscustomobject]@{
            PolicyName     = $Policy.Name
            PolicyId       = $Policy.Id
            Setting        = $Name
            ActualValue    = $Actual
            ExpectedValue  = $ExpectedValue
            Compliance     = $Compliance
            TenantId       = $TenantID
            RunTimestamp   = $RunTimestamp
        }
    }
}

# Quick sanity check (optional)
Write-Host ("Collected {0} settings across {1} policies." -f $Results.Count, $Policies.Count) -ForegroundColor Cyan
$Results | Select-Object -First 5 | Format-Table -AutoSize

# ===================================
# BUILD SUMMARY SHEET
# ===================================
$Summary = $Results |
    Group-Object PolicyName |
    ForEach-Object {
        $Total = $_.Group.Count
        $Pass  = ($_.Group | Where-Object Compliance -eq "PASS").Count
        $Fail  = ($_.Group | Where-Object Compliance -eq "FAIL").Count

        [pscustomobject]@{
            PolicyName    = $_.Name
            TotalSettings = $Total
            Pass          = $Pass
            Fail          = $Fail
            CompliancePct = [math]::Round(($Pass / $Total) * 100, 1)
        }
    }

# =========================
# Export
# =========================
$desktop = Join-Path $env:USERPROFILE 'Desktop\IntuneReports'
if (-not (Test-Path $desktop)) { $desktop = (Get-Location).Path }  # fallback for non-Windows environments

# Date-stamped file name (yyyyMMMdd)
$stamp = (Get-Date).ToString('yyyyMMMdd')
$mergedPath = Join-Path $desktop "TEST_Intune_AppleConfigs_$stamp.xlsx"
if ($mergedRows.Count -gt 0) {
    $mergedRows | Sort-Object Type, DaysUntilExpiry, TokenName | Export-Csv -NoTypeInformation -Path $mergedPath
    Write-Host "Merged CSV saved: $mergedPath" -ForegroundColor Green
} else {
    Write-Warning "No token data available to export (merged CSV not created)."
}