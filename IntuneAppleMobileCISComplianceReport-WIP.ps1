<#
Intune iOS/iPadOS CIS Compliance Report (Excel)
- Reads CIS mapping (CIS section + level) from "Translation to DDM" sheet
- Validates Settings Catalog policies against expected values
- Outputs Excel (Detailed, Summary, Missing Settings) with conditional formatting

Author: M365 Copilot helper
#>

param(
    # Path to your CIS mapping workbook (the one you uploaded)
    [string]$MappingPath = "C:\Temp\iOS26_Intune_Settings_for_CIS_Controls.xlsx",
    [string]$MappingSheet = "Translation to DDM",

    # Output workbook path
    [string]$OutputPath  = "C:\Temp\Intune_iOS_CIS_Compliance.xlsx",

    # Optional: filter to a single Settings Catalog policy name
    [string]$PolicyNameFilter = "",

    # Optional: pin tenant
    [string]$TenantId = ""
)

# -------------------------------
# Ensure modules & connect Graph
# -------------------------------
function Ensure-Module {
    param([string]$Name)
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        Install-Module $Name -Scope CurrentUser -Force -ErrorAction Stop
    }
    Import-Module $Name -ErrorAction Stop
}
Ensure-Module -Name Microsoft.Graph
Ensure-Module -Name ImportExcel

$scopes = @("DeviceManagementConfiguration.Read.All")
if ($TenantId) { Connect-MgGraph -TenantId $TenantId -Scopes $scopes } else { Connect-MgGraph -Scopes $scopes }

# ----------------------------------------
# Expected values (from your requirements)
# ----------------------------------------
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

    "Lock Screen Footnote / Asset Tag Information"             = "Configured"
}

# -----------------------------------------
# Name resolution for common display aliases
# -----------------------------------------
$Aliases = @{
    "Locked Screen Experience > Block Wallet notifications in lock screen" = "Block Wallet notifications in lock screen"
    "Allow Shared Stream"                                                  = "Allow Photo Stream (Deprecated) / Allow Shared Stream"
    "Allow Photo Stream (Deprecated)"                                      = "Allow Photo Stream (Deprecated) / Allow Shared Stream"
}

function Resolve-ExpectedKey {
    param([string]$DisplayName)
    if ($Expected.ContainsKey($DisplayName)) { return $DisplayName }
    if ($Aliases.ContainsKey($DisplayName))  { return $Aliases[$DisplayName] }
    # Trim well-known prefix
    $n = ($DisplayName -replace '^Locked Screen Experience\s*>\s*','').Trim()
    if ($Expected.ContainsKey($n)) { return $n }
    return $null
}

# ------------------------------------------------------
# Load CIS mapping (CIS section number + Level L1/L2)
# ------------------------------------------------------
if (-not (Test-Path $MappingPath)) {
    throw "CIS mapping workbook not found: $MappingPath"
}

$CisRows = Import-Excel -Path $MappingPath -WorksheetName $MappingSheet -ErrorAction Stop
# Expected columns in "Translation to DDM":
# "Intune Setting Name/Location", "CIS Control", "Level"
$CisMap = @{}
foreach ($r in $CisRows) {
    $name = ($r.'Intune Setting Name/Location').ToString().Trim()
    if (-not [string]::IsNullOrWhiteSpace($name)) {
        $CisMap[$name] = [pscustomobject]@{
            CisSection = $r.'CIS Control'
            Level      = $r.'Level'
        }
    }
}
# Also load aliases into CIS map for lookups
foreach ($kvp in $Aliases.GetEnumerator()) {
    if ($CisMap.ContainsKey($kvp.Key) -and -not $CisMap.ContainsKey($kvp.Value)) {
        $CisMap[$kvp.Value] = $CisMap[$kvp.Key]
    }
}

# -------------------------------------------------------
# Pull iOS/iPadOS Settings Catalog configuration policies
# -------------------------------------------------------
$Policies = Get-MgDeviceManagementConfigurationPolicy -All |
    Where-Object { $_.Platforms -and ($_.Platforms -contains "iOS" -or $_.Platforms -like "*iOS*") }

if ($PolicyNameFilter) {
    $Policies = $Policies | Where-Object { $_.Name -eq $PolicyNameFilter }
}

Write-Host "Found $($Policies.Count) iOS/iPadOS Settings Catalog policies."

# Definitions (fallback when displayName isn't on the setting)
$Defs = Get-MgDeviceManagementSettingDefinition -All
$DefById = @{}
foreach ($d in $Defs) { $DefById[$d.Id] = $d }

# ------------------------
# Normalize actual values
# ------------------------
function Get-NormalizedValue {
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

# -------------------------------
# Compare actual vs expected text
# -------------------------------
function Compare-Values {
    param(
        [Parameter(Mandatory)]$Actual,
        [Parameter(Mandatory)][string]$ExpectedValue
    )
    $a = $Actual.ToString().Trim().ToLower()
    $e = $ExpectedValue.ToString().Trim().ToLower()
    return ($a -eq $e)
}

# Collect results
$detailed = New-Object System.Collections.Generic.List[object]
$foundKeys = New-Object System.Collections.Generic.HashSet[string]

foreach ($p in $Policies) {
    $settings = Get-MgDeviceManagementConfigurationPolicySetting `
        -DeviceManagementConfigurationPolicyId $p.Id -All

    foreach ($s in $settings) {
        # Prefer the setting's embedded displayName; fall back to definition
        $disp = $s.AdditionalProperties.displayName
        if (-not $disp) {
            $def = $DefById[$s.SettingDefinitionId]
            $disp = $def.DisplayName
        }

        $actual = Get-NormalizedValue -ValueJson $s.ValueJson

        $ekey = Resolve-ExpectedKey -DisplayName $disp
        if ($null -ne $ekey) {

            $expected = $Expected[$ekey]
            $pass = Compare-Values -Actual $actual -ExpectedValue $expected

            # CIS section/level lookup: try exact, alias, and prefix-trim
            $cis = $CisMap[$disp]
            if (-not $cis) {
                $cis = $CisMap[$ekey]
            }
            if (-not $cis) {
                $trimName = ($disp -replace '^Locked Screen Experience\s*>\s*','').Trim()
                $cis = $CisMap[$trimName]
            }

            $detailed.Add([pscustomobject]@{
                PolicyName    = $p.Name
                PolicyId      = $p.Id
                Setting       = $disp
                ActualValue   = $actual
                ExpectedValue = $expected
                Compliance    = if ($pass) { 'PASS' } else { 'FAIL' }
                CisSection    = if ($cis) { $cis.CisSection } else { '' }
                Level         = if ($cis) { $cis.Level } else { '' }
            })

            $null = $foundKeys.Add($ekey)
        }
        else {
            # Setting present, but not in Expected list -> we ignore for compliance,
            # but you could capture here if you want a "non-evaluated" sheet.
            continue
        }
    }
}

# -------------------------------------------------
# Build "Missing Settings" sheet (not found at all)
# -------------------------------------------------
$missing = foreach ($k in $Expected.Keys) {
    if (-not $foundKeys.Contains($k)) {
        $cis = $CisMap[$k]
        [pscustomobject]@{
            Setting       = $k
            ExpectedValue = $Expected[$k]
            CisSection    = if ($cis) { $cis.CisSection } else { '' }
            Level         = if ($cis) { $cis.Level } else { '' }
            Note          = "No matching setting found in any iOS Settings Catalog policy"
        }
    }
}

# ------------------------
# Summary sheet per policy
# ------------------------
$summary =
    $detailed |
    Group-Object PolicyName |
    ForEach-Object {
        $name  = $_.Name
        $rows  = $_.Group
        $pass  = ($rows | Where-Object Compliance -eq 'PASS').Count
        $fail  = ($rows | Where-Object Compliance -eq 'FAIL').Count
        $unk   = ($rows | Where-Object { $_.Compliance -notin @('PASS','FAIL') }).Count
        $total = $rows.Count
        [pscustomobject]@{
            PolicyName     = $name
            TotalSettings  = $total
            Pass           = $pass
            Fail           = $fail
            Unknown        = $unk
            CompliancePct  = if ($total -gt 0) { [math]::Round(($pass / $total) * 100, 1) } else { 0 }
        }
    } |
    Sort-Object -Property CompliancePct -Descending

# -------------------------
# Export to Excel + styling
# -------------------------
# Detailed first, with -PassThru so we can add conditional formatting
$xl = $detailed | Export-Excel -Path $OutputPath -WorksheetName "Detailed" `
    -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter -PassThru

# Add conditional formatting to Detailed (Compliance column)
# Column order: PolicyName, PolicyId, Setting, ActualValue, ExpectedValue, Compliance, CisSection, Level
Add-ConditionalFormatting -WorkSheet $xl.Workbook.Worksheets["Detailed"] `
    -Address "F2:F5000" -RuleType ContainsText -Text "PASS" `
    -BackgroundColor "#C6EFCE" -StopIfTrue

Add-ConditionalFormatting -WorkSheet $xl.Workbook.Worksheets["Detailed"] `
    -Address "F2:F5000" -RuleType ContainsText -Text "FAIL" `
    -BackgroundColor "#FFC7CE" -StopIfTrue

Add-ConditionalFormatting -WorkSheet $xl.Workbook.Worksheets["Detailed"] `
    -Address "F2:F5000" -RuleType ContainsText -Text "Unknown" `
    -BackgroundColor "#FFEB9C"

# Summary sheet
$summary | Export-Excel -ExcelPackage $xl -WorksheetName "Summary" `
    -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter

# Three-color scale on CompliancePct (green/yellow/red)
Add-ConditionalFormatting -WorkSheet $xl.Workbook.Worksheets["Summary"] `
    -Address "F2:F5000" -RuleType ThreeColorScale `
    -Color "#63BE7B","#FFEB84","#F8696B"

# Missing Settings sheet
$missing | Export-Excel -ExcelPackage $xl -WorksheetName "Missing Settings" `
    -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter

# Save
Close-ExcelPackage $xl

Write-Host "`nExcel report created: $OutputPath" -ForegroundColor Green