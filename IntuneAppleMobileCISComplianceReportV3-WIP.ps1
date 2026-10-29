<#
Intune iOS/iPadOS CIS 26 Compliance Audit
Author: Matthew Vagner
Scope: Enumerate Settings Catalog policies targeting iOS/iPadOS;
       compare configured values vs. CIS recommendations (from Excel);
       optionally include legacy templates & iOS compliance policies.

Output: Intune_iOS26_CIS_Compliance_Report.xlsx (Summary + Detailed)
#>

param(
    [string]$MappingPath = "C:\Temp\iOS26_Intune_Settings_for_CIS_Controls.xlsx",
    [string]$OutputPath  = "C:\Temp\Intune_iOS26_CIS_Compliance_Report.xlsx",
    [string[]]$Levels    = @('L1','L2'),     # Filter by CIS Level
    [switch]$IncludeLegacy,                  # Also scan legacy template profiles
    [switch]$IncludeCompliancePolicies,      # Also scan iOS compliance policies (CIS §4.x)
    [string]$TenantId = ""                   # Optional: pin to a tenant ID
)

#region Helper: Module bootstrap
function Ensure-Module {
    param([string]$Name)
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        Write-Host "Installing module: $Name" -ForegroundColor Cyan
        Install-Module $Name -Scope CurrentUser -Force -ErrorAction Stop
    }
    Import-Module $Name -ErrorAction Stop
}
Ensure-Module -Name Microsoft.Graph
Ensure-Module -Name ImportExcel
#endregion

#region Connect Graph
$scopes = @(
    "DeviceManagementConfiguration.Read.All",
    "DeviceManagementServiceConfig.Read.All"
)
if ($TenantId) {
    Connect-MgGraph -TenantId $TenantId -Scopes $scopes
} else {
    Connect-MgGraph -Scopes $scopes
}

# NOTE: configurationPolicies/settings are available in v1.0, but some fields
# are richer in beta. If needed, uncomment the beta profile:
# Select-MgProfile -Name "beta"
#endregion

#region Load CIS mapping from Excel ("Translation to DDM")
# Expected columns in the sheet: CIS Control, Level, Policy Control or Setting (Intune),
# Configuration Type, Intune Setting Name/Location, Recommended Setting, Description
$sheetName = "Translation to DDM"

if (-not (Test-Path $MappingPath)) {
    throw "Mapping workbook not found at: $MappingPath"
}

$rawMap = Import-Excel -Path $MappingPath -WorksheetName $sheetName -ErrorAction Stop

# Normalize mapping rows for iOS/iPadOS (Settings Catalog or Template - Device Restrictions)
$cisMap = $rawMap |
    Where-Object {
        $_.Level -in $Levels -and
        ($_.ConfigurationType -match 'Settings Catalog' -or $_.ConfigurationType -match 'Template')
    } |
    ForEach-Object {
        # Build a canonical record
        [pscustomobject]@{
            CisSection   = $_.'CIS Control'
            Level        = $_.Level
            IntuneName   = $_.'Intune Setting Name/Location' -replace '\s+', ' ' # display name
            ConfigType   = $_.'Configuration Type'
            Recommended  = $_.'Recommended Setting'
            Description  = $_.Description
        }
    }

# Basic heuristic for rows where "Recommended Setting" is blank:
function Infer-Recommended {
    param(
        [string]$IntuneName,
        [string]$Existing,
        [string]$Description
    )
    if ($Existing) { return $Existing }

    # Heuristics: Allow/Force/Block naming
    if ($IntuneName -match '^(Allow )') { return 'No' }
    if ($IntuneName -match '^(Force )') { return 'Yes' }
    if ($IntuneName -match '^(Allow .*While Locked|Allow .*For Unlock)') { return 'No' }
    if ($IntuneName -match '^(Allow .*Cloud|Allow .*Keychain|Allow .*Document|Allow .*Backup)') { return 'No' }
    if ($IntuneName -match '^(Allow .*Password Sharing|Allow .*Password Proximity)') { return 'No' }
    if ($IntuneName -match '^(Allow .*App Installation|Allow .*Enterprise App Trust)') { return 'No' }
    if ($IntuneName -match '^(Allow .*Photo|Allow .*Stream)') { return 'No' }
    if ($IntuneName -match '^(Allow .*VPN Creation)') { return 'No' }
    if ($IntuneName -match '^(Allow Assistant While Locked)') { return 'No' }
    if ($IntuneName -match '^(Allow Device Name Modification)') { return 'No' }
    if ($IntuneName -match '^(Allow Screen Shot)') { return 'No' }
    if ($IntuneName -match '^(Allow Untrusted TLS Prompt)') { return 'No' }
    if ($IntuneName -match '^(Allow UI Configuration Profile Installation)') { return 'No' }
    if ($IntuneName -match '^(Allow Open From Managed To Unmanaged)') { return 'No' }
    if ($IntuneName -match '^(Safari Force Fraud Warning)') { return 'Yes' }
    if ($IntuneName -match '^(Force Limit Ad Tracking)') { return 'Yes' }
    if ($IntuneName -match '^(Force Encrypted Backup|Force Automatic Date And Time)') { return 'Yes' }
    if ($IntuneName -match '^(Locked Screen Experience .*Wallet|Block Wallet notifications in lock screen)') { return 'Yes' }

    # Fall back on description hints
    if ($Description -match 'prevents' -or $Description -match 'disables') { return 'Yes' }

    return 'Unknown'
}

$cisMap = $cisMap | ForEach-Object {
    $_.Recommended = Infer-Recommended -IntuneName $_.IntuneName -Existing $_.Recommended -Description $_.Description
    $_
}

# Quick index for lookup by setting display name
$mapByDisplayName = @{}
foreach ($row in $cisMap) {
    $mapByDisplayName[$row.IntuneName] = $row
}
#endregion

#region Pull iOS/iPadOS setting definitions (to resolve display names <-> IDs)
$defsAll = Get-MgDeviceManagementSettingDefinition -All
# iOS platform filter (definition.Platforms can be array or string)
$defsIos = $defsAll | Where-Object {
    $_.Platforms -and ($_.Platforms -contains "iOS" -or $_.Platforms -like "*iOS*")
}

# Index by Id and by DisplayName
$defById   = @{}
$defByName = @{}
foreach ($d in $defsIos) {
    $defById[$d.Id]     = $d
    $defByName[$d.DisplayName] = $d
}
#endregion

#region Enumerate Settings Catalog policies for iOS/iPadOS
$policies = Get-MgDeviceManagementConfigurationPolicy -All |
    Where-Object { $_.Platforms -and ($_.Platforms -contains "iOS" -or $_.Platforms -like "*iOS*") }

Write-Host "Found $($policies.Count) iOS/iPadOS Settings Catalog configuration policies." -ForegroundColor Green

# Helper: normalize actual values from ValueJson shapes
function Normalize-Actual {
    param($valueJson)
    try {
        $obj = $valueJson
        if ($valueJson -is [string]) {
            $obj = $valueJson | ConvertFrom-Json -ErrorAction Stop
        }
        switch ($obj.GetType().Name) {
            'Hashtable' { 
                # pick common keys
                foreach ($k in @('value','BooleanValue','StringValue','IntValue','EnumValue','simpleValue')) {
                    if ($obj.ContainsKey($k)) { return $obj[$k] }
                }
                return ($obj | ConvertTo-Json -Compress)
            }
            default { return $obj }
        }
    } catch {
        return $valueJson
    }
}

# Helper: compare normalized actual vs recommended (Yes/No/strings/numbers)
function Compare-Recommended {
    param(
        [Parameter(Mandatory)]$Actual,
        [Parameter(Mandatory)][string]$Recommended
    )
    # Map Yes/No to $true/$false
    $rec = $Recommended.Trim()
    $recBool = $null
    if ($rec -match '^(Yes|True)$') { $recBool = $true }
    elseif ($rec -match '^(No|False)$') { $recBool = $false }

    # Actual may be boolean, string, int
    $act = $Actual
    if ($recBool -ne $null) {
        # Normalize actual to bool if possible
        if ($act -is [string]) {
            if ($act -match '^(true|yes)$') { $act = $true }
            elseif ($act -match '^(false|no)$') { $act = $false }
        }
        return ($act -eq $recBool)
    }

    # String/integer comparisons
    if ([int]::TryParse($rec, [ref]([int]$null))) {
        return ([int]$act -eq [int]$rec)
    }
    # Examples: "Alphanumeric", "Immediately"
    return ($act.ToString().Trim() -ieq $rec)
}

# Collect results
$detailed = New-Object System.Collections.Generic.List[object]

foreach ($p in $policies) {
    $settings = Get-MgDeviceManagementConfigurationPolicySetting -DeviceManagementConfigurationPolicyId $p.Id -All
    foreach ($s in $settings) {
        $defId     = $s.SettingDefinitionId
        $def       = $defById[$defId]
        if (-not $def) { continue }

        $display   = $def.DisplayName
        $actual    = Normalize-Actual -valueJson $s.ValueJson

        if ($mapByDisplayName.ContainsKey($display)) {
            $cis = $mapByDisplayName[$display]
            $isOk = $false
            if ($cis.Recommended -and $cis.Recommended -ne 'Unknown') {
                $isOk = Compare-Recommended -Actual $actual -Recommended $cis.Recommended
            } else {
                # If recommendation can't be inferred, mark Unknown
                $isOk = $false
            }

            $detailed.Add([pscustomobject]@{
                PolicyName      = $p.Name
                PolicyId        = $p.Id
                SettingDisplay  = $display
                ActualValue     = $actual
                Recommended     = $cis.Recommended
                CisSection      = $cis.CisSection
                Level           = $cis.Level
                ConfigType      = $cis.ConfigType
                Compliance      = if ($cis.Recommended -eq 'Unknown') { 'Unknown' } elseif ($isOk) { 'PASS' } else { 'FAIL' }
                Notes           = $cis.Description
            })
        } else {
            # Setting not mapped in the sheet — capture for analysis
            $detailed.Add([pscustomobject]@{
                PolicyName      = $p.Name
                PolicyId        = $p.Id
                SettingDisplay  = $display
                ActualValue     = $actual
                Recommended     = 'Not in mapping'
                CisSection      = ''
                Level           = ''
                ConfigType      = 'Settings Catalog'
                Compliance      = 'Unknown'
                Notes           = 'No row in Translation to DDM'
            })
        }
    }
}

#region Optional: legacy template-based iOS profiles
if ($IncludeLegacy) {
    $legacy = Get-MgDeviceManagementDeviceConfiguration -All |
        Where-Object { $_.OdataType -like "#microsoft.graph.ios*" }

    foreach ($l in $legacy) {
        # Legacy objects have properties rather than per-setting instances.
        # We’ll enumerate public properties and attempt fuzzy matches to CIS mapping.
        $props = $l | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name
        foreach ($pn in $props) {
            $val = $l.$pn
            if (-not $val) { continue }

            # Try fuzzy mapping by display name fragment
            $hit = $cisMap | Where-Object { $pn -replace '_' -like "*$($_.IntuneName.Split()[0])*" } | Select-Object -First 1
            if ($hit) {
                $isOk = Compare-Recommended -Actual $val -Recommended $hit.Recommended
                $detailed.Add([pscustomobject]@{
                    PolicyName      = $l.DisplayName
                    PolicyId        = $l.Id
                    SettingDisplay  = $pn
                    ActualValue     = $val
                    Recommended     = $hit.Recommended
                    CisSection      = $hit.CisSection
                    Level           = $hit.Level
                    ConfigType      = 'Template (legacy)'
                    Compliance      = if ($isOk) { 'PASS' } else { 'FAIL' }
                    Notes           = "Legacy template mapped heuristically"
                })
            } else {
                $detailed.Add([pscustomobject]@{
                    PolicyName      = $l.DisplayName
                    PolicyId        = $l.Id
                    SettingDisplay  = $pn
                    ActualValue     = $val
                    Recommended     = 'Not in mapping'
                    CisSection      = ''
                    Level           = ''
                    ConfigType      = 'Template (legacy)'
                    Compliance      = 'Unknown'
                    Notes           = 'No mapping for legacy property'
                })
            }
        }
    }
}
#endregion

#region Optional: iOS Compliance Policies (CIS §4.x)
if ($IncludeCompliancePolicies) {
    $comp = Get-MgDeviceManagementDeviceCompliancePolicy -All |
        Where-Object { $_.OdataType -eq "#microsoft.graph.iosCompliancePolicy" }

    foreach ($cp in $comp) {
        $props = $cp | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name
        foreach ($pn in $props) {
            $val = $cp.$pn
            if ($null -eq $val) { continue }

            $hit = $cisMap | Where-Object {
                $_.CisSection -like '4.%' -and $_.IntuneName -like "*$pn*"
            } | Select-Object -First 1

            if ($hit) {
                $isOk = Compare-Recommended -Actual $val -Recommended $hit.Recommended
                $detailed.Add([pscustomobject]@{
                    PolicyName      = $cp.DisplayName
                    PolicyId        = $cp.Id
                    SettingDisplay  = $pn
                    ActualValue     = $val
                    Recommended     = $hit.Recommended
                    CisSection      = $hit.CisSection
                    Level           = $hit.Level
                    ConfigType      = 'Compliance Policy'
                    Compliance      = if ($isOk) { 'PASS' } else { 'FAIL' }
                    Notes           = "Compliance policy check"
                })
            }
        }
    }
}
#endregion

#region Build Summary sheet
$summary =
    $detailed |
    Group-Object PolicyName |
    ForEach-Object {
        $name = $_.Name
        $rows = $_.Group
        $pass = ($rows | Where-Object { $_.Compliance -eq 'PASS' }).Count
        $fail = ($rows | Where-Object { $_.Compliance -eq 'FAIL' }).Count
        $unk  = ($rows | Where-Object { $_.Compliance -eq 'Unknown' }).Count
        $total = $rows.Count
        [pscustomobject]@{
            PolicyName     = $name
            TotalSettings  = $total
            PassCount      = $pass
            FailCount      = $fail
            UnknownCount   = $unk
            CompliancePct  = if ($total -gt 0) { [math]::Round(($pass / $total) * 100, 1) } else { 0 }
        }
    } |
    Sort-Object -Property CompliancePct -Descending

#endregion

#region Export to Excel
$excelParams = @{
    Path              = $OutputPath
    WorksheetName     = "Detailed"
    AutoSize          = $true
    FreezeTopRow      = $true
    BoldTopRow        = $true
}
$detailed | Export-Excel @excelParams

$summary | Export-Excel -Path $OutputPath -WorksheetName "Summary" -AutoSize -FreezeTopRow -BoldTopRow

Write-Host "Report written to: $OutputPath" -ForegroundColor Green
#endregion
