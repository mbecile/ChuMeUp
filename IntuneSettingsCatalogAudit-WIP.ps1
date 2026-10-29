# Intune Settings Catalog
# Install required modules if needed
# Install-Module Microsoft.Graph -Scope CurrentUser

# Connect to Graph with the right scopes
Connect-MgGraph -Scopes "DeviceManagementConfiguration.Read.All"
Select-MgProfile -Name "beta"  # Settings Catalog APIs are typically under beta

# Helper: GET with paging
function Invoke-GraphPagedJson {
    param([string]$Uri)
    $results = @()
    $resp = Invoke-MgGraphRequest -Uri $Uri -Method GET -OutputType PSObject
    $results += $resp.value
    $next = $resp.'@odata.nextLink'
    while ($next) {
        $resp = Invoke-MgGraphRequest -Uri $next -Method GET -OutputType PSObject
        $results += $resp.value
        $next = $resp.'@odata.nextLink'
    }
    return $results
}

# 1) Get all Settings Catalog policies (MDM tech)
$policies = Invoke-GraphPagedJson -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies?`$filter=technologies has 'mdm'"

# 2) For each policy, get its settings collection and flatten
$rows = foreach ($p in $policies) {

    # NOTE: The settings endpoint is available for settings catalog policies; replace with the correct path if needed in your tenant.
    $settings = Invoke-GraphPagedJson -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/$($p.id)/settings"

    foreach ($s in $settings) {
        # Each $s typically includes settingDefinitionId, settingInstance, and value payloads for the setting
        # Flatten some common fields; adjust as needed for your environment.
        [pscustomobject]@{
            PolicyName   = $p.name
            PolicyId     = $p.id
            Platform     = $p.platforms
            SettingName  = $s.settingDefinitionId
            CategoryPath = ($s.settingDefinition?.categoryPath)  # if available
            Value        = ($s.settingInstance?.simpleSettingValue?.value ?? $s.settingInstance?.choiceSettingValue?.value ?? $s.settingInstance?.complexSettingValue)
        }
    }
}

# 3) Export to CSV and JSON for analysis
$timestamp = Get-Date -Format "yyyyMMdd-HHmm"
$csvPath = "$PWD/Intune_SettingsCatalogMap_$timestamp.csv"
$jsonPath = "$PWD/Intune_SettingsCatalogMap_$timestamp.json"

$rows | Export-Csv -NoTypeInformation -Path $csvPath
$rows | ConvertTo-Json -Depth 6 | Out-File $jsonPath -Encoding utf8

Write-Host "Export complete:"
Write-Host "CSV:  $csvPath"
