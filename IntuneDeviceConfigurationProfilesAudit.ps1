# Get all device configuration profiles (Settings Catalog and Templates)
$profiles = Get-MgDeviceManagementDeviceConfiguration

# Prepare an array to hold all settings
$allSettings = @()

foreach ($profile in $profiles) {
    # Get settings for each profile (beta endpoint provides more detail)
    $settings = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations/$($profile.Id)/settings"
    foreach ($setting in $settings.value) {
        $allSettings += [PSCustomObject]@{
            ProfileName   = $profile.displayName
            ProfileType   = $profile.'@odata.type'
            SettingName   = $setting.displayName
            SettingValue  = $setting.value
            SettingType   = $setting.'@odata.type'
        }
    }
}

# Export to CSV
$allSettings | Export-Csv -Path ".\\Intune_DeviceConfiguration_Settings_Beta.csv" -NoTypeInformation

Write-Host "Export complete: Intune_DeviceConfiguration_Settings_Beta.csv"