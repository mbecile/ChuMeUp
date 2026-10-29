# ---------------------------------------------
# Validate iOS/iPadOS Settings Catalog profiles
# Against expected recommended settings
# ---------------------------------------------

Import-Module Microsoft.Graph -ErrorAction Stop

# Connect to Graph
Connect-MgGraph -Scopes "DeviceManagementConfiguration.Read.All"

# Expected values
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

# Retrieve all iOS/iPadOS Settings Catalog policies
$Policies = Get-MgDeviceManagementConfigurationPolicy -All |
    Where-Object { $_.Platforms -contains "iOS" -or $_.Platforms -like "*iOS*" }

$Results = @()

foreach ($Policy in $Policies) {

    $Settings = Get-MgDeviceManagementConfigurationPolicySetting `
        -DeviceManagementConfigurationPolicyId $Policy.Id -All

    foreach ($Setting in $Settings) {

        $Actual = $Setting.ValueJson
        if ($Actual -is [string]) {
            try { $Actual = $Actual | ConvertFrom-Json } catch {}
        }

        # Normalize: pick common value keys
        if ($Actual -isnot [string]) {
            foreach ($key in "value","simpleValue","StringValue","BooleanValue","IntValue","EnumValue") {
                if ($Actual.$key) {
                    $Actual = $Actual.$key
                    break
                }
            }
        }

        $Name = $Setting.AdditionalProperties.displayName

        if ($Expected.ContainsKey($Name)) {

            $ExpectedValue = $Expected[$Name]

            $Pass = $Actual.ToString().Trim().ToLower() -eq $ExpectedValue.ToString().Trim().ToLower()

            $Results += [pscustomobject]@{
                PolicyName     = $Policy.Name
                PolicyId       = $Policy.Id
                Setting        = $Name
                ActualValue    = $Actual
                ExpectedValue  = $ExpectedValue
                Compliance     = if ($Pass) { "PASS" } else { "FAIL" }
            }
        }
    }
}

$Results | Format-Table -AutoSize
$Results | Export-Csv "iOS_CIS_Check.csv" -NoTypeInformation

Write-Host "`nResults exported to iOS_CIS_Check.csv"