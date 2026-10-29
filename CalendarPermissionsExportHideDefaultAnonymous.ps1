Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline -ShowBanner:$false

$mailboxes = Get-EXOMailbox -ResultSize Unlimited `
  -RecipientTypeDetails UserMailbox,SharedMailbox,RoomMailbox,EquipmentMailbox `
  -Filter "CustomAttribute12 -eq 'AGENCY'"

$Results = foreach ($Mailbox in $mailboxes) {
    $CalendarPath = "$($Mailbox.PrimarySmtpAddress):\Calendar"
    $Permissions = Get-MailboxFolderPermission -Identity $CalendarPath -ErrorAction SilentlyContinue

    foreach ($Perm in $Permissions) {
        if ($Perm.User -in @('Default','Anonymous')) { continue }
        $Rec = Get-Recipient -Identity $Perm.User -ErrorAction SilentlyContinue
        $smtp = if ($Rec) { $Rec.PrimarySmtpAddress } else { $Perm.User.ToString() }

        [PSCustomObject]@{
            Mailbox              = $Mailbox.PrimarySmtpAddress
            MailboxDisplayName   = $Mailbox.DisplayName
            Grantee              = $smtp
            GranteeDisplayName   = $Perm.User.ToString()
            AccessRights         = ($Perm.AccessRights -join ',')
            SharingPermissionFlags = $Perm.SharingPermissionFlags
            RecipientTypeDetails = $Mailbox.RecipientTypeDetails
        }
    }
}

$desktop = Join-Path $env:USERPROFILE 'Desktop'
$Results | Export-Csv -Path "$desktop\AGENCY_calendar_permissions.csv" -NoTypeInformation -Encoding UTF8
Write-Host "Export complete: $desktop\AGENCY_calendar_permissions.csv"