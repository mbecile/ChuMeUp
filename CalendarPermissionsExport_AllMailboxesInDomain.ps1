# 1) Connect to Exchange Online
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline -ShowBanner:$false

# 2) Get all relevant mailboxes (user/shared/room/equipment) in agency.domain.com
$mailboxes = Get-EXOMailbox -ResultSize Unlimited `
  -RecipientTypeDetails UserMailbox,SharedMailbox,RoomMailbox,EquipmentMailbox |
  Where-Object { $_.PrimarySmtpAddress -like '*@agency.domain.com' }

# 3) Enumerate Calendar permissions for each mailbox in the agency
$results = foreach ($mbx in $mailboxes) {
    $calendarIdentity = "$($mbx.PrimarySmtpAddress):\Calendar"
    $perms = Get-MailboxFolderPermission -Identity $calendarIdentity -ErrorAction SilentlyContinue
    foreach ($perm in $perms) {
        # Skip built-in entries to avoid filler
        if ($perm.User -in @('Default','Anonymous')) { continue }
        # Resolve principal to recipient, if possible
        $rec = Get-Recipient -Identity $perm.User -ErrorAction SilentlyContinue
        # Determine best-available representation of the grantee
        $smtp = if ($rec) {
            $rec.PrimarySmtpAddress
        } else {
            $perm.User.ToString()
        }
        
        [PSCustomObject]@{
            Mailbox              = $mbx.PrimarySmtpAddress
            MailboxDisplayName   = $mbx.DisplayName
            Grantee              = $smtp
            GranteeRaw           = $perm.User.ToString()
            AccessRights         = ($perm.AccessRights -join ',')
            IsInherited          = $perm.IsInherited
            FolderPath           = $calendarIdentity
            RecipientTypeDetails = $mbx.RecipientTypeDetails
        }
    }
}

# 4) Export to CSV
$desktop = Join-Path $env:USERPROFILE 'Desktop'
$results | Export-Csv -Path "$desktop\agency_calendar_permissions.csv" -NoTypeInformation -Encoding UTF8