<#
.SYNOPSIS
Export calendar permissions for all mailboxes in a specific agency
.REQUIREMENTS
- Appropriate permissions to run Get-EXOMailbox and Get-MailboxFolderPermission
#>
# ---------------- CONFIGURATION ----------------
$AGY = "<AgencyAcronym>" # Agency acronym to filter mailboxes via CustomAttribute12
$AgencyDomain = "AGENCY.domain.com" # Agency domain to filter mailboxes via PrimarySmtpAddress
$Desktop = Join-Path $env:USERPROFILE 'Desktop'
$ExportPath = "$Desktop\$AGY_CalendarPermissions.csv" # Path to export the CSV file

# ---------------- CONNECT TO EXCHANGE ONLINE ----------------
#Install-Module ExchangeOnlineManagement
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline -ShowBanner:$false 

# ---------------- GET MAILBOXES ----------------
$Mailboxes = Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox,SharedMailbox,RoomMailbox,EquipmentMailbox | Where-Object { ($_.PrimarySmtpAddress -like "*@$AgencyDomain") -or ($_.CustomAttribute12 -eq '$AGY') }

# ---------------- GET CALENDAR PERMISSIONS ----------------
$Results = @()
foreach ($Mailbox in $Mailboxes) {
$CalendarPath = "$($Mailbox.PrimarySmtpAddress):\Calendar"
$Permissions = Get-MailboxFolderPermission -Identity $CalendarPath -ErrorAction Stop
foreach ($Perm in $Permissions) {
if ($Perm.User -in @('Default','Anonymous')) { continue }
$Rec = Get-Recipient -Identity $Perm.User -ErrorAction SilentlyContinue
$smtp = if ($Rec) { $Rec.PrimarySmtpAddress } else { $Perm.User.ToString() }
$Results += [PSCustomObject]@{
RecipientTypeDetails = $Mailbox.RecipientTypeDetails
Mailbox              = $Mailbox.PrimarySmtpAddress
MailboxDisplayName   = $Mailbox.DisplayName
Folder               = "Calendar"
Grantee              = $smtp
GranteeDisplayName   = $Perm.User.ToString()
AccessRights         = ($Perm.AccessRights -join ',')
SharingPermissionFlags = $Perm.SharingPermissionFlags
}
}
}
# ---------------- EXPORT TO CSV ----------------
$Results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
Write-Host "Report exported to $ExportPath"