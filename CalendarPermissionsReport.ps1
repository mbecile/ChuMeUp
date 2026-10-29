<#
.SYNOPSIS
Export calendar permissions for all mailboxes in a specific agency and email the CSV to the requestor.
.REQUIREMENTS
- SMTP relay or credentials for Send-MailMessage
- Appropriate permissions to run Get-EXOMailbox and Get-MailboxFolderPermission
#>
# ---------------- CONFIGURATION ----------------
#$AgencyAcronym = "DAS" # Agency acronym to filter mailboxes
$ExportPath = "$env:TEMP\CalendarPermissions.csv"
$SmtpServer = "relay.domain.com" # SMTP server
$FromAddress = "noreply@domain.com" # Sender email
#$Requestor = "Your.Admin@domain.com" # Recipient email
$Subject = "Calendar Permissions Report"
$Body = "Attached is the latest calendar permissions report."

# ---------------- GET MAILBOXES ----------------
# Write-Host "Retrieving mailboxes for $AgencyAcronym..."
$mailboxes = Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox,SharedMailbox,RoomMailbox,EquipmentMailbox -Filter "CustomAttribute12 -eq '$AgencyAcronym'"

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
# Write-Host "Report exported to $ExportPath"

# ---------------- EMAIL THE REPORT ----------------
Send-MailMessage -From $FromAddress -To $Requestor -Subject $Subject -Body $Body -SmtpServer $SmtpServer -Attachments $ExportPath -ErrorAction Stop
# Write-Host "Email sent to $Requestor"