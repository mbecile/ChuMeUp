
# ---------------- CONFIGURATION ----------------
$AGY = "<AgencyAcronym>" # Agency acronym to filter mailboxes via CustomAttribute12
$AgencyDomain = "AGENCY.domain.com" # Agency domain to filter mailboxes via PrimarySmtpAddress
$DelegateUser = "<DelegateUserEmail>" # Email address of the user or Mail-Enabled Security Group to delegate permissions to
$Desktop = Join-Path $env:USERPROFILE 'Desktop'
$ExportPath = "$Desktop\CalendarPermissions.csv" # Path to export the CSV file

# ---------------- CONNECT TO EXCHANGE ONLINE ----------------
#Install-Module ExchangeOnlineManagement
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline -ShowBanner:$false

# ---------------- GET ALL MAILBOXES & UPDATE PERMISSIONS ----------------
$AllMailboxes = Get-EXOMailbox -ResultSize Unlimited | Where-Object { ($_.PrimarySmtpAddress -like "*@$AgencyDomain") -or ($_.CustomAttribute12 -eq '$AGY') }
foreach ($Mailbox in $AllMailboxes) {
    try {
        Add-MailboxFolderPermission -Identity ($Mailbox.PrimarySmtpAddress + ":\Calendar") -User $DelegateUser -AccessRights Reviewer -ErrorAction Stop
        Write-Host "Updated: $($Mailbox.PrimarySmtpAddress)"
        # Add success to results
        $Results += [PSCustomObject]@{
            PrimarySmtpAddress     = $Mailbox.PrimarySmtpAddress
            RecipientTypeDetails   = $Mailbox.RecipientTypeDetails
            Status                 = "Success"
            ErrorMessage           = $null
        }
        Write-Host "Updated: $($Mailbox.PrimarySmtpAddress) [$($Mailbox.RecipientTypeDetails)]"
    }
    catch {
        # Add failure to results
        $Results += [PSCustomObject]@{
            PrimarySmtpAddress     = $Mailbox.PrimarySmtpAddress
            RecipientTypeDetails   = $Mailbox.RecipientTypeDetails
            Status                 = "Failed"
            ErrorMessage           = $_.Exception.Message
        }

        Write-Warning "Failed for $($Mailbox.PrimarySmtpAddress): $_"
    }
}
# ---------------- EXPORT TO CSV ----------------
$Results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
Write-Host "Report exported to $ExportPath"

# ---------------- GET USER MAILBOXES ONLY & UPDATE PERMISSIONS ----------------
# Or adjust recipient types as needed
# UserMailbox,SharedMailbox,RoomMailbox,EquipmentMailbox

$UserMailboxes = Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox | Where-Object { ($_.PrimarySmtpAddress -like "*@$AgencyDomain") -or ($_.CustomAttribute12 -eq '$AGY') }
foreach ($Mailbox in $UserMailboxes) {
    try {
        Add-MailboxFolderPermission -Identity ($Mailbox.PrimarySmtpAddress + ":\Calendar") -User $DelegateUser -AccessRights Reviewer -ErrorAction Stop
        Write-Host "Updated: $($Mailbox.PrimarySmtpAddress)"
    }
    catch {
        Write-Warning "Failed for $($Mailbox.PrimarySmtpAddress): $_"
    }
}


# Install/Import Exchange Online module
#Install-Module ExchangeOnlineManagement
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline

# Output file
$ReportPath = "C:\Temp\SendCopy_Report.csv"

# Prepare a results array
$Results = @()

# Get mailbox set
$Mailboxes = Get-EXOMailbox -ResultSize Unlimited `
    -RecipientTypeDetails SharedMailbox,RoomMailbox,EquipmentMailbox,GroupMailbox |
    Where-Object {
        ($_.PrimarySmtpAddress -like '*@AGENCY.domain.com') -or
        ($_.CustomAttribute12 -eq 'AGENCY')
    }

foreach ($Mailbox in $Mailboxes) {

    try {
        Set-Mailbox -Identity $Mailbox.PrimarySmtpAddress `
            -MessageCopyForSentAsEnabled $true `
            -MessageCopyForSendOnBehalfEnabled $true

        # Add success to results
        $Results += [PSCustomObject]@{
            PrimarySmtpAddress     = $Mailbox.PrimarySmtpAddress
            RecipientTypeDetails   = $Mailbox.RecipientTypeDetails
            Status                 = "Success"
            ErrorMessage           = $null
        }

        Write-Host "Updated: $($Mailbox.PrimarySmtpAddress) [$($Mailbox.RecipientTypeDetails)]"
    }
    catch {
        # Add failure to results
        $Results += [PSCustomObject]@{
            PrimarySmtpAddress     = $Mailbox.PrimarySmtpAddress
            RecipientTypeDetails   = $Mailbox.RecipientTypeDetails
            Status                 = "Failed"
            ErrorMessage           = $_.Exception.Message
        }

        Write-Warning "Failed for $($Mailbox.PrimarySmtpAddress): $_"
    }
}

# Export results
$Results | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8

Write-Host "Report saved to $ReportPath"