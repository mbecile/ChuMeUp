#Install-Module ExchangeOnlineManagement
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline

# Output file
$ReportPath = "C:\Temp\SendCopy_Report.csv"

# Prepare results array
$Results = @()

# Get mailbox set
$Mailboxes = Get-EXOMailbox -ResultSize Unlimited `
    -RecipientTypeDetails SharedMailbox,RoomMailbox,EquipmentMailbox,GroupMailbox |
    Where-Object {
        ($_.PrimarySmtpAddress -like '*@DOMAIN.COM') -or
        ($_.CustomAttribute12 -eq 'AGENCY')
    }

# Enable Send-Copy for each mailbox and log results
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