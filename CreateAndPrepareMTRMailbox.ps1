$MTRMailbox = ConfRoom01@yourtenant.onmicrosoft.com

Connect-ExchangeOnline
Set-CalendarProcessing -Identity "ConfRoom01@yourtenant.onmicrosoft.com" `
    -AutomateProcessing AutoAccept `
    -BookingWindowInDays 740 `
    -MaximumDurationInMinutes 550000 `
    -AddOrganizerToSubject $false `
    -DeleteComments $false `
    -DeleteAttachments $false `
    -DeleteSubject $false `
    -ProcessExternalMeetingMessages $false `
    -RemovePrivateProperty $false `
    -DeleteNonCalendarItems $false `
    -EnableResponseDetails $true `
    -TentativePendingApproval $true `
    -RemoveOldMeetingMessages $true `
    -AddNewRequestsTentatively $true `
    -ForwardRequestsToDelegates $true `
    -AllBookInPolicy $true `
    -AddAdditionalResponse $true `
    -AdditionalResponse "This is the Teams Room for ConfRoom01@yourtenant.onmicrosoft.com."


# Create Cloud-Only Room Mailbox for use with Microsoft Teams Rooms Device
Connect-MgGraph -Scopes "User.ReadWrite.All","Directory.ReadWrite.All","Organization.Read.All","Directory.AccessAsUser.All"
New-Mailbox -Room `
    -Name "ConfRoom-01" `
    -DisplayName "Conference Room 01" `
    -Alias "ConfRoom01" `
    -MicrosoftOnlineServicesID "MTR-$AGY.$RoomName@stateoforegon.onmicrosoft.com" `
    -Password (ConvertTo-SecureString "RoomPasswordHere123!" -AsPlainText -Force)

Update-MgUser -UserId "ConfRoom01@yourtenant.onmicrosoft.com" -PasswordPolicies "DisablePasswordExpiration" -AccountEnabled $true
Set-MgUserLicense -UserId "ConfRoom01@yourtenant.onmicrosoft.com" `
    -AddLicenses @("ecc74eae-eeb7-4ad5-9c88-e8b2bfca75b8")



Connect-MgGraph -Scopes "User.ReadWrite.All","Directory.ReadWrite.All","Organization.Read.All","Directory.AccessAsUser.All"

# Install modules
Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
Install-Module Microsoft.Graph -Scope CurrentUser -Force
Install-Module Microsoft.Graph.Beta -Scope CurrentUser -Force  # only if needed

# Connect
Connect-ExchangeOnline
Connect-MgGraph -Scopes "User.ReadWrite.All","Directory.ReadWrite.All","Organization.Read.All","Directory.AccessAsUser.All"

# Create room mailbox
New-Mailbox -Room `
    -Name "ConfRoom01" `
    -DisplayName "Conference Room 01" `
    -Alias "ConfRoom01" `
    -MicrosoftOnlineServicesID "ConfRoom01@yourtenant.onmicrosoft.com" `
    -Password (ConvertTo-SecureString "RoomPass123!" -AsPlainText -Force)

# License room
$sku = (Get-MgSubscribedSku | Where-Object {$_.SkuPartNumber -like "*TEAMSROOM*"}).SkuId
Set-MgUserLicense -UserId "ConfRoom01@yourtenant.onmicrosoft.com" -AddLicenses @($sku)

# Disable password expiration
Update-MgUser -UserId "ConfRoom01@yourtenant.onmicrosoft.com" -PasswordPolicies "DisablePasswordExpiration"

# Ensure account enabled
Update-MgUser -UserId "ConfRoom01@yourtenant.onmicrosoft.com" -AccountEnabled $true

# Calendar processing
Set-CalendarProcessing -Identity "ConfRoom01@yourtenant.onmicrosoft.com" `
    -AutomateProcessing AutoAccept `
    -BookingWindowInDays 740 `
    -MaximumDurationInMinutes 550000 `
    -AddOrganizerToSubject $false `
    -DeleteComments $false `
    -DeleteAttachments $false `
    -DeleteSubject $false `
    -ProcessExternalMeetingMessages $false `
    -RemovePrivateProperty $false `
    -DeleteNonCalendarItems $false `
    -EnableResponseDetails $true `
    -TentativePendingApproval $true `
    -RemoveOldMeetingMessages $true `
    -AddNewRequestsTentatively $true `
    -ForwardRequestsToDelegates $true `
    -AllBookInPolicy $true `
    -AddAdditionalResponse $true `
    -AdditionalResponse "This is the Teams Room for Conference Room 01."

# Create room list
New-DistributionGroup -Name "MainBuilding" -PrimarySmtpAddress "MainBuilding@yourtenant.onmicrosoft.com" -RoomList

# Add room to list
Add-DistributionGroupMember -Identity "MainBuilding@yourtenant.onmicrosoft.com" -Member "ConfRoom01@yourtenant.onmicrosoft.com"

# Verify place
Get-Place -Identity "ConfRoom01@yourtenant.onmicrosoft.com" | fl
