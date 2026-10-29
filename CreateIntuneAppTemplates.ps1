#Install-Module Microsoft.Graph.Users
#Install-Module Microsoft.Graph.Groups
Import-Module Microsoft.Graph.Users
Import-Module Microsoft.Graph.Groups
Connect-MgGraph -Scopes "Group.ReadWrite.All", "User.Read.All"

$ownerUpn = "<USERNAME@DOMAIN.COM>"
$ownerId  = (Get-MgUser -UserId $ownerUpn).Id

$groupNames = @(
"INTUNE_App_AGENCY_Template1",
"INTUNE_App_AGENCY_Template2",
"INTUNE_App_AGENCY_Template3",
"INTUNE_App_AGENCY_Template4",
"INTUNE_App_AGENCY_Template5"
)

foreach ($gn in $groupNames) {

    # Create the group
    $newGroup = New-MgGroup -DisplayName $gn `
        -MailEnabled:$false `
        -MailNickname $gn.Replace("-","").Replace("_","") `
        -SecurityEnabled

    Write-Host "Created group: $($newGroup.DisplayName)  ID:$($newGroup.Id)"

    # Assign owner
    New-MgGroupOwnerByRef -GroupId $newGroup.Id `
        -OdataId "https://graph.microsoft.com/v1.0/users/$ownerId"

    Write-Host "Assigned owner $ownerUpn to group $gn"
}