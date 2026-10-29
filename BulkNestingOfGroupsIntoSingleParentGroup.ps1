Import-Module Microsoft.Graph.Groups
Connect-MgGraph -Scopes "Group.ReadWrite.All"

$agencies = @(
    "BIZ","BOLI","BSPA","CCB","DAS","DCBS","DELC","DEQ","DLCD","DOGAMI",
    "DOR","DPSST","DSL","HECC","LCB","LTCO","OBD","OBMI","OBNM","OBOP",
    "OBPT","OCB","ODA","ODF","ODFW","ODOE","ODVA","OED","OEM",
    "OHCS","OLCC","OMB","OMD","OPDC","OPRD","OPSC","ORC","OSMB",
    "OTIC","OTLB","OVMEB","OWB","PUC","SLO","WRD"
)

foreach ($agency in $agencies) {

    $groupName = "MEM_Updates_{0}_Enforced" -f $agency
    $nick = $groupName.Replace("_","")

    $params = @{
        DisplayName     = $groupName
        SecurityEnabled = $true
        MailEnabled     = $false
        MailNickname    = $nick
        GroupTypes      = @()
    }

    $newGroup = New-MgGroup @params
    Write-Host "Created group: $groupName  ObjectId: $($newGroup.Id)"
}

$parentGroupId = "<InsertParentGroupId>" # Replace with the actual ObjectId of the parent group "MEM_Users_Updates_Enforced"

foreach ($agency in $agencies) {

    $childGroupName = "MEM_Updates_{0}_Enforced" -f $agency
    $childGroup = Get-MgGroup -Filter "displayName eq '$childGroupName'"

    if ($childGroup) {
        New-MgGroupMember -GroupId $parentGroupId -DirectoryObjectId $childGroup.Id
        Write-Host "Added $childGroupName to MEM_Users_Updates_Enforced"
    }
    else {
        Write-Host "Group not found: $childGroupName"
    }
}