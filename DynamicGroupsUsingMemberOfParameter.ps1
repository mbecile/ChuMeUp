# Find DynamicMembership Groups using 'memberOf' operator that is being discontinued
#Install-Module Microsoft.Graph.Groups
#Install-Module Microsoft.Graph.Identity.DirectoryManagement
Import-Module Microsoft.Graph.Groups
Import-Module Microsoft.Graph.Identity.DirectoryManagement
Connect-MgGraph -Scopes "Group.Read.All", "AdministrativeUnit.Read.All"
# Retrieve all dynamic groups and examine their membership rules
$dynamicGroups = Get-MgGroup -All `
  -Filter "groupTypes/any(c:c eq 'DynamicMembership')" `
  -Select "DisplayName,Id,MembershipRule"
# Keep only groups whose rule uses the memberOf operator
$memberOfGroups = $dynamicGroups | Where-Object {
  $_.MembershipRule -match 'memberOf'
}
# Display results
$memberOfGroups | Select-Object DisplayName, Id, MembershipRule | Format-Table -AutoSize


# Find and display Administrative Units using 'memberOf' operator that is being discontinued
Get-MgDirectoryAdministrativeUnit -All -Property Id, DisplayName, MembershipRule, MembershipType |
    Where-Object {
        $_.MembershipType -eq "Dynamic" -and
        $_.MembershipRule -match "memberOf"
    } |
    Select-Object DisplayName, Id, MembershipRule | Format-Table -AutoSize