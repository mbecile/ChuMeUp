Connect-MgGraph 
$context = Get-MgContext
$currentUser = (Get-MgUser -UserId $context.Account).Id

# Get all available roles
$myRoles = Get-MgRoleManagementDirectoryRoleEligibilitySchedule -ExpandProperty RoleDefinition -All -Filter "principalId eq '$currentuser'"

# Get Global Reader 
$myRole = $myroles | Where-Object {$_.RoleDefinition.DisplayName -eq "Global Reader"}

# Setup parameters for activation
$params = @{
    Action = "selfActivate"
    PrincipalId = $myRole.PrincipalId
    RoleDefinitionId = $myRole.RoleDefinitionId
    DirectoryScopeId = $myRole.DirectoryScopeId
    Justification = "Work."
    ScheduleInfo = @{
        StartDateTime = Get-Date
        Expiration = @{
            Type = "AfterDuration"
            Duration = "PT8H"
        }
    }
   }

# Activate the role
New-MgRoleManagementDirectoryRoleAssignmentScheduleRequest -BodyParameter $params

# Get Exch Admin 
$myRole = $myroles | Where-Object {$_.RoleDefinition.DisplayName -eq "Exchange Administrator"}

# Setup parameters for activation
$params = @{
    Action = "selfActivate"
    PrincipalId = $myRole.PrincipalId
    RoleDefinitionId = $myRole.RoleDefinitionId
    DirectoryScopeId = $myRole.DirectoryScopeId
    Justification = "Work."
    ScheduleInfo = @{
        StartDateTime = Get-Date
        Expiration = @{
            Type = "AfterDuration"
            Duration = "PT8H"
        }
    }
   }

# Activate the role
New-MgRoleManagementDirectoryRoleAssignmentScheduleRequest -BodyParameter $params