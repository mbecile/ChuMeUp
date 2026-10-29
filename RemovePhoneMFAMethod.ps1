#Install-Module Microsoft.Graph.Reports
#Install-Module Microsoft.Graph.Identity.SignIns
Import-Module Microsoft.Graph.Reports
Import-Module Microsoft.Graph.Identity.SignIns
Connect-MgGraph -Scopes "AuditLog.Read.All","UserAuthenticationMethod.ReadWrite.All"

# Retrieve full MFA registration overview for all users
$RegistrationDetails = Get-MgReportAuthenticationMethodUserRegistrationDetail -All

# Identify users with at least one phone method
$UsersWithPhone = $RegistrationDetails |
    Where-Object { $_.MethodsRegistered -match "phone" }

# Remove the discovered phone authentication methods
foreach ($user in $UsersWithPhone) {
    $methods = Get-MgUserAuthenticationMethod -UserId $user.Id
    foreach ($method in $methods) {
        if ($method.AdditionalProperties['@odata.type'] -eq "#microsoft.graph.phoneAuthenticationMethod") {
    Remove-MgUserAuthenticationPhoneMethod -UserId $user.Id -PhoneAuthenticationMethodId $method.Id
        }
    }
}