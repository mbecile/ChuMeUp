Connect-MgGraph -Scopes "Device.Read.All"

$hybridDevices = Get-MgDevice -All | Where-Object {
    $_.TrustType -eq "ServerAd" -and
    $_.ProfileType -eq "RegisteredDevice"
}

foreach ($dev in $hybridDevices) {
    Invoke-Command -ComputerName $dev.DisplayName -ScriptBlock {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned -Force
        Install-Script -Name Get-WindowsAutopilotInfo -Force
        Get-WindowsAutopilotInfo -Online
    }
}

# --------------- SPECIFIC GROUPS ---------------
Connect-MgGraph -Scopes "Group.Read.All","Device.Read.All"

# Get your group's object ID
$groupId = (Get-MgGroup -Filter "DisplayName eq 'DEVICE-GROUP-NAME'").Id

# Get all members in the group and filter for device objects
$devices = Get-MgGroupMember -GroupId $groupId -All |
           Where-Object { $_.'@odata.type' -eq '#microsoft.graph.device' }

# Further filter to Hybrid Joined devices
$hybridDevices = foreach ($dev in $devices) {
    $d = Get-MgDevice -DeviceId $dev.Id
    if ($d.TrustType -eq "ServerAd" -and $d.ProfileType -eq "RegisteredDevice") { $d }
}

foreach ($dev in $hybridDevices) {
    Invoke-Command -ComputerName $dev.DisplayName -ScriptBlock {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned -Force
        Install-Script -Name Get-WindowsAutopilotInfo -Force
        Get-WindowsAutopilotInfo -Online
    }
}
