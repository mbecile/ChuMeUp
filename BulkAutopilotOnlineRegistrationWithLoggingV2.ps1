# ============================================================
# Bulk Autopilot Online Registration with Logging
# ============================================================

Connect-MgGraph -Scopes "Device.Read.All"

# Identify all Hybrid Joined devices
#$hybridDevices = Get-MgDevice -All | Where-Object {
#    $_.TrustType -eq "ServerAd" -and
#    $_.ProfileType -eq "RegisteredDevice"
#}

# Get the device group's object ID
$groupId = (Get-MgGroup -Filter "DisplayName eq 'YourGroupName'").Id

# Filter group members for device objects only
$GroupDevices = Get-MgGroupMember -GroupId $groupId -All |
           Where-Object { $_.'@odata.type' -eq '#microsoft.graph.device' }

# Further filter to Hybrid Joined devices
$hybridDevices = foreach ($dev in $GroupDevices) {
    $d = Get-MgDevice -DeviceId $dev.Id
    if ($d.TrustType -eq "ServerAd" -and $d.ProfileType -eq "RegisteredDevice") { $d }
}

# Log file paths
$logCSV = "AutopilotUploadLog.csv"
$deviceLogFolder = "DeviceLogs"
New-Item -ItemType Directory -Force -Path $deviceLogFolder | Out-Null

# Prepare CSV
"DeviceName,DeviceId,Timestamp,Result,ErrorMessage" | Out-File $logCSV

foreach ($device in $hybridDevices) {

    $deviceName = $device.DisplayName
    $deviceId   = $device.Id
    $ts         = (Get-Date).ToString("s")

    $logPath = Join-Path $deviceLogFolder "$deviceName.log"

    Write-Host "Processing $deviceName ($deviceId)..."

    try {

        Invoke-Command -ComputerName $deviceName -ScriptBlock {

            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned -Force
            Install-Script -Name Get-WindowsAutopilotInfo -Force

            # Try uploading directly to Intune
            Get-WindowsAutopilotInfo -Online

            # Return success indicator
            return "SUCCESS"
        } -ErrorAction Stop | Out-File $logPath

        Add-Content -Path $logCSV -Value "$deviceName,$deviceId,$ts,SUCCESS,""None"""
        Write-Host "SUCCESS: $deviceName"

    }
    catch {
        $err = $_.Exception.Message
        $err | Out-File $logPath

        Add-Content -Path $logCSV -Value "$deviceName,$deviceId,$ts,FAILED,""$err"""
        Write-Warning "FAILED: $deviceName — $err"
    }
}

Write-Host "Completed Autopilot upload processing. Results saved to $logCSV."