<#
    GCC High Exchange Online Diagnostic Script
    -------------------------------------
    Purpose:
      - Validate PowerShell version & TLS capabilities
      - Validate ExchangeOnlineManagement module health
      - Validate network access to GCC High endpoints (HEAD requests)
      - Validate EXO REST connectivity using a lightweight EXO cmdlet
      - Detect common causes of "The underlying connection was closed"
#>

Write-Host "`n=== GCC High Diagnostic Script ===`n" -ForegroundColor Cyan

#############################
# 1. PowerShell / TLS Check #
#############################

Write-Host "Checking PowerShell version..."
$psVer = $PSVersionTable.PSVersion
Write-Host "PowerShell version: $psVer"

if ($PSVersionTable.PSEdition -eq "Desktop") {
    Write-Host "Desktop edition detected (Windows PowerShell). Forcing TLS 1.2..."
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Write-Host "TLS now set to: $([Net.ServicePointManager]::SecurityProtocol)"
} else {
    Write-Host "PowerShell 7+ detected. TLS handled automatically." -ForegroundColor Green
}

############################################
# 2. ExchangeOnlineManagement Module Check #
############################################

Write-Host "`nChecking ExchangeOnlineManagement module..."
$exo = Get-Module ExchangeOnlineManagement -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1

if ($exo) {
    Write-Host "Module found: $($exo.Name) v$($exo.Version)" -ForegroundColor Green
} else {
    Write-Warning "ExchangeOnlineManagement module NOT found. Install with:"
    Write-Host "Install-Module ExchangeOnlineManagement -Force"
}

##########################################
# 3. GCC Endpoint Reachability Test (HEAD)
##########################################

Write-Host "`nTesting GCC High endpoints reachability..."
$endpoints = @(
    "https://outlook.office365.us",
    "https://login.microsoftonline.us",
    "https://graph.microsoft.us",
    "https://outlook.office365.us/EWS/Exchange.asmx",
    "https://outlook.office365.us/adminapi/beta/mailbox"
)

foreach ($ep in $endpoints) {
    Write-Host "Checking $ep ..."
    try {
        $resp = Invoke-WebRequest -Uri $ep -Method Head -TimeoutSec 10
        Write-Host "  Reachable (HTTP $($resp.StatusCode))" -ForegroundColor Green
    }
    catch {
        Write-Warning "  UNREACHABLE or INSPECTED — $($_.Exception.Message)"
    }
}

##########################################
# 4. EXO REST Connectivity Test (GCC)
##########################################

Write-Host "`nTesting Exchange Online REST connectivity..."

try {
    Disconnect-ExchangeOnline -Confirm:$false 2>$null

    Connect-ExchangeOnline -ExchangeEnvironmentName O365USGov -ShowProgress:$false

    Write-Host "Connected successfully to GCC EXO." -ForegroundColor Green

    # Lightweight EXO REST cmdlet
    $test = Get-EXORecipient -ResultSize 1
    if ($test) {
        Write-Host "REST call succeeded. Sample recipient:" -ForegroundColor Green
        $test | Format-Table DisplayName, RecipientTypeDetails
    } else {
        Write-Warning "REST call returned no data but did not error."
    }

    Disconnect-ExchangeOnline -Confirm:$false
}
catch {
    Write-Warning "`nEXO REST FAILURE:"
    Write-Warning $_.Exception.Message
    Write-Host "`nCommon causes:" -ForegroundColor Yellow
    Write-Host "  - SSL inspection breaking REST traffic"
    Write-Host "  - Proxy not bypassing *.office365.us and login.microsoftonline.us"
    Write-Host "  - Outdated ExchangeOnlineManagement module"
    Write-Host "  - PS 5.1 without TLS 1.2 forced"
}

Write-Host "`n=== Diagnostics Complete ===`n" -ForegroundColor Cyan