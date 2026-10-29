<#
    Commercial / GCC Exchange Online Diagnostic Script
    -------------------------------------------------------
    Purpose:
      - Validate PowerShell version & TLS
      - Validate ExchangeOnlineManagement module
      - Test commercial Microsoft 365 endpoints
      - Validate REST connectivity to Exchange Online (commercial)
#>

Write-Host "`n=== Commercial Exchange Online Diagnostic Script ===`n" -ForegroundColor Cyan

#############################
# 1. PowerShell / TLS Check #
#############################

Write-Host "Checking PowerShell version..."
$psVer = $PSVersionTable.PSVersion
Write-Host "PowerShell version: $psVer"

if ($PSVersionTable.PSEdition -eq "Desktop") {
    Write-Host "Windows PowerShell detected. Forcing TLS 1.2..."
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
    Write-Warning "ExchangeOnlineManagement module NOT found."
    Write-Host "Install with: Install-Module ExchangeOnlineManagement -Force"
}

##########################################
# 3. Commercial Endpoint Reachability    #
##########################################

Write-Host "`nTesting commercial Microsoft 365 endpoints..."

$endpoints = @(
    "https://outlook.office365.com",           # EXO commercial
    "https://login.microsoftonline.com",       # AAD commercial
    "https://graph.microsoft.com",             # Graph API commercial
    "https://outlook.office365.com/EWS/Exchange.asmx",
    "https://outlook.office365.com/adminapi/beta/mailbox"
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
# 4. EXO REST Connectivity Test          #
##########################################

Write-Host "`nTesting commercial Exchange Online REST connectivity..."

try {
    Disconnect-ExchangeOnline -Confirm:$false 2>$null

    Connect-ExchangeOnline -ShowProgress:$false

    Write-Host "Connected successfully to Exchange Online." -ForegroundColor Green

    # Lightweight REST test
    $test = Get-EXORecipient -ResultSize 1
    if ($test) {
        Write-Host "REST call succeeded. Sample recipient:" -ForegroundColor Green
        $test | Format-Table DisplayName, RecipientTypeDetails
    } else {
        Write-Warning "REST call returned no data but did not error."
    }
}
catch {
    Write-Warning "`nEXO REST FAILURE:"
    Write-Warning $_.Exception.Message
    Write-Host "`nCommon causes:" -ForegroundColor Yellow
    Write-Host "  - SSL inspection breaking REST traffic"
    Write-Host "  - Proxy not bypassing commercial endpoints (*.office365.com, login.microsoftonline.com)"
    Write-Host "  - Outdated ExchangeOnlineManagement module"
    Write-Host "  - Windows PowerShell not enforcing TLS 1.2"
}

Write-Host "`n=== Diagnostics Complete ===`n" -ForegroundColor Cyan