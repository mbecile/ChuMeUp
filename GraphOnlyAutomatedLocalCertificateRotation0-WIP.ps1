<#
AGENCY-AppName – Graph-only automated certificate rotation
- Creates new self-signed cert (LocalMachine\My)
- Exports Base-64 X.509 (.CER)
- Adds the cert to the Entra app using Microsoft Graph (addKey + PoP JWT)
- Updates Credential Manager (Generic) with the new thumbprint

Run in an elevated (Administrator) PowerShell session.
Requires Microsoft.Graph PowerShell SDK and Application.ReadWrite.All scope.
#>

[CmdletBinding()]
param(
    [string]$CertCommonName       = 'AGENCY-AppName',
    [string]$ExportCerPath        = 'C:\Desktop\Certificates\AGENCY-AppName.cer',

    # Tenant and app identifiers
    [string]$TenantId             = '<InsertM365TenantId>',
    [string]$AppClientId          = '<InsertM365AppClientId>',
    [string]$AppObjectId          = '<InsertM365AppObjectId>',

    # Lifetime of new cert
    [int]$ValidityMonths          = 12,

    # Windows Credential Manager (Generic) where the current/old thumbprint is stored
    [string]$CredentialTargetName = 'AGENCY-AppName-CertThumbprint',

    # Optional cleanup to remove older local certs with the same CN
    [switch]$RemoveOldLocalCertsAfterUpload
)

$ErrorActionPreference = 'Stop'

# --- Helpers ---------------------------------------------------------------

function Assert-Admin {
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()
               ).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')
    if (-not $isAdmin) { throw 'Run this script in an elevated (Administrator) PowerShell session.' }
}

# Read a Generic credential (Windows Credential Manager)
Add-Type -Namespace NativeCred -Name Win32 -MemberDefinition @"
using System;
using System.Runtime.InteropServices;

[StructLayout(LayoutKind.Sequential, CharSet=CharSet.Unicode)]
public struct CREDENTIAL {
    public uint Flags;
    public uint Type;
    public string TargetName;
    public string Comment;
    public System.Runtime.InteropServices.ComTypes.FILETIME LastWritten;
    public uint CredentialBlobSize;
    public IntPtr CredentialBlob;
    public uint Persist;
    public uint AttributeCount;
    public IntPtr Attributes;
    public string TargetAlias;
    public string UserName;
}

public static class Cred {
    [DllImport("advapi32", SetLastError=true, CharSet=CharSet.Unicode)]
    public static extern bool CredRead(string target, uint type, uint reservedFlag, out IntPtr credentialPtr);
    [DllImport("advapi32", SetLastError=true)]
    public static extern void CredFree(IntPtr buffer);
    [DllImport("advapi32", SetLastError=true, CharSet=CharSet.Unicode)]
    public static extern bool CredWrite(ref CREDENTIAL Credential, uint Flags);
    [DllImport("kernel32.dll")]
    public static extern uint GetLastError();
}
"@

function Get-GenericCredential {
    param([Parameter(Mandatory)][string]$Target)
    $CRED_TYPE_GENERIC = 1
    $ptr = :Zero
    $ok = [NativeCred.Win32+Cred]::CredRead($Target, $CRED_TYPE_GENERIC, 0, [ref]$ptr)
    if (-not $ok) { return $null }
    try {
        $cred = [NativeCred.Win32+CREDENTIAL]
        $size = [int]$cred.CredentialBlobSize
        $blob = if ($size -gt 0 -and $cred.CredentialBlob -ne :Zero) {
            $bytes = New-Object byte[] $size
            [System.Runtime.InteropServices.Marshal]::Copy($cred.CredentialBlob, $bytes, 0, $size)
            [System.Text.Encoding]::Unicode.GetString($bytes)
        } else { "" }
        [pscustomobject]@{
            TargetName = $cred.TargetName
            Persist    = $cred.Persist
            Secret     = $blob
            UserName   = $cred.UserName
        }
    } finally {
        if ($ptr -ne :Zero) { [NativeCred.Win32+Cred]::CredFree($ptr) }
    }
}

function Set-GenericCredential {
    param(
        [Parameter(Mandatory)][string]$Target,
        [Parameter(Mandatory)][string]$Secret,
        [string]$UserName = "",
        [ValidateSet(0,1,2,3)] [int]$Persist = 2 # 2 = LocalMachine; 1 = LocalProfile; 3 = Enterprise
    )

    $bytes = [System.Text.Encoding]::Unicode.GetBytes($Secret)
    $ptr   = [System.Runtime.InteropServices.Marshal]::AllocCoTaskMem($bytes.Length)
    [System.Runtime.InteropServices.Marshal]::Copy($bytes, 0, $ptr, $bytes.Length)

    $cred = New-Object NativeCred.Win32+CREDENTIAL
    $cred.Type               = 1                 # CRED_TYPE_GENERIC
    $cred.TargetName         = $Target
    $cred.CredentialBlobSize = [uint32]$bytes.Length
    $cred.CredentialBlob     = $ptr
    $cred.Persist            = [uint32]$Persist
    $cred.UserName           = $UserName

    $ok = [NativeCred.Win32+Cred]::CredWrite([ref]$cred, 0)
    [System.Runtime.InteropServices.Marshal]::FreeCoTaskMem($ptr)

    if (-not $ok) {
        $code = [NativeCred.Win32+Cred]::GetLastError()
        throw "CredWrite failed. Win32Error=$code"
    }
}

function Ensure-GraphModule {
    if (-not (Get-Module Microsoft.Graph.Applications -ListAvailable)) {
        Write-Host "Installing Microsoft.Graph.Applications PowerShell SDK..." -ForegroundColor Yellow
        Install-Module Microsoft.Graph.Applications -Scope CurrentUser -Force
    }
    Import-Module Microsoft.Graph.Applications -ErrorAction Stop
}

# --- Run -------------------------------------------------------------------

Assert-Admin

# Prepare export folder
$destFolder = Split-Path -Path $ExportCerPath -Parent
if (-not (Test-Path $destFolder)) {
    New-Item -ItemType Directory -Path $destFolder -Force | Out-Null
}

# Create new self-signed certificate in LocalMachine\My
$subject = "CN=$CertCommonName"
$start   = Get-Date
$end     = $start.AddMonths($ValidityMonths)

Write-Host "==> Creating new self-signed certificate: $subject" -ForegroundColor Cyan
$cert = New-SelfSignedCertificate -Subject $subject `
          -CertStoreLocation "Cert:\LocalMachine\My" `
          -KeyExportPolicy Exportable `
          -KeySpec Signature `
          -KeyLength 2048 `
          -KeyAlgorithm RSA `
          -HashAlgorithm SHA256 `
          -NotBefore $start `
          -NotAfter  $end

Write-Host ("    New thumbprint: {0}" -f $cert.Thumbprint)

# Export as Base-64 encoded X.509 (.CER) — overwrite if present
if (Test-Path $ExportCerPath) {
    Write-Host "==> Removing pre-existing file: $ExportCerPath" -ForegroundColor Yellow
    Remove-Item $ExportCerPath -Force
}

Write-Host "==> Exporting Base-64 (.CER) to $ExportCerPath" -ForegroundColor Cyan
Export-Certificate -Cert $cert -FilePath $ExportCerPath -CertEncoding Base64 | Out-Null

# Compute Base64 DER for Graph (use in-memory RawData to avoid PEM headers)
$base64Der = :ToBase64String($cert.RawData)

# Connect to Microsoft Graph (requires Application.ReadWrite.All)
Ensure-GraphModule
Write-Host "==> Connecting to Microsoft Graph (tenant: $TenantId)" -ForegroundColor Cyan
Connect-MgGraph -TenantId $TenantId -Scopes Application.ReadWrite.All

# Pull application (to get existing keyCredentials)
Write-Host "==> Reading application keyCredentials (App Object Id: $AppObjectId)" -ForegroundColor Cyan
$app = Get-MgApplication -ApplicationId $AppObjectId -Property Id,AppId,KeyCredentials

# Retrieve the old/current thumbprint from Credential Manager (to locate signing cert for PoP JWT)
Write-Host "==> Reading current thumbprint from Credential Manager: $CredentialTargetName" -ForegroundColor Cyan
$existingCred = Get-GenericCredential -Target $CredentialTargetName
if (-not $existingCred -or :IsNullOrWhiteSpace($existingCred.Secret)) {
    throw "Cannot read existing Generic credential '$CredentialTargetName'. A prior certificate thumbprint is required to build PoP JWT."
}

$oldThumb = $existingCred.Secret.Trim()
Write-Host ("    Current (old) thumbprint: {0}" -f $oldThumb)

# Load the existing certificate from LocalMachine\My (must have private key)
$store = New-Object System.Security.Cryptography.X509Certificates.X509Store('My','LocalMachine')
$store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly)
try {
    $signingCert = $store.Certificates | Where-Object { $_.Thumbprint -eq $oldThumb }
} finally { $store.Close() }

if (-not $signingCert) {
    throw "Existing cert with thumbprint $oldThumb not found in LocalMachine\My. Cannot build PoP JWT."
}
if (-not $signingCert.HasPrivateKey) {
    throw "Existing cert with thumbprint $oldThumb does not have a private key. Cannot build PoP JWT."
}

# Map the existing local cert to the app's keyCredentials (by customKeyIdentifier = Base64(SHA1 hash))
$oldHashBase64 = :ToBase64String($signingCert.GetCertHash())
$existingKey = $app.KeyCredentials | Where-Object { $_.CustomKeyIdentifier -eq $oldHashBase64 }

if (-not $existingKey) {
    throw "No matching keyCredential found in the app for the existing certificate (customKeyIdentifier=$oldHashBase64). Ensure the app currently has this certificate configured."
}

# Build a proof-of-possession JWT signed by the existing certificate private key
Write-Host "==> Building PoP JWT for addKey" -ForegroundColor Cyan

# Load JWT assembly
$null = [System.Reflection.Assembly]::LoadWithPartialName('System.IdentityModel.Tokens.Jwt')
$null = [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.IdentityModel.Tokens')

# Create SigningCredentials with KeyId = existing keyCredential.KeyId
$rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($signingCert)
$rsaKey = [Microsoft.IdentityModel.Tokens.RsaSecurityKey]::new($rsa)
$rsaKey.KeyId = $existingKey.KeyId.ToString()

$signingCreds = [Microsoft.IdentityModel.Tokens.SigningCredentials]::new(
    $rsaKey, 
    [Microsoft.IdentityModel.Tokens.SecurityAlgorithms]::RsaSha256
)

$handler = [System.IdentityModel.Tokens.Jwt.JwtSecurityTokenHandler]::new()
$claims  = [System.Security.Claims.ClaimsIdentity]::new(@(
    [System.Security.Claims.Claim]::new('sub', $AppClientId)   # subject = application (client) id
))

$descriptor = [Microsoft.IdentityModel.Tokens.SecurityTokenDescriptor]::new()
$descriptor.Issuer            = $AppClientId
$descriptor.Audience          = 'https://graph.microsoft.com'
$descriptor.NotBefore         = :UtcNow
$descriptor.Expires           = :UtcNow.AddMinutes(10)
$descriptor.Subject           = $claims
$descriptor.SigningCredentials= $signingCreds

$token = $handler.CreateJwtSecurityToken($descriptor)
$proofJwt = $handler.WriteToken($token)

# Build the new keyCredential payload (Base64 DER)
$newKeyCredential = @{
    Type            = 'AsymmetricX509Cert'
    Usage           = 'Verify'
    Key             = $base64Der
    StartDateTime   = $start.ToUniversalTime()
    EndDateTime     = $end.ToUniversalTime()
}

# Add the new certificate to the application (Graph addKey)
Write-Host "==> Adding new certificate to application (Graph addKey)" -ForegroundColor Cyan
$result = Add-MgApplicationKey -ApplicationId $AppObjectId `
          -KeyCredential $newKeyCredential `
          -Proof $proofJwt

Write-Host ("    Added Key (KeyId): {0}" -f $result.KeyId)

# Update Windows Credential Manager with the new thumbprint
Write-Host "==> Updating Credential Manager: $CredentialTargetName" -ForegroundColor Cyan
Set-GenericCredential -Target $CredentialTargetName -Secret $cert.Thumbprint -Persist 2
Write-Host ("    Stored new thumbprint: {0}" -f $cert.Thumbprint)

# Optional cleanup: remove older local certs with same CN (keep newest)
if ($RemoveOldLocalCertsAfterUpload.IsPresent) {
    Write-Host "==> Cleaning up older LocalMachine\My certs with Subject: $subject" -ForegroundColor Cyan
    $all = Get-ChildItem Cert:\LocalMachine\My | Where-Object { $_.Subject -eq $subject }
    $toRemove = $all | Where-Object { $_.Thumbprint -ne $cert.Thumbprint }
    if ($toRemove) {
        $toRemove | Remove-Item -Force
        Write-Host ("    Removed {0} older certificate(s)." -f ($toRemove | Measure-Object).Count)
    } else {
        Write-Host "    No older certificates found to remove."
    }
}

Write-Host "==> Rotation completed successfully." -ForegroundColor Green

# Summary
Write-Host ""
Write-Host "Summary:"
Write-Host ("  CN: {0}" -f $CertCommonName)
Write-Host ("  New Thumbprint: {0}" -f $cert.Thumbprint)
Write-Host ("  Exported Base-64 .CER: {0}" -f $ExportCerPath)
Write-Host ("  App Client ID: {0}" -f $AppClientId)
Write-Host ("  App Object ID: {0}" -f $AppObjectId)
Write-Host ("  Validity (UTC): {0} -> {1}" -f $start.ToUniversalTime(), $end.ToUniversalTime())