
<#
COPY OWNERSHIP OF GROUPS - Interactive Script
.SYNOPSIS
Copies group ownership from a source user to a target user across:
- Microsoft 365 (Unified) groups             -> Graph owners
- Security groups                            -> Graph owners
- Distribution groups                        -> Exchange "ManagedBy"
- Mail-Enabled Security groups               -> Exchange "ManagedBy"
.SKIPS
- Dynamic membership groups
.NOTES
- Requires Microsoft.Graph and ExchangeOnlineManagement modules.
- In hybrid environments, synced groups must be managed on-prem.
#>
# -------------------------------
# Set UPN variables
# -------------------------------
$SourceUserUpn = "username1@agency.domain.com"
$TargetUserUpn = "username2@agency.domain.com"
# -------------------------------
# Connect to Microsoft Graph
# -------------------------------
Connect-MgGraph -ContextScope Process `
  -Scopes "User.Read.All","Group.ReadWrite.All"
# Modules
Import-Module Microsoft.Graph.Users
Import-Module Microsoft.Graph.Groups
# -------------------------------
# Connect to Exchange Online for DG/MESG ownership
# -------------------------------
# Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue
try {
    if (-not (Get-Module ExchangeOnlineManagement)) {
        Write-Warning "ExchangeOnlineManagement module not loaded; DG/MESG ownership updates will be skipped."
    } else {
        Connect-ExchangeOnline -ShowProgress $true
    }
} catch {
    Write-Warning "Exchange Online connection failed: $($_.Exception.Message). DG/MESG ownership updates will be skipped."
}
# -------------------------------
# Resolve users (grab GUIDs)
# -------------------------------
$SourceUser = Get-MgUser -UserId $SourceUserUpn
$TargetUser = Get-MgUser -UserId $TargetUserUpn
# -------------------------------
# Helper: classify group type from Graph properties
# -------------------------------
function Get-GroupInfo {
    param([Parameter(Mandatory)][object]$DirObj)
    $p = $DirObj.AdditionalProperties
    $isUnified  = ($p.groupTypes -contains "Unified")
    $isDynamic  = ($p.groupTypes -contains "DynamicMembership")
    $mail       = $p.mail
    $secEnabled = [bool]$p.securityEnabled
    $mailEnabled= [bool]$p.mailEnabled
    $type = if ($isUnified) { "Microsoft365" }
            elseif ($mailEnabled -and -not $secEnabled) { "Distribution" }
            elseif ($mailEnabled -and $secEnabled) { "MailEnabledSecurity" }
            elseif ($secEnabled) { "Security" }
            else { "Unknown" }
    [PSCustomObject]@{
        Id              = $DirObj.Id
        DisplayName     = $p.displayName
        Mail            = $mail
        GroupTypes      = $p.groupTypes
        SecurityEnabled = $secEnabled
        MailEnabled     = $mailEnabled
        IsDynamic       = $isDynamic
        Type            = $type
    }
}
# -------------------------------
# Get all groups the source user owns (Graph)
# -------------------------------
$AllOwned = Get-MgUserOwnedObject -UserId $SourceUser.Id -All | Where-Object {
    $_.Id -ne $null -and $_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.group'
}
# Exclude dynamic groups and classify
$Groups = $AllOwned |
    ForEach-Object { Get-GroupInfo $_ } |
    Where-Object { -not $_.IsDynamic }
# Buckets
$M365Groups   = $Groups | Where-Object { $_.Type -eq "Microsoft365" }
$SecGroups    = $Groups | Where-Object { $_.Type -eq "Security" }
$DistGroups   = $Groups | Where-Object { $_.Type -eq "Distribution" }
$MesgGroups   = $Groups | Where-Object { $_.Type -eq "MailEnabledSecurity" }
Write-Host "Found: $($M365Groups.Count) M365, $($SecGroups.Count) Security, $($DistGroups.Count) Distribution, $($MesgGroups.Count) Mail-Enabled Security (excluding Dynamic)."
# -------------------------------
# Add Graph owners for M365 & Security groups
# -------------------------------
foreach ($g in ($M365Groups + $SecGroups)) {
    try {
        $existing = Get-MgGroupOwner -GroupId $g.Id -All |
                    Where-Object { $_.Id -eq $TargetUser.Id }
        if (-not $existing) {
            New-MgGroupOwnerByRef -GroupId $g.Id -BodyParameter @{
                "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($TargetUser.Id)"
            }
            Write-Host "[Graph] Added $($TargetUser.DisplayName) as owner of '$($g.DisplayName)' ($($g.Type))."
        } else {
            Write-Host "[Graph] Already owner: $($TargetUser.DisplayName) on '$($g.DisplayName)' ($($g.Type))."
        }
    } catch {
        Write-Warning "[Graph] Failed on '$($g.DisplayName)' ($($g.Id)): $($_.Exception.Message)"
    }
}
# -------------------------------
# Update Exchange 'ManagedBy' for Distribution & Mail-Enabled Security groups
# -------------------------------
$exchangeConnected = (Get-Module ExchangeOnlineManagement) -and (Get-ConnectionInformation)
if ($exchangeConnected) {
    # Resolve the target as an Exchange recipient to use a stable identity value
    $TargetRecipient = $null
    try { $TargetRecipient = Get-Recipient -Identity $TargetUserUpn -ErrorAction Stop } catch { }
    if (-not $TargetRecipient) {
        Write-Warning "Could not resolve Target recipient in Exchange for '$TargetUserUpn'. Skipping DG/MESG ownership updates."
    } else {
        foreach ($g in ($DistGroups + $MesgGroups)) {
            # Prefer the group's primary SMTP address when present
            $dgIdentity = if ($g.Mail) { $g.Mail } else { $g.DisplayName }
            try {
                # Check current owners
                $dg = Get-DistributionGroup -Identity $dgIdentity -ErrorAction Stop
                $existingOwners = $dg.ManagedBy
                # If already present, skip
                $alreadyOwner = $false
                foreach ($owner in $existingOwners) {
                    if ($owner -eq $TargetRecipient.Identity -or $owner.PrimarySmtpAddress -eq $TargetUserUpn) { 
                        $alreadyOwner = $true; break 
                    }
                }
                if (-not $alreadyOwner) {
                    # Add owner and bypass manager check if needed
                    Set-DistributionGroup -Identity $dg.Identity `
                        -ManagedBy @{ Add = $TargetRecipient.Identity } `
                        -BypassSecurityGroupManagerCheck:$true
                    Write-Host "[EXO] Added $($TargetUser.DisplayName) to ManagedBy of '$($dg.DisplayName)' ($($g.Type))."
                } else {
                    Write-Host "[EXO] Already ManagedBy: $($TargetUser.DisplayName) on '$($dg.DisplayName)' ($($g.Type))."
                }
            } catch {
                Write-Warning "[EXO] Failed on '$($dgIdentity)': $($_.Exception.Message)"
            }
        }
    }
} else {
    Write-Warning "Exchange Online not connected. Distribution and Mail-Enabled Security group ownership (ManagedBy) not updated."
}
Write-Host "Ownership copy complete."
