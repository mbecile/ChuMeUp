<#
.BULK FORMS OWNERSHIP TRANSFER
Author: EIS M365 Team
Date: 2026-05-19
Purpose: Discover all Microsoft Forms for a given user and transfer ownership in bulk.

Phases:
  1) Discovery via OneDrive (Apps/Microsoft Forms)
  2) Transfer via Graph (beta Forms permissions) if available
  3) Optional UI automation fallback using Playwright

References:
  - Internal KB: KB XX -M How to change the owner for a Microsoft Forms (Owner-to-Group preferred; 30-day rule; Owner-to-Owner requires active account)
  - NOTE: Graph Forms endpoints are in beta in many tenants and may not be enabled.

Usage examples:
  # Preferred: Transfer all forms owned by user to a M365 Group
  .\Bulk-Transfer-UserForms.ps1 -SourceUserUpn "user@agency.gov" -TargetGroupId "00000000-0000-0000-0000-000000000000" -Mode Group -DryRun

  # Owner-to-owner (only if the source is still active per KB)
  .\Bulk-Transfer-UserForms.ps1 -SourceUserUpn "user@agency.gov" -TargetUserUpn "new.owner@agency.gov" -Mode User

#>

param(
  [Parameter(Mandatory=$true)]
  [string] $SourceUserUpn,

  [Parameter(Mandatory=$false)]
  [string] $TargetGroupId,      # Required if -Mode Group

  [Parameter(Mandatory=$false)]
  [string] $TargetUserUpn,      # Required if -Mode User

  [Parameter(Mandatory=$true)]
  [ValidateSet('Group','User')]
  [string] $Mode,               # 'Group' (preferred) or 'User'

  [switch] $DryRun,             # No changes; just report

  [string] $OutCsv = ".\Forms-Discovery-$((Get-Date).ToString('yyyyMMdd-HHmmss')).csv",
  [string] $LogPath = ".\BulkFormsTransfer-$((Get-Date).ToString('yyyyMMdd-HHmmss')).log"
)

# ---------------------------
# Utility: Logging
# ---------------------------
function Write-Log {
  param([string]$Message,[string]$Level="INFO")
  $line = "[{0}] [{1}] {2}" -f (Get-Date).ToString("s"), $Level.ToUpper(), $Message
  Write-Host $line
  Add-Content -Path $LogPath -Value $line
}

# ---------------------------
# Connect to Microsoft Graph
# ---------------------------
function Connect-GraphScopes {
  param([string[]]$Scopes)
  try {
    Write-Log "Connecting to Microsoft Graph with scopes: $($Scopes -join ', ')"
    Connect-MgGraph -Scopes $Scopes -NoWelcome
    $profile = Get-MgContext
    Write-Log "Connected. Tenant: $($profile.TenantId), App: $($profile.AppName)"
  } catch {
    Write-Log "Failed to connect to Graph: $($_.Exception.Message)" "ERROR"
    throw
  }
}

# ---------------------------
# Discovery: OneDrive Forms footprint
# ---------------------------
function Get-UserFormsFromOneDrive {
  param([string]$UserUpn)

  # Scopes needed: Files.Read.All, User.Read.All
  Connect-GraphScopes -Scopes @("Files.Read.All","User.Read.All")

  # Get user's drive
  $drive = Get-MgUserDrive -UserId $UserUpn
  if (-not $drive) {
    Write-Log "No OneDrive found for $UserUpn" "WARN"
    return @()
  }

  # Locate "Apps/Microsoft Forms" folder
  Write-Log "Locating 'Apps' and 'Microsoft Forms' folders for $UserUpn"
  $appsFolder = Get-MgUserDriveRootChild -UserId $UserUpn -Filter "name eq 'Apps'" -All
  if (-not $appsFolder) { return @() }

  $formsFolder = Get-MgDriveItemChild -DriveId $drive.Id -ItemId $appsFolder.Id -Filter "name eq 'Microsoft Forms'" -All
  if (-not $formsFolder) { return @() }

  Write-Log "Enumerating items under 'Apps/Microsoft Forms' for $UserUpn"
  $items = Get-MgDriveItemChild -DriveId $drive.Id -ItemId $formsFolder.Id -All

  # Build candidate list (best-effort)
  $candidates = @()
  foreach ($i in $items) {
    # Many tenants store response workbooks (.xlsx) and form assets here.
    $isResponseXlsx = $i.File -and $i.Name -like "*.xlsx"
    $formTitle = [System.IO.Path]::GetFileNameWithoutExtension($i.Name)

    $candidates += [pscustomobject]@{
      SourceUserUpn   = $UserUpn
      DriveId         = $drive.Id
      ItemId          = $i.Id
      ItemName        = $i.Name
      FormTitleGuess  = $formTitle
      PossibleFormId  = $null # We attempt to enrich in the next step if possible
      Path            = ($i.ParentReference.Path + "/" + $i.Name)
      Type            = if ($isResponseXlsx) {"ResponseWorkbook"} else {"Asset"}
    }
  }

  return $candidates
}

# ---------------------------
# Enrichment: Try to resolve FormId (best-effort)
# NOTE: There is no guaranteed way to map every workbook to its FormId,
# but some tenants embed IDs in file metadata or sibling JSON files.
# ---------------------------
function Try-ResolveFormIds {
  param([object[]]$Candidates, [string]$DriveId)

  Write-Log "Attempting to resolve Form IDs from available metadata (best-effort)"
  foreach ($c in $Candidates) {
    # Attempt to read workbook properties or inspect sibling files
    try {
      $props = Get-MgDriveItem -DriveId $DriveId -ItemId $c.ItemId -Property "name,webUrl,createdDateTime,lastModifiedDateTime,size"
      # Heuristics go here (left as placeholders; tenant implementations vary).
      # If you maintain a registry mapping titles->FormIds, enrich here.
      $c.PossibleFormId = $null
    } catch {
      Write-Log "Failed metadata read for $($c.ItemName): $($_.Exception.Message)" "WARN"
    }
  }
  return $Candidates
}

# ---------------------------
# Transfer: Graph beta endpoint (if available)
# ---------------------------
function Transfer-FormOwnershipGraphBeta {
  param(
    [string] $FormId,
    [ValidateSet('Group','User')]
    [string] $Mode,
    [string] $TargetGroupId,
    [string] $TargetUserUpn
  )

  # Connect with transfer scopes
  $scopes = @("Forms.ReadWrite.All")
  if ($Mode -eq "Group") { $scopes += "Group.ReadWrite.All" }
  Connect-GraphScopes -Scopes $scopes

  $body = if ($Mode -eq "Group") {
    @{
      "grantee" = @{
        "@odata.type" = "microsoft.graph.group"
        "id" = $TargetGroupId
      };
      "role" = "owner"
    }
  } else {
    @{
      "grantee" = @{
        "@odata.type" = "microsoft.graph.user"
        "userPrincipalName" = $TargetUserUpn
      };
      "role" = "owner"
    }
  }

  $uri = "https://graph.microsoft.com/beta/forms/$FormId/permissions"
  Write-Log "Posting transfer for FormId $FormId to $Mode target"
  if ($script:DryRun) {
    Write-Log "Dry-run: would call $uri with body: $($body | ConvertTo-Json -Depth 5)"
    return
  }

  try {
    Invoke-MgGraphRequest -Method POST -Uri $uri -Body $body | Out-Null
    Write-Log "Transfer invoked for FormId $FormId"
  } catch {
    Write-Log "Transfer failed for FormId $FormId: $($_.Exception.Message)" "ERROR"
    throw
  }
}

# ---------------------------
# Fallback: UI Automation with Playwright (bulk via Admin portal)
# ---------------------------
function Transfer-FormsViaUiAutomation {
  param(
    [object[]] $FormList,
    [ValidateSet('Group','User')] [string] $Mode,
    [string] $TargetGroupId,
    [string] $TargetUserUpn,
    [string] $AdminUpn,
    [string] $AdminPassword
  )

  # Install and initialize Playwright once (requires dotnet; run with pwsh)
  # pwsh -c "dotnet tool install --global Microsoft.Playwright.CLI; playwright install"

  # The automation below is a placeholder outline; implementation depends
  # on your tenant’s Forms Admin Center DOM.
  Write-Log "Starting UI automation fallback (Playwright)."
  Write-Log "NOTE: You will need to adapt DOM selectors to your Admin Center instance."

  # Pseudo-steps:
  # 1) Launch Chromium
  # 2) Login to admin.microsoft.com with $AdminUpn / $AdminPassword
  # 3) Navigate to Forms Admin Center
  # 4) For each form in $FormList:
  #       - Search form by title/owner
  #       - Open transfer dialog
  #       - Select Group/User target
  #       - Click 'Proceed' (as per KB)
  # 5) Log results

  Write-Log "UI automation is not executed in this sample to avoid exposing credentials in code."
}

# ---------------------------
# MAIN
# ---------------------------
$script:DryRun = [bool]$DryRun

if ($Mode -eq 'Group' -and [string]::IsNullOrWhiteSpace($TargetGroupId)) {
  Write-Log "TargetGroupId is required for Group mode." "ERROR"; exit 1
}
if ($Mode -eq 'User' -and [string]::IsNullOrWhiteSpace($TargetUserUpn)) {
  Write-Log "TargetUserUpn is required for User mode." "ERROR"; exit 1
}

Write-Log "Starting bulk discovery for $SourceUserUpn"
$candidates = Get-UserFormsFromOneDrive -UserUpn $SourceUserUpn
$candidates = Try-ResolveFormIds -Candidates $candidates -DriveId ($candidates | Select-Object -First 1).DriveId

# Export discovery for review
Write-Log "Discovered $($candidates.Count) candidate items; exporting to $OutCsv"
$candidates | Export-Csv -Path $OutCsv -NoTypeInformation

# Ask admin to provide FormIds (if discovery couldn't resolve them all)
$inputCsvPath = Read-Host "If you have a CSV mapping FormTitle -> FormId, enter its path (or press Enter to skip)"
$formMappings = @()
if (-not [string]::IsNullOrWhiteSpace($inputCsvPath) -and (Test-Path $inputCsvPath)) {
  $formMappings = Import-Csv $inputCsvPath # Expected columns: FormTitle,FormId
  Write-Log "Loaded $($formMappings.Count) mappings from $inputCsvPath"
}

# Build final transfer list
$transferList = @()
foreach ($c in $candidates) {
  $formId = $c.PossibleFormId
  if (-not $formId -and $formMappings.Count -gt 0) {
    $match = $formMappings | Where-Object { $_.FormTitle -eq $c.FormTitleGuess }
    if ($match) { $formId = $match.FormId }
  }
  if ($formId) {
    $transferList += [pscustomobject]@{
      FormTitle = $c.FormTitleGuess
      FormId    = $formId
      Source    = $SourceUserUpn
      Mode      = $Mode
      Target    = if ($Mode -eq "Group") { $TargetGroupId } else { $TargetUserUpn }
    }
  }
}

Write-Log "Forms with resolvable IDs: $($transferList.Count)"
if ($transferList.Count -eq 0) {
  Write-Log "No FormIds resolved. Consider supplying a CSV mapping or use UI automation fallback." "WARN"
  exit 0
}

# Execute transfers (Graph beta)
foreach ($t in $transferList) {
  try {
    Transfer-FormOwnershipGraphBeta -FormId $t.FormId -Mode $t.Mode -TargetGroupId $TargetGroupId -TargetUserUpn $TargetUserUpn
  } catch {
    Write-Log "Graph transfer failed for $($t.FormTitle) ($($t.FormId)). Consider UI automation fallback." "ERROR"
  }
}

Write-Log "Bulk transfer completed."