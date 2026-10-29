# Requires: ExchangeOnlineManagement module
Import-Module ExchangeOnlineManagement

# ---------------------------------------
# Configuration (serial-only execution)
# ---------------------------------------
$RetryCount        = 3      # Get-Place retry attempts per identity
$RetryDelaySeconds = 2      # Delay between retries
$ReconnectOnce     = $true  # Allow one forced reconnect on the specific HttpResponseMessage error
$UseGovGCCH        = $false # Set $true if your tenant is GCC High (most state tenants are O365Default)
$UseDoD            = $false # Set $true if your tenant is DoD

# ---------------------------------------
# Helper: Connect (with environment awareness)
# ---------------------------------------
function Ensure-ExchangeSession {
    param(
        [switch]$Reconnect
    )

    try {
        if ($Reconnect) {
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        }

        $envName = 'O365Default'
        if ($UseGovGCCH) { $envName = 'O365USGovGCCHigh' }
        if ($UseDoD)     { $envName = 'O365USGovDoD' }

        Connect-ExchangeOnline -ExchangeEnvironmentName $envName -ShowBanner:$false
    }
    catch {
        Write-Warning "Reconnect failed: $($_.Exception.Message)"
    }
}

# Initial connect
Ensure-ExchangeSession

# ---------------------------------------
# Helper: Safe join for arrays / nulls
# ---------------------------------------
function Join-IfAny {
    param(
        [Parameter(Mandatory=$false)] $Value,
        [string] $Separator = ", "
    )

    if ($null -eq $Value) { return $null }
    if ($Value -is [string]) { return $Value }

    try {
        if ($Value.Count -gt 0) { return ($Value -join $Separator) }
        else { return $null }
    }
    catch { return $Value }
}

# ---------------------------------------
# Helper: Try Get-Place with intelligent fallback and retry
# ---------------------------------------
function Get-PlaceWithRetry {
    param(
        [Parameter(Mandatory=$true)][psobject] $Room,   # Pass the mailbox object
        [int] $MaxRetries = 3,
        [int] $DelaySeconds = 2
    )

    $errorMessage = $null
    $place        = $null

    # Try identities in order: PrimarySmtpAddress, Alias, DisplayName
    $identities = @(
        $Room.PrimarySmtpAddress,
        $Room.Alias,
        $Room.DisplayName
    ) | Where-Object { $_ -and $_.ToString().Trim() -ne '' } | Select-Object -Unique

    foreach ($identity in $identities) {
        for ($i = 1; $i -le $MaxRetries -and -not $place; $i++) {
            try {
                $place = Get-Place -Identity $identity -ErrorAction Stop
            }
            catch {
                $errorMessage = $_.Exception.Message

                # If we hit the specific HttpResponseMessage/GetResponseHeader issue, try one reconnect (once only)
                if ($ReconnectOnce -and $errorMessage -match 'HttpResponseMessage.*GetResponseHeader') {
                    Write-Warning "Encountered HttpResponseMessage/GetResponseHeader. Attempting session reconnect..."
                    $script:ReconnectOnce = $false  # avoid looping reconnects
                    Ensure-ExchangeSession -Reconnect
                    Start-Sleep -Seconds 2
                    continue  # next loop iteration will retry the same identity
                }

                if ($i -lt $MaxRetries) { Start-Sleep -Seconds $DelaySeconds }
            }
        }

        if ($place) { break } # stop if we got it for any identity
    }

    [PSCustomObject]@{
        Place = $place
        Error = if ($place) { $null } else { $errorMessage }
    }
}

# ---------------------------------------
# Collect all Room Mailboxes
# ---------------------------------------
Write-Host "Collecting Room Mailboxes and their associated properties" -ForegroundColor Cyan    
$AllRooms = Get-Mailbox -RecipientTypeDetails RoomMailbox -ResultSize Unlimited

# ---------------------------------------
# Process Rooms (serial-only)
# ---------------------------------------
$results = foreach ($Room in $AllRooms) {
    # Retry-aware Get-Place with identity fallbacks
    $placeResult = Get-PlaceWithRetry -Room $Room -MaxRetries $RetryCount -DelaySeconds $RetryDelaySeconds
    $place       = $placeResult.Place
    $errorText   = $placeResult.Error

    if (-not $place) {
        Write-Warning "Get-Place failed for $($Room.PrimarySmtpAddress): $errorText"
    }

    # Build output object
    $obj = [PSCustomObject]@{
        Identity               = $Room.PrimarySmtpAddress
        DisplayName            = $Room.DisplayName
        MTREnabled             = if ($place) { $place.IsMicrosoftTeamsRooms } else { $null }
        OfficeLocation         = if ($place) { $place.Office } else { $null }
        Localities             = if ($place) { Join-IfAny $place.Localities } else { $null }
        AudioDeviceName        = if ($place) { $place.AudioDeviceName } else { $null }
        VideoDeviceName        = if ($place) { $place.VideoDeviceName } else { $null }
        DisplayDeviceName      = if ($place) { $place.DisplayDeviceName } else { $null }
        Building               = if ($place) { $place.Building } else { $null }
        Floor                  = if ($place) { $place.Floor } else { $null }
        FloorLabel             = if ($place) { $place.FloorLabel } else { $null }
        Street                 = if ($place) { $place.Street } else { $null }
        City                   = if ($place) { $place.City } else { $null }
        State                  = if ($place) { $place.State } else { $null }
        PostalCode             = if ($place) { $place.PostalCode } else { $null }
        CountryOrRegion        = $Room.CountryOrRegion
        UsageLocation          = $Room.UsageLocation
        Capacity               = if ($place) { $place.Capacity } else { $null }
        IsWheelChairAccessible = if ($place) { $place.IsWheelChairAccessible } else { $null }
        GeoCoordinates         = if ($place) { $place.GeoCoordinates } else { $null }
        Tags                   = if ($place) { Join-IfAny $place.Tags } else { $null }
        RoomPhone              = $Room.Phone
        PlacePhone             = if ($place) { $place.Phone } else { $null }
        ManagedBy              = Join-IfAny $Room.ManagedBy
        ResourceDelegates      = Join-IfAny $Room.ResourceDelegates

        # Error metadata columns
        GetPlaceSuccess        = [bool]$place
        GetPlaceError          = if ($place) { $null } else { $errorText }
    }
}

# ---------------------------------------
# Export to CSV
# ---------------------------------------
$Desktop    = Join-Path $env:USERPROFILE 'Desktop'
$ExportPath = Join-Path $Desktop 'RoomMailbox_Audit.csv'
$results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8

Write-Host "Audit complete. CSV saved to $ExportPath" -ForegroundColor Green