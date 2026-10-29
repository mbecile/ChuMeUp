<#
<. DESCRIPTION 
This script connects to Exchange Online, retrieves all room mailboxes, and gathers their associated Mailbox and Microsoft Places information, including Microsoft Teams Rooms (MTR) metadata. 
It also checks the Room List distribution group membership to confirm if the room is available in Room Finder. 
The results are exported to a CSV file on the user's desktop.
#>

Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline
# ---------------------------------------
# Configuration
# ---------------------------------------
$RetryCount        = 3      # Get-Place retry attempts
$RetryDelaySeconds = 2      # Delay between retries
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
    catch {
        return $Value
    }
}
# ---------------------------------------
# Helper: Get-Place with retry + error
# ---------------------------------------
function Get-PlaceWithRetry {
    param(
        [Parameter(Mandatory=$true)][string] $Identity,
        [int] $MaxRetries = 3,
        [int] $DelaySeconds = 2
    )
    $place        = $null
    $errorMessage = $null
    for ($i = 1; $i -le $MaxRetries -and -not $place; $i++) {
        try {
            $place = Get-Place -Identity $Identity -ErrorAction Stop
        }
        catch {
            $errorMessage = $_.Exception.Message
            if ($i -lt $MaxRetries) { Start-Sleep -Seconds $DelaySeconds }
        }
    }
    [PSCustomObject]@{
        Place = $place
        Error = if ($place) { $null } else { $errorMessage }
    }
}
# ---------------------------------------
# Collect all Room Mailboxes
# ---------------------------------------
Write-Host "Collecting all room mailboxes..." -ForegroundColor Cyan
$AllRooms = Get-Mailbox -RecipientTypeDetails RoomMailbox -ResultSize Unlimited
# ---------------------------------------
# Process Rooms - Include reattempts for Get-Place if it fails due to transient issues
# ---------------------------------------
$results = foreach ($Room in $AllRooms) {
    $placeResult = Get-PlaceWithRetry -Identity $Room.PrimarySmtpAddress -MaxRetries $RetryCount -DelaySeconds $RetryDelaySeconds
    $place       = $placeResult.Place
    $errorText   = $placeResult.Error

    if (-not $place) {
        Write-Warning "Get-Place failed for $($Room.PrimarySmtpAddress): $errorText"
    }
    
    [PSCustomObject]@{
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