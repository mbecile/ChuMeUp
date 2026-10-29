Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline

# Collect all Room Mailboxes
$AllRooms = Get-Mailbox -RecipientTypeDetails RoomMailbox -ResultSize Unlimited

$results = foreach ($Room in $AllRooms) {
    try {
        $place = Get-Place -Identity $Room.PrimarySmtpAddress
    }
    catch {
        Write-Warning "Get-Place failed for $($Room.PrimarySmtpAddress): $($_.Exception.Message)"
        $place = $null
    }

    [PSCustomObject]@{
        Identity               = $Room.PrimarySmtpAddress
        DisplayName            = $Room.DisplayName
        MTREnabled             = $place.IsMicrosoftTeamsRooms
        OfficeLocation         = $place.Office
        Localities             = ($place.Localities -join ", ")
        AudioDeviceName        = $place.AudioDeviceName
        VideoDeviceName        = $place.VideoDeviceName
        DisplayDeviceName      = $place.DisplayDeviceName
        Building               = $place.Building
        Floor                  = $place.Floor
        FloorLabel             = $place.FloorLabel
        Street                 = $place.Street
        City                   = $place.City
        State                  = $place.State
        PostalCode             = $place.PostalCode
        CountryOrRegion        = $Room.CountryOrRegion
        UsageLocation          = $Room.UsageLocation
        Capacity               = $place.Capacity
        IsWheelChairAccessible = $place.IsWheelChairAccessible
        GeoCoordinates         = $place.GeoCoordinates
        Tags                   = ($place.Tags -join ", ")
        RoomPhone              = $Room.Phone
        PlacePhone             = $place.Phone
        ManagedBy              = ($Room.ManagedBy -join ", ")
        ResourceDelegates      = ($Room.ResourceDelegates -join ", ")
    }
}

# Export to CSV
$Desktop    = Join-Path $env:USERPROFILE 'Desktop'
$ExportPath = Join-Path $Desktop 'RoomMailbox_Audit.csv'
$results | Export-Csv -Path $ExportPath -NoTypeInformation

Write-Host "Audit complete. CSV saved to $ExportPath" -ForegroundColor Green