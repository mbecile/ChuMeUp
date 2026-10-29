# Get all dynamic membership groups
$dynGroups = Get-MgGroup -All -Property "id,displayName,groupTypes,membershipRule" |
    Where-Object { $_.GroupTypes -contains "DynamicMembership" }

Write-Host "`nFound $($dynGroups.Count) dynamic groups.`n"

foreach ($g in $dynGroups) {

    $rule = $g.membershipRule

    if (-not $rule) {
        Write-Warning "Group '$($g.DisplayName)' has no membershipRule."
        continue
    }

    # Detect hidden or non‑ASCII characters
    $badChars = $rule.ToCharArray() | Where-Object { [int]$_ -gt 127 }

    # Detect known invisible Unicode problem characters
    $invisiblePattern = '[\u200B-\u200F\uFEFF\u00A0]'
    $invisibleMatches = ($rule | Select-String -Pattern $invisiblePattern -AllMatches).Matches

    if ($badChars -or $invisibleMatches.Count -gt 0) {
        Write-Warning "BAD CHARACTERS DETECTED in group '$($g.DisplayName)' (ID: $($g.Id))"

        if ($badChars) {
            Write-Host " Non-ASCII characters:" -ForegroundColor Yellow
            $badChars | ForEach-Object {
                '{0} (U+{1:X4})' -f $_, [int]$_
            }
        }

        if ($invisibleMatches.Count -gt 0) {
            Write-Host " Invisible characters found (zero-width, BOM, or NBSP)." -ForegroundColor Yellow
        }

        Write-Host " Rule:" $rule "`n"
    }
}
