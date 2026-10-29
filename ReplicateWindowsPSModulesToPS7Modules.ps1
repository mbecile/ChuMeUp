# -------------------------------
# Copy and install same modules from:
#   C:\Program Files\WindowsPowerShell\Modules
# to:
#   C:\Program Files\PowerShell\Modules
# -------------------------------

$OldPath  = "C:\Program Files\WindowsPowerShell\Modules"
$NewPath  = "C:\Program Files\PowerShell\Modules"

Write-Host "Scanning modules in $OldPath ..." -ForegroundColor Cyan

# Get all module names (folder names)
$Modules = Get-ChildItem -Path $OldPath -Directory | Select-Object -ExpandProperty Name
$total   = $Modules.Count

Write-Host "Found $total modules. Beginning..." -ForegroundColor Cyan

# Ensure PSGallery trusted
Set-PSRepository -Name PSGallery -InstallationPolicy Trusted

# Loop through each module with progress bar
for ($i = 0; $i -lt $total; $i++) {

    $Module = $Modules[$i]

    # Write progress
    $percent = ($i / $total * 100)
    Write-Progress -Activity "Migrating PowerShell Modules" `
                   -Status "Processing module: $Module ($percent% complete)" `
                   -PercentComplete $percent

    Write-Host "`nProcessing module: $Module" -ForegroundColor Yellow
    
    try {
        # Check if installed already in PowerShell 7 location
        $existing = Get-InstalledModule -Name $Module -ErrorAction SilentlyContinue

        if ($existing) {
            Write-Host "Already installed in PowerShell 7. Updating to latest version..." -ForegroundColor Green
            Update-Module -Name $Module -ErrorAction SilentlyContinue
        }
        else {
            Write-Host "Installing latest version of $Module into PowerShell 7 AllUsers location..." -ForegroundColor Green
            Install-Module -Name $Module -Scope AllUsers -Force -ErrorAction Stop
        }
    }
    catch {
        Write-Warning "❌ Failed to install/update module $Module : $($_.Exception.Message)"
    }
}

# Complete progress bar
Write-Progress -Activity "Copy PowerShell Modules to PowerShell 7" -Completed

Write-Host "`n✔ Migration complete. All modules processed." -ForegroundColor Cyan