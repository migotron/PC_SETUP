# Start logging
try {
    Start-Transcript -Path C:\temp\OfficeUninstall.log -Append
    Write-Output "This script is intended for uninstalling Office installed from the Microsoft Store."
} catch {
    Write-Output "Error starting transcript: $_"
    return
}

# Get all apps installed from the Microsoft Store
$allApps = Get-AppxPackage

# Filter out Office 365 apps
$officeApps = $allApps | Where-Object { $_.Name -like "Microsoft.*Office*" -or $_.Name -like "Microsoft.*365*" -or $_.Name -like "Microsoft.MicrosoftOfficeHub" }

if ($null -eq $officeApps) {
    Write-Output "Office is not installed or not installed from the Microsoft Store. Skipping uninstallation..."
    Stop-Transcript
    return
}

# Display a numbered list of Office apps and let the user select which one to uninstall
Write-Output "Select the Office app to uninstall:"
$i = 0
$officeApps | ForEach-Object {
    $i++
    Write-Output "${i}: $($_.Name)"
}

$officeAppIndex = Read-Host -Prompt 'Enter the number of the Office app you want to uninstall'
$officeAppToUninstall = $officeApps[$officeAppIndex - 1].Name

# Uninstall selected Office app
Write-Output "Uninstalling ${officeAppToUninstall}..."
try {
    $officeApps | Where-Object { $_.Name -eq $officeAppToUninstall } | Remove-AppxPackage
} catch {
    Write-Output "Error uninstalling ${officeAppToUninstall}: ${_}"
    Stop-Transcript
    return
}

# Wait for the uninstallation process to complete
Write-Output "Waiting for ${officeAppToUninstall} to uninstall..."
try {
    while ($null -ne (Get-AppxPackage -name $officeAppToUninstall)) {
        # Wait for 10 seconds before checking again
        Start-Sleep -Seconds 10
    }
} catch {
    Write-Output "Error waiting for ${officeAppToUninstall} to uninstall: ${_}"
    Stop-Transcript
    return
}

# Verify Office app was removed
Write-Output "Verifying ${officeAppToUninstall} was removed..."
if ($null -eq (Get-AppxPackage -name $officeAppToUninstall)) {
    Write-Output "${officeAppToUninstall} was successfully removed."
} else {
    Write-Output "${officeAppToUninstall} was not successfully removed."
}

# Close the PowerShell window
Write-Output "Closing PowerShell..."

# Stop logging
try {
    Stop-Transcript
} catch {
    Write-Output "Error stopping transcript: ${_}"
}
