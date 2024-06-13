Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
Install-Module PSWindowsUpdate

#set time zone to CST
Set-TimeZone "Central Standard Time"

$wholeserialNumber = (Get-WmiObject Win32_BIOS).SerialNumber
$serialNumber = $wholeserialNumber.Substring($wholeserialNumber.Length - 7)
$newName = "DESKTOP-$serialNumber"
Rename-Computer -NewName $newName -Force

$UserName = Read-Host "Enter the user's name"

# Define the folder containing software files
$softwareFolder = "C:\Temp"

# Function to install Windows updates
function Install-WindowsUpdates {
    try {
        # Import the PSWindowsUpdate module
        Import-Module PSWindowsUpdate

        # Get and install all updates
        Get-WUInstall -AcceptAll -IgnoreReboot -Install -Verbose
    } catch {
        Write-Host "Failed to install Windows updates. Error: $($_.Exception.Message)"
    }
}

# Function to set power settings
function Set-PowerSettings {
	# Get the active power scheme
	$currScheme = (powercfg /getactivescheme).split()[3]

	# Set the power settings for the active power scheme to never turn off the screen or go to sleep.
	powercfg /setacvalueindex $currScheme SUB_VIDEO VIDEOIDLE 0
	powercfg /setacvalueindex $currScheme SUB_SLEEP STANDBYIDLE 0
    	# Set hard drive to never turn off when on AC power
    	powercfg /setacvalueindex $currScheme SUB_DISK DISKIDLE 0
	
	# Make the changes effective immediately
	powercfg /setactive $currScheme
}


# Function to install software
function Install-Software {
    param (
        [Parameter(Mandatory=$true)]
        [string]$softwareFolder
    )

    # List of .exe and .msi files to install
    $msiFiles = Get-ChildItem -Path $softwareFolder -Filter *.msi
    $exeFiles = Get-ChildItem -Path $softwareFolder -Filter *.exe

    # Install .msi files silently
    foreach ($msiFile in $msiFiles) {
        try {
            Write-Host "Installing $($msiFile.Name)..."
            Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$($msiFile.FullName)`" /qn" -Wait
            Write-Host "Installation of $($msiFile.Name) complete."
        } catch {
            Write-Host "Failed to install $($msiFile.Name). Error: $($_.Exception.Message)"
        }
    }

    # Install .exe files
    foreach ($exeFile in $exeFiles) {
        try {
            Write-Host "Installing $($exeFile.Name)..."
            Start-Process -FilePath "$($exeFile.FullName)" -Wait
            Write-Host "Installation of $($exeFile.Name) complete."
        } catch {
            Write-Host "Failed to install $($exeFile.Name). Error: $($_.Exception.Message)"
        }
    }
}

# Function to change local user password and optionally activate the account
function Change-LocalUserPassword {
    param(
        [Parameter(Mandatory=$true)]
        [string]$username,
        [Parameter(Mandatory=$false)]
        [bool]$activate = $false
    )

    # Check if the username is 'Admin' or 'barcom'
    if ($username -eq 'Administrator') {
        # Change the password 
        net user $username Pass1
    } elseif ($username -eq 'barcom') {
        # Change the password
        net user $username Pass2
    }

    # If activate is true, activate the account
    if ($activate) {
        net user $username /active:yes
    }
}

# Function to create a local admin user
function Create-LocalAdminUser {
    param(
        [Parameter(Mandatory=$true)]
        [string]$username
    )

    # Create the user as an administrator
    net user $username "" /add /active:yes

    # Add the user to the 'Administrators' group
    net localgroup Administrators $username /add

    # Require the user to change their password at next logon
    net user $username /logonpasswordchg:yes
}

# Set power settings
Set-PowerSettings

# Call the function to change the password and activate the account
Change-LocalUserPassword -username "Administrator" -activate $true

# Call the function to change the password without activating the account
Change-LocalUserPassword -username "Company"

# Call the function to create a local admin user
Create-LocalAdminUser -username $UserName

# Install software from the specified folder
Install-Software -softwareFolder $softwareFolder

# Install Windows updates
Install-WindowsUpdates

Write-Host "All software installations, uninstallations, power plan settings, and Windows updates are complete."

# Prompt for domain name and join the computer to the domain (you will be prompted for credentials)
$domainName = Read-Host -Prompt 'Enter your domain name'
Add-Computer -DomainName $domainName -Credential (Get-Credential)
