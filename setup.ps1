# Ensure the script can run with elevated privileges
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "Please run this script as an Administrator!"
    break
}

# Function to test internet connectivity
function Test-InternetConnection {
    try {
        $testConnection = Test-Connection -ComputerName www.google.com -Count 1 -ErrorAction Stop
        return $true
    }
    catch {
        Write-Warning "Internet connection is required but not available. Please check your connection."
        return $false
    }
}

# Check for internet connectivity before proceeding
if (-not (Test-InternetConnection)) {
    break
}

# Profile creation or update
if (!(Test-Path -Path $PROFILE -PathType Leaf)) {
    try {
        # Detect Version of PowerShell & Create Profile directories if they do not exist.
        $profilePath = ""
        if ($PSVersionTable.PSEdition -eq "Core") { 
            $profilePath = "$env:userprofile\Documents\Powershell"
        }
        elseif ($PSVersionTable.PSEdition -eq "Desktop") {
            $profilePath = "$env:userprofile\Documents\WindowsPowerShell"
        }

        if (!(Test-Path -Path $profilePath)) {
            New-Item -Path $profilePath -ItemType "directory"
        }

        Invoke-RestMethod https://github.com/F5T3/powershell-profile/raw/main/Microsoft.PowerShell_profile.ps1 -OutFile $PROFILE
        Write-Host "The profile @ [$PROFILE] has been created."
        Write-Host "If you want to add any persistent components, please do so at [$profilePath\Profile.ps1] as there is an updater in the installed profile which uses the hash to update the profile and will lead to loss of changes"
    }
    catch {
        Write-Error "Failed to create or update the profile. Error: $_"
    }
}
else {
    try {
        Get-Item -Path $PROFILE | Move-Item -Destination "oldprofile.ps1" -Force
        Invoke-RestMethod https://github.com/F5T3/powershell-profile/raw/main/Microsoft.PowerShell_profile.ps1 -OutFile $PROFILE
        Write-Host "The profile @ [$PROFILE] has been created and old profile removed."
        Write-Host "Please back up any persistent components of your old profile to [$HOME\Documents\PowerShell\Profile.ps1] as there is an updater in the installed profile which uses the hash to update the profile and will lead to loss of changes"
    }
    catch {
        Write-Error "Failed to backup and update the profile. Error: $_"
    }
}

# Font Install
try {
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
    $fontFamilies = (New-Object System.Drawing.Text.InstalledFontCollection).Families.Name

    if ($fontFamilies -notcontains "RobotoMono Nerd Font") {
        $webClient = New-Object System.Net.WebClient
        $webClient.DownloadFileAsync((New-Object System.Uri("https://github.com/ryanoasis/nerd-fonts/releases/download/v3.2.1/RobotoMono.zip")), ".\RobotoMono.zip")
        
        while ($webClient.IsBusy) {
            Start-Sleep -Seconds 2
        }

        Expand-Archive -Path ".\RobotoMono.zip" -DestinationPath ".\RobotoMono" -Force
        $destination = (New-Object -ComObject Shell.Application).Namespace(0x14)
        Get-ChildItem -Path ".\RobotoMono" -Recurse -Filter "*.ttf" | ForEach-Object {
            If (-not(Test-Path "C:\Windows\Fonts\$($_.Name)")) {        
                $destination.CopyHere($_.FullName, 0x10)
            }
        }

        Remove-Item -Path ".\RobotoMono" -Recurse -Force
        Remove-Item -Path ".\RobotoMono.zip" -Force
    }
}
catch {
    Write-Error "Failed to download or install the Roboto Mono Nerd Font. Error: $_"
}

# Final check and message to the user
if ((Test-Path -Path $PROFILE) -and (winget list --name "OhMyPosh" -e) -and ($fontFamilies -contains "RobotoMono Nerd Font")) {
    Write-Host "Setup completed successfully. Please restart your PowerShell session to apply changes."
} else {
    Write-Warning "Setup completed with errors. Please check the error messages above."
}
function InstallPackages {
    $packages = @("Zoxide", "Starship", "Neovim", "Terminal-Icon", "Neofetch", "Everything", "EverythingToolbar", "Docker", "GlazeWM", "Oh My Posh", "Chocolatey")
    $installedPackages = @()
    $missingPackages = @()

    foreach ($package in $packages) {
        if (-not (Get-Command $package -ErrorAction SilentlyContinue)) {
            $missingPackages += $package
        } else {
            $installedPackages += $package
        }
    }

    if ($missingPackages.Count -eq 0) {
        Write-Output "All apps are already installed"
    } else {
        foreach ($package in $missingPackages) {
            # Attempt to install the missing package using the appropriate package manager
            if ($package -eq "Zoxide") {
                Install-Package -Name zoxide -Source winget -Force
            } elseif ($package -eq "Starship") {
                winget install starship
            } elseif ($package -eq "Neovim") {
                winget install neovim
            } elseif ($package -eq "Terminal-Icon") {
                Install-Module -Name Terminal-Icons -Repository PSGallery -Force
            } elseif ($package -eq "Neofetch") {
                winget install neofetch 
            } elseif ($package -eq "Everything") {
                winget install --id voidtools.Everything
            } elseif ($package -eq "EverythingToolbar") {
                winget install --id stnkl.EverythingToolbar
            } elseif ($package -eq "Docker") {
                winget install docker
            } elseif ($package -eq "GlazeWM") {
                winget install GlazeWM
            } elseif ($package -eq "Oh My Posh") {
                winget install -e --accept-source-agreements --accept-package-agreements OhMyPosh
            } elseif ($package -eq "Chocolatey") {
                Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
            } 
        }
        $installed = ($missingPackages -join ", ") -replace ",([^,]+)$"," and`$1"
        Write-Output "The apps, $installed got installed"
    }
}
InstallPackages
