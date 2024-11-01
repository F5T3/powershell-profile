# Ensure the script can run with elevated privileges
if ($Env:OS -eq "Windows_NT") {
    if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        Write-Warning "Please run this script as an Administrator!"
        break
    }
} else {
    Write-Host "Executing in MacOS or Linux mode" -ForegroundColor Green
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
            $profilePath = "$env:PROFILE"
        }
        elseif ($PSVersionTable.PSEdition -eq "Desktop") {
            $profilePath = "$env:PROFILE"
        }

        if (!(Test-Path -Path $profilePath)) {
            New-Item -Path $profilePath -ItemType "directory"
        }

        Invoke-RestMethod https://github.com/F5T3/powershell-profile/raw/main/Microsoft.PowerShell_profile.ps1 -OutFile $PROFILE
        Write-Host "The profile @ [$PROFILE] has been created." -ForegroundColor Green
        Write-Host "If you want to add any persistent components, please do so at [$profilePath\Profile.ps1] as there is an updater in the installed profile which uses the hash to update the profile and will lead to loss of changes" -ForegroundColor Yellow
    }
    catch {
        Write-Error "Failed to create or update the profile. Error: $_" -ForegroundColor red
    }
}
else {
    try {
        Get-Item -Path $PROFILE | Move-Item -Destination "oldprofile.ps1" -Force
        Invoke-RestMethod https://github.com/F5T3/powershell-profile/raw/main/Microsoft.PowerShell_profile.ps1 -OutFile $PROFILE
        Write-Host "The profile @ [$PROFILE] and has replace the old profile." -ForegroundColor Green
        Write-Host "Please back up any persistent components of your old profile to [$PROFILE\..\Profile.ps1] as there is an updater in the installed profile which uses the hash to update the profile and will lead to loss of changes" -ForegroundColor Yellow  
    }
    catch {
        Write-Error "Failed to backup and update the profile. Error: $_" -ForegroundColor Red
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
    Write-Error "Failed to download or install the Roboto Mono Nerd Font. Error: $_" -ForegroundColor Red
}

# Final check and message to the user
if ((Test-Path -Path $PROFILE) -and (winget list --name "OhMyPosh" -e) -and ($fontFamilies -contains "RobotoMono Nerd Font")) {
    Write-Host "Setup completed successfully. Please restart your PowerShell session to apply changes." -ForegroundColor Green
} else {
    Write-Warning "Setup completed with errors. Please check the error messages above."
}
#Install packages
try {
    $packages = @("Zoxide", "Starship", "Neovim", "Terminal-Icon", "Neofetch", "Everything", "EverythingToolbar", "Docker", "GlazeWM", "OhMyPosh", "Chocolatey", "Shell")
    $missingPackages = @()

    foreach ($package in $packages) {
        $packageInfo = winget list --id $package

        if (-not $packageInfo) {
            $missingPackages += $package
        }
    }

    if ($missingPackages.Count -eq 0) {
        Write-Host "The apps are already installed" -ForegroundColor Green
    } else {
        foreach ($packageName in $missingPackages) {
            winget install --id $packageName
        }
        $installed = ($missingPackages -join ", ") -replace ",([^,]+)$"," and`$1"
        Write-Host "The apps, $installed got installed" -ForegroundColor Green
    }
}
catch {
    Write-Error "Failed to download or install the apps. Error: $_" -ForegroundColor Red
}
