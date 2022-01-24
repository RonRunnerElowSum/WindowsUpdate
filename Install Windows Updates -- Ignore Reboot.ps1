$BlacklistedPatches = (Invoke-WebRequest -URI "https://raw.githubusercontent.com/RonRunnerElowSum/WindowsUpdate/Prod-Branch/BlackListedPatches.cfg" -UseBasicParsing).Content

function InstallPSWindowsUpdate () {
    Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope CurrentUser -Force -ErrorAction SilentlyContinue
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Install-PackageProvider -Name NuGet -Force | Out-Null
    Install-Module PSWindowsUpdate -Force | Out-Null
    Import-Module PSWindowsUpdate -Force | Out-Null
    if(!(Get-Module -Name "PSWindowsUpdate")){
        Write-Warning "The module PSWindowsUpdate failed to install...exiting..."
        Write-PatchLog "Warning: the module PSWindowsUpdate failed to install...exiting..."
        EXIT
    }
}

function InstallUpdates () {
    Write-Host "Checking for Windows Updates..."
    Write-PatchLog "Checking for Windows Updates..."
    $MissingUpdates = (Get-WindowsUpdate -MicrosoftUpdate -NotCategory Drivers -NotTitle "Feature update to Windows 10" -NotKBArticleID $BlacklistedPatches).KB
    if(!($Null -eq $MissingUpdates)){
        $NumberOfMissingPatches = $MissingUpdates.Count
        if($MissingUpdates.Count -eq "1"){
            $FormattedMissingUpdates = $MissingUpdates
        }
        else{
            $FormattedMissingUpdates = [string]::Join("`r`n",($MissingUpdates))
        }
        Write-Warning "$Env:COMPUTERNAME is missing the following ($NumberOfMissingPatches) patches:`r`n$FormattedMissingUpdates"
        Write-PatchLog "$Env:COMPUTERNAME is missing the following ($NumberOfMissingPatches) patches:`r`n$FormattedMissingUpdates"
        Write-Host "Installing missing updates..."
        Write-PatchLog "Installing missing updates..."
        $FormattedMissingUpdates | ForEach-Object {
            Install-WindowsUpdate -KBArticleID "$_" -IgnoreReboot -Confirm:$False
        }
        CheckPendingRebootStatus
    }
    else{
        Write-Host "Windows is up-to-date!"
        Write-PatchLog "Windows is up-to-date!"
        EXIT
    }
}

function Write-PatchLog ($PatchLogEntryValue) {
    $CurrentMonthYear = Get-Date -Format MMyyyy
    if(!(Test-Path -Path "C:\Windows\Temp")){New-Item -Path "C:\Windows" -Name "Temp" -ItemType "Directory" | Out-Null}
    if(!(Test-Path -Path "C:\Windows\Temp\MSP")){New-Item -Path "C:\Windows\Temp" -Name "MSP" -ItemType "Directory" | Out-Null}
    if(!(Test-Path -Path "C:\Windows\Temp\MSP\Logs")){New-Item -Path "C:\Windows\Temp\MSP" -Name "Logs" -ItemType "Directory" | Out-Null}
    if(!(Test-Path -Path "C:\Windows\Temp\MSP\Logs\Patch Health")){New-Item -Path "C:\Windows\Temp\MSP\Logs" -Name "Patch Health" -ItemType "Directory" | Out-Null}
    if(!(Test-Path -Path "C:\Windows\Temp\MSP\Logs\Patch Health\PatchHealthLog-$CurrentMonthYear.log")){New-Item -Path "C:\Windows\Temp\MSP\Logs\Patch Health" -Name "PatchHealthLog-$CurrentMonthYear.log" -ItemType "File" | Out-Null}
    Add-Content -Path "C:\Windows\Temp\MSP\Logs\Patch Health\PatchHealthLog-$CurrentMonthYear.log" -Value "$(Get-Date) -- $PatchLogEntryValue"
}

function CheckPendingRebootStatus () {
    $PendingRebootStatus = Get-WURebootStatus -Silent -CancelReboot
    Write-Host "Pending Reboot Status: $PendingRebootStatus"
    Write-PatchLog "Pending Reboot Status: $PendingRebootStatus"
    if($PendingRebootStatus -eq "True"){
        <#
        if(!(Get-ScheduledTask -TaskName '(MSP) Pending Reboot Checker' -ErrorAction SilentlyContinue)){
            Write-PatchLog "PRC is not installed...installing now..."
            Invoke-WebRequest -URI "https://raw.githubusercontent.com/RonRunnerElowSum/PendingRebootChecker/Prod-Branch/PRC%20Installer.ps1" -UseBasicParsing | Invoke-Expression; PunchIt | Out-Null
        }
        #>
        if(!(Get-ScheduledTask -TaskName '(MSP) Throw Reboot Required Toast Notification' -ErrorAction SilentlyContinue)){
            Write-PatchLog "The scheduled task ((MSP) Throw Reboot Required Toast Notification) does not exist...creating now..."
            Invoke-WebRequest -URI "https://raw.githubusercontent.com/RonRunnerElowSum/PendingRebootChecker/Prod-Branch/Create%20Reboot%20Required%20Toast%20Scheduled%20Task.ps1" -UseBasicParsing | Invoke-Expression; PunchIt | Out-Null
        }
        Write-Host "Throwing reboot required toast notification..."
        Write-PatchLog "Throwing reboot required toast notification..."
        Start-ScheduledTask -TaskName "(MSP) Throw Reboot Required Toast Notification"
    }
    else{
        Write-Host "$env:ComputerName is not currently in a pending reboot state..."
        Write-PatchLog "$env:ComputerName is not currently in a pending reboot state..."
    }
}

function PunchIt () {
    $ClientOS = Get-WmiObject -class Win32_OperatingSystem | Select-Object -ExpandProperty Caption
    if(($ClientOS | Select-String "Windows 7") -or ($ClientOS | Select-String "Server 2003") -or ($ClientOS | Select-String "2008")){
        Write-Host "OS: $ClientOS"
        Write-Warning "This operating system is no longer supported...exiting..."
        EXIT
    }
    if($ClientOS | Select-String "Server"){
        Write-Host "OS: $ClientOS"
        Write-Host "Install patches on servers between 11pm and 5am...exiting..."
        EXIT
    }
    $Win10CurrentBuildNumber = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -ErrorAction SilentlyContinue).CurrentBuildNumber
    if($Win10CurrentBuildNumber -eq "14393"){
        $Win10Build = "1607"
    }
    if($Win10CurrentBuildNumber -eq "15063"){
        $Win10Build = "1703"
    }
    if($Win10CurrentBuildNumber -eq "16299"){
        $Win10Build = "1709"
    }
    if($Win10CurrentBuildNumber -eq "17134"){
        $Win10Build = "1803"
    }
    if($Win10CurrentBuildNumber -eq "18363"){
        $Win10Build = "1909"
    }
    if($Win10CurrentBuildNumber -eq "19041"){
        $Win10Build = "2004"
    }
    if($Win10CurrentBuildNumber -eq "19042"){
        $Win10Build = "20H2"
    }
    if($Win10CurrentBuildNumber -eq "19043"){
        $Win10Build = "21H1"
    }
    if($Win10CurrentBuildNumber -eq "19044"){
        $Win10Build = "21H2"
    }
    Write-Host "Computer Name: $Env:COMPUTERNAME"
    Write-PatchLog "Computer Name: $Env:COMPUTERNAME"
    Write-Host "OS: $ClientOS $Win10Build"
    Write-PatchLog "OS: $ClientOS $Win10Build"
    if(!(Get-Module -Name "PSWindowsUpdate")){
        InstallPSWindowsUpdate
    }
    InstallUpdates
}
