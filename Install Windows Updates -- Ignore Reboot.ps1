[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$BlacklistedPatches = @((Invoke-WebRequest -URI "https://raw.githubusercontent.com/RonRunnerElowSum/WindowsUpdate/Prod-Branch/BlackListedPatches.cfg" -UseBasicParsing).Content)

function InstallPSWindowsUpdate () {
    Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope CurrentUser -Force -ErrorAction SilentlyContinue
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Install-PackageProvider -Name NuGet -Force | Out-Null
    Install-Module PSWindowsUpdate -Force | Out-Null
    Import-Module PSWindowsUpdate -Force | Out-Null
    if(!(Get-Module -Name "PSWindowsUpdate")){
        Write-Warning "The module PSWindowsUpdate failed to install...exiting..."
        Write-MSPLog -LogSource "MSP Patch Health" -LogType "Warning" -LogMessage "Warning: the module PSWindowsUpdate failed to install...exiting..."
        EXIT
    }
}

function InstallUpdates () {
    Write-Host "Checking for Windows Updates..."
    Write-MSPLog -LogSource "MSP Patch Health" -LogType "Information" -LogMessage "Checking for Windows Updates..."
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
        Write-MSPLog -LogSource "MSP Patch Health" -LogType "Warning" -LogMessage "$Env:COMPUTERNAME is missing the following ($NumberOfMissingPatches) patches:`r`n$FormattedMissingUpdates"
        Write-Host "Installing missing updates..."
        Write-MSPLog -LogSource "MSP Patch Health" -LogType "Information" -LogMessage "Installing missing updates..."
        $FormattedMissingUpdates | ForEach-Object {
            [string]$InstallUpdates = Install-WindowsUpdate -KBArticleID "$_" -IgnoreReboot -Confirm:$False | Out-String
            Write-MSPLog -LogSource "MSP Patch Health" -LogType "Information" -LogMessage $InstallUpdates
        }
        CheckPendingRebootStatus
    }
    else{
        Write-Host "Windows is up-to-date!"
        Write-MSPLog -LogSource "MSP Patch Health" -LogType "Information" -LogMessage "Windows is up-to-date!"
        EXIT
    }
}

function Write-MSPLog {
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [ValidateSet('MSP Patch Health','MSP PendingRebootChecker','MSP Disk Health')]
         [string] $LogSource,
         [Parameter(Mandatory=$true, Position=1)]
         [ValidateSet('Information','Warning','Error')]
         [string] $LogType,
         [Parameter(Mandatory=$true, Position=2)]
         [string] $LogMessage
    )

    New-EventLog -LogName MSP-IT -Source 'MSP' -ErrorAction SilentlyContinue
    if(!(Get-EventLog -LogName MSP-IT -Source 'MSP Patch Health' -ErrorAction SilentlyContinue)){
        New-EventLog -LogName MSP-IT -Source 'MSP Patch Health' -ErrorAction SilentlyContinue
    }
    if(!(Get-EventLog -LogName MSP-IT -Source 'MSP PendingRebootChecker' -ErrorAction SilentlyContinue)){
        New-EventLog -LogName MSP-IT -Source 'MSP PendingRebootChecker' -ErrorAction SilentlyContinue
    }
    if(!(Get-EventLog -LogName MSP-IT -Source 'MSP Disk Health' -ErrorAction SilentlyContinue)){
        New-EventLog -LogName MSP-IT -Source 'MSP Disk Health' -ErrorAction SilentlyContinue
    }
    Write-EventLog -Log MSP-IT -Source $LogSource -EventID 0 -EntryType $LogType -Message "$LogMessage"
}

function CheckPendingRebootStatus () {
    $PendingRebootStatus = Get-WURebootStatus -Silent -CancelReboot
    Write-MSPLog -LogSource "MSP PendingRebootChecker" -LogType "Information" -LogMessage "Pending Reboot Status: $PendingRebootStatus"
    if($PendingRebootStatus -eq "True"){
        <#
        if(!(Get-ScheduledTask -TaskName '(MSP) Pending Reboot Checker' -ErrorAction SilentlyContinue)){
            Write-MSPLog -LogSource "MSP PendingRebootChecker" -LogType "Information" -LogMessage "PRC is not installed...installing now..."
            Invoke-WebRequest -URI "https://raw.githubusercontent.com/RonRunnerElowSum/PendingRebootChecker/Prod-Branch/PRC%20Installer.ps1" -UseBasicParsing | Invoke-Expression; PunchIt | Out-Null
        }
        #>
        if(!(Get-ScheduledTask -TaskName '(MSP) Throw Reboot Required Toast Notification' -ErrorAction SilentlyContinue)){
            Write-MSPLog -LogSource "MSP PendingRebootChecker" -LogType "Information" -LogMessage "The scheduled task ((MSP) Throw Reboot Required Toast Notification) does not exist...creating now..."
            Invoke-WebRequest -URI "https://raw.githubusercontent.com/RonRunnerElowSum/PendingRebootChecker/Prod-Branch/Create%20Reboot%20Required%20Toast%20Scheduled%20Task.ps1" -UseBasicParsing | Invoke-Expression; PunchIt | Out-Null
        }
        Write-MSPLog -LogSource "MSP PendingRebootChecker" -LogType "Information" -LogMessage "Throwing reboot required toast notification..."
        Start-ScheduledTask -TaskName "(MSP) Throw Reboot Required Toast Notification"
    }
}

function PunchIt () {
    $ClientOS = Get-WmiObject -class Win32_OperatingSystem | Select-Object -ExpandProperty Caption
    if(($ClientOS | Select-String "Windows 7") -or ($ClientOS | Select-String "Server 2003") -or ($ClientOS | Select-String "2008")){
        Write-Host "OS: $ClientOS"
        Write-Warning "This operating system is no longer supported...exiting..."
        Write-MSPLog -LogSource "MSP Patch Health" -LogType "Warning" -LogMessage "OS: $ClientOS`r`n`This operating system is no longer supported...exiting..."
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
    Write-Host "OS: $ClientOS $Win10Build"
    Write-MSPLog -LogSource "MSP Patch Health" -LogType "Information" -LogMessage "Computer Name: $Env:COMPUTERNAME`r`nOS: $ClientOS $Win10Build"
    if(!(Get-Module -Name "PSWindowsUpdate")){
        InstallPSWindowsUpdate
    }
    InstallUpdates
}
