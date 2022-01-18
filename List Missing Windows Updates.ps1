$BlacklistedPatches = (Invoke-WebRequest -URI "https://raw.githubusercontent.com/RonRunnerElowSum/WindowsUpdate/Prod-Branch/BlackListedPatches.cfg" -UseBasicParsing).Content

function InstallPSWindowsUpdate () {
    Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope CurrentUser -Force -ErrorAction SilentlyContinue
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Install-PackageProvider -Name NuGet -Force | Out-Null
    Install-Module PSWindowsUpdate -Force | Out-Null
    Import-Module PSWindowsUpdate -Force | Out-Null
    if(!(Get-Module -Name "PSWindowsUpdate")){
        Write-Warning "The module PSWindowsUpdate failed to install...exiting..."
        EXIT
    }
    else{
        ListUpdates
    }  
}

function ListUpdates () {
    Write-Host "Checking for Windows Updates..."
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
        EXIT
    }
    else{
        Write-Host "Windows is up-to-date!"
        EXIT
    }
}

function PunchIt () {
    $ClientOS = (Get-WmiObject -class Win32_OperatingSystem).Caption
    if(($ClientOS | Select-String "Windows 7") -or ($ClientOS | Select-String "Server 2003") -or ($ClientOS | Select-String "2008")){
        Write-Host "OS: $ClientOS"
        Write-Warning "This operating system is no longer supported...exiting..."
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
    if(!(Get-Module -Name "PSWindowsUpdate")){
        InstallPSWindowsUpdate
    }
    else{
        ListUpdates
    }
}