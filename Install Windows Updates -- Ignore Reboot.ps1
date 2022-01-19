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
        InstallUpdates
    }  
}

function InstallUpdates () {
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
        Write-Host "Installing missing updates..."
        $FormattedMissingUpdates | ForEach-Object {
            Install-WindowsUpdate -KBArticleID "$_" -IgnoreReboot -Confirm:$False
        }
    }
    else{
        Write-Host "Windows is up-to-date!"
        EXIT
    }
}

function PunchIt () {
    $ClientOS = Get-WmiObject -class Win32_OperatingSystem | Select-Object -ExpandProperty Caption
    if(($ClientOS | Select-String "Windows 7") -or ($ClientOS | Select-String "Server 2003") -or ($ClientOS | Select-String "2008")){
        Write-Host "OS: $ClientOS"
        Write-Warning "This operating system is no longer supported...exiting..."
        EXIT
    }
    Write-Host "Computer Name: $Env:COMPUTERNAME"
    Write-Host "OS: $ClientOS"
    if(!(Get-Module -Name "PSWindowsUpdate")){
        InstallPSWindowsUpdate
    }
    else{
        InstallUpdates
    }
}
