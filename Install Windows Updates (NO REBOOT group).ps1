$BlacklistedPatches = (Invoke-WebRequest -URI "https://raw.githubusercontent.com/RonRunnerElowSum/WindowsUpdate/Prod-Branch/BlackListedPatches.cfg" -UseBasicParsing).Content

$ClientOS = Get-WmiObject -class Win32_OperatingSystem | Select-Object -ExpandProperty Caption
if(($ClientOS | Select-String "Windows 7") -or ($ClientOS | Select-String "Server 2003") -or ($ClientOS | Select-String "2008")){
    Write-Host "OS: $ClientOS"
    Write-Warning "This operating system is no longer supported...exiting..."
    EXIT
}

function Get-ThisMonthsPatchTuesday {
    [CmdletBinding()]
    Param
    (
      [Parameter(position = 0)]
      [ValidateSet("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")]
      [String]$weekDay = 'Tuesday',
      [ValidateRange(0, 5)]
      [Parameter(position = 1)]
      [int]$findNthDay = 2
    )
    [datetime]$today = Get-Date -Format d
    $todayM = $today.Month.ToString()
    $todayY = $today.Year.ToString()
    [datetime]$strtMonth = $todayM + '/1/' + $todayY
    while ($strtMonth.DayofWeek -ine $weekDay ) { $strtMonth = $StrtMonth.AddDays(1) }
    $firstWeekDay = $strtMonth
    if ($findNthDay -eq 1) {
      $dayOffset = 0
    }
    else {
      $dayOffset = ($findNthDay - 1) * 7
    }
    $patchTuesday = $firstWeekDay.AddDays($dayOffset)
    return $patchTuesday
}
function Get-LastMonthsPatchTuesday {
    [CmdletBinding()]
    Param
    (
        [Parameter(position = 0)]
        [ValidateSet("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")]
        [String]$WeekDay = 'Tuesday',
        [ValidateRange(0, 5)]
        [Parameter(position = 1)]
        [int]$FindNthDay = 2
    )
    [datetime]$Today = Get-Date -Format d
    $ThisMonth = $Today.Month
    $ThisYear = $Today.Year
    if($ThisMonth -eq "1"){
        $ThisMonth = "12"
        $DecreaseValue = 1
        $Thisyear = $ThisYear - $DecreaseValue
    }
    else{
        $DecreaseValue = 1
        $ThisMonth = $ThisMonth - $DecreaseValue
    }
    [datetime]$DateString = $ThisMonth + '/1/' + $ThisYear
    while ($DateString.DayofWeek -ine $WeekDay ) { $DateString = $DateString.AddDays(1) }
    $firstWeekDay = $DateString
    if($FindNthDay -eq 1){
        $DayOffSet = 0
    }
    else{
        $DayOffSet = ($FindNthDay - 1) * 7
    }
    $LastMonthsPatchTuesday = $FirstWeekDay.AddDays($DayOffSet)
    return $LastMonthsPatchTuesday
}

if((Get-Date) -lt (Get-ThisMonthsPatchTuesday)){
    $PatchTuesday = Get-LastMonthsPatchTuesday
}
else{
    $PatchTuesday = Get-ThisMonthsPatchTuesday
}

$PatchGroupProduction = @()

$PatchGroupProduction += ($PatchTuesday).AddDays(12).ToShortDateString()
$PatchGroupProduction += ($PatchTuesday).AddDays(13).ToShortDateString()
$PatchGroupProduction += ($PatchTuesday).AddDays(14).ToShortDateString()
$PatchGroupProduction += ($PatchTuesday).AddDays(15).ToShortDateString()
$PatchGroupProduction += ($PatchTuesday).AddDays(16).ToShortDateString()
$PatchGroupProduction += ($PatchTuesday).AddDays(17).ToShortDateString()
$PatchGroupProduction += ($PatchTuesday).AddDays(18).ToShortDateString()
$PatchGroupProduction += ($PatchTuesday).AddDays(19).ToShortDateString()
$PatchGroupProduction += ($PatchTuesday).AddDays(20).ToShortDateString()
$PatchGroupProduction += ($PatchTuesday).AddDays(21).ToShortDateString()
$PatchGroupProduction += ($PatchTuesday).AddDays(22).ToShortDateString()
$PatchGroupProduction += ($PatchTuesday).AddDays(23).ToShortDateString()
$PatchGroupProduction += ($PatchTuesday).AddDays(24).ToShortDateString()
$PatchGroupProduction += ($PatchTuesday).AddDays(25).ToShortDateString()
$PatchGroupProduction += ($PatchTuesday).AddDays(26).ToShortDateString()

$PatchGroupPilot = @()

$PatchGroupPilot += ($PatchTuesday).AddDays(7).ToShortDateString()
$PatchGroupPilot += ($PatchTuesday).AddDays(8).ToShortDateString()
$PatchGroupPilot += ($PatchTuesday).AddDays(9).ToShortDateString()
$PatchGroupPilot += ($PatchTuesday).AddDays(10).ToShortDateString()
$PatchGroupPilot += ($PatchTuesday).AddDays(11).ToShortDateString()
$PatchGroupPilot += $PatchGroupProduction

$PatchGroupTest = @()
$PatchGroupTest += ($PatchTuesday).ToShortDateString()
$PatchGroupTest += ($PatchTuesday).AddDays(1).ToShortDateString()
$PatchGroupTest += ($PatchTuesday).AddDays(2).ToShortDateString()
$PatchGroupTest += ($PatchTuesday).AddDays(3).ToShortDateString()
$PatchGroupTest += ($PatchTuesday).AddDays(4).ToShortDateString()
$PatchGroupTest += ($PatchTuesday).AddDays(5).ToShortDateString()
$PatchGroupTest += ($PatchTuesday).AddDays(6).ToShortDateString()
$PatchGroupTest += $PatchGroupPilot
$PatchGroupTest += $PatchGroupProductions

$CurrentDate = Get-Date -Format d

if(($CurrentDate -eq $PatchGroupProduction) -and ((Get-Date).TimeOfDay.TotalHours -lt "5")){
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
        $DetailedMissingUpdates = (Get-WindowsUpdate -MicrosoftUpdate -NotCategory Drivers -NotTitle "Feature update to Windows 10" -NotKBArticleID $BlacklistedPatches)
        $MissingUpdates = $DetailedMissingUpdates | Select-Object -ExpandProperty "KB" -ErrorAction SilentlyContinue
        if(!($Null -eq $MissingUpdates)){
            if($MissingUpdates.Count -eq "1"){
                $FormattedMissingUpdates = $MissingUpdates
            }
            else{
                $FormattedMissingUpdates = [string]::Join("`r`n",($MissingUpdates))
            }
            Write-Warning "$Env:COMPUTERNAME is missing the following patches:`r`n$FormattedMissingUpdates"
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
        InstallUpdates
    }  
}
    else{
        Write-Host "Outside of patch window...$Env:COMPUTERNAME patches from 12am to 5am on the following days this patch cycle:`r`n`r`n$PatchGroupProduction"
        EXIT
    }

