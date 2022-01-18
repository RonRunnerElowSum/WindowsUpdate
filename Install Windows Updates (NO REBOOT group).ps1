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
}

function InstallUpdatesWithNoReboot () {
    $BlacklistedPatches = (Invoke-WebRequest -URI "https://raw.githubusercontent.com/RonRunnerElowSum/WindowsUpdate/Prod-Branch/BlackListedPatches.cfg" -UseBasicParsing).Content
    Write-Host "Checking for Windows Updates..."
    $DetailedMissingUpdates = (Get-WindowsUpdate -MicrosoftUpdate -NotCategory Drivers -NotTitle "Feature update to Windows 10" -NotKBArticleID $BlacklistedPatches)
    $MissingUpdates = ($DetailedMissingUpdates).KB
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
        CheckPendingRebootStatus
    }
    else{
        Write-Host "Windows is up-to-date!"
        EXIT
    }
}

function CheckPendingRebootStatus () {
    $PendingRebootStatus = Get-WURebootStatus -Silent -CancelReboot
    if($PendingRebootStatus -eq "True"){
        if(!(Get-ScheduledTask -TaskName "(MSP) Pending Reboot Checker" -ErrorAction SilentlyContinue)){
            Invoke-WebRequest -URI "https://raw.githubusercontent.com/RonRunnerElowSum/PendingRebootChecker/Prod-Branch/PRC%20Installer.ps1" -UseBasicParsing | Invoke-Expression; PunchIt | Out-Null
        }
        Start-ScheduledTask -TaskName "(MSP) Pending Reboot Checker"
    }
}

function PunchIt () {
    $ClientOS = Get-WmiObject -class Win32_OperatingSystem | Select-Object -ExpandProperty Caption
    if(($ClientOS | Select-String "Windows 7") -or ($ClientOS | Select-String "Server 2003") -or ($ClientOS | Select-String "2008")){
        Write-Host "OS: $ClientOS"
        Write-Warning "This operating system is no longer supported...exiting..."
        EXIT
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
    
    if($PatchGroupProduction | Select-String $CurrentDate){
        Write-Host "Computer Name: $Env:COMPUTERNAME"
        Write-Host "OS: $ClientOS"
        if(!(Get-Module -Name "PSWindowsUpdate")){
            InstallPSWindowsUpdate
        }
        InstallUpdatesWithNoReboot
    }
    else{
        Write-Host "Outside of patch window...$Env:COMPUTERNAME patches on the following days this patch cycle:`r`n`r`n$PatchGroupProduction"
        EXIT
    }
}
