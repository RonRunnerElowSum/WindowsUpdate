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
        Write-PatchLog "The module PSWindowsUpdate failed to install...exiting..."
        EXIT
    }
}

function InstallUpdatesWithNoReboot () {
    $BlacklistedPatches = (Invoke-WebRequest -URI "https://raw.githubusercontent.com/RonRunnerElowSum/WindowsUpdate/Prod-Branch/BlackListedPatches.cfg" -UseBasicParsing).Content
    Write-PatchLog "Checking for Windows Updates..."
    $DetailedMissingUpdates = (Get-WindowsUpdate -MicrosoftUpdate -NotCategory Drivers -NotTitle "Feature update to Windows 10" -NotKBArticleID $BlacklistedPatches)
    $MissingUpdates = ($DetailedMissingUpdates).KB
    if(!($Null -eq $MissingUpdates)){
        if($MissingUpdates.Count -eq "1"){
            $FormattedMissingUpdates = $MissingUpdates
        }
        else{
            $FormattedMissingUpdates = [string]::Join("`r`n",($MissingUpdates))
        }
        Write-PatchLog "$Env:COMPUTERNAME is missing the following ($($MissingUpdates.Count)) patches:`r`n$FormattedMissingUpdates"
        Write-PatchLog "Installing missing updates..."
        $FormattedMissingUpdates | ForEach-Object {
            $CurrentMonthYear = Get-Date -Format MMyyyy
            Install-WindowsUpdate -KBArticleID "$_" -IgnoreReboot -Confirm:$False | Out-File "C:\Windows\Temp\MSP\Logs\Patch Health\PatchHealthLog-$CurrentMonthYear.log" -Append
        }
        CheckPendingRebootStatus
    }
    else{
        Write-PatchLog "Windows is up-to-date!"
    }
}

function InstallUpdatesWithForcedReboot () {
    $BlacklistedPatches = (Invoke-WebRequest -URI "https://raw.githubusercontent.com/RonRunnerElowSum/WindowsUpdate/Prod-Branch/BlackListedPatches.cfg" -UseBasicParsing).Content
    Write-PatchLog "Checking for Windows Updates..."
    $DetailedMissingUpdates = (Get-WindowsUpdate -MicrosoftUpdate -NotCategory Drivers -NotTitle "Feature update to Windows 10" -NotKBArticleID $BlacklistedPatches)
    $MissingUpdates = ($DetailedMissingUpdates).KB
    if(!($Null -eq $MissingUpdates)){
        if($MissingUpdates.Count -eq "1"){
            $FormattedMissingUpdates = $MissingUpdates
        }
        else{
            $FormattedMissingUpdates = [string]::Join("`r`n",($MissingUpdates))
        }
        Write-PatchLog "$Env:COMPUTERNAME is missing the following ($($MissingUpdates.Count)) patches:`r`n$FormattedMissingUpdates"
        Write-PatchLog "Installing missing updates..."
        $FormattedMissingUpdates | ForEach-Object {
            $CurrentMonthYear = Get-Date -Format MMyyyy
            Install-WindowsUpdate -KBArticleID "$_" -AutoReboot -Confirm:$False | Out-File "C:\Windows\Temp\MSP\Logs\Patch Health\PatchHealthLog-$CurrentMonthYear.log" -Append
        }
    }
    else{
        Write-PatchLog "Windows is up-to-date!"
        EXIT
    }
}

function CheckPendingRebootStatus () {
    $PendingRebootStatus = Get-WURebootStatus -Silent -CancelReboot
    Write-PatchLog "Pending Reboot Status: $PendingRebootStatus"
    if($PendingRebootStatus -eq "True"){
        if(!(Get-ScheduledTask -TaskName '(MSP) Pending Reboot Checker' -ErrorAction SilentlyContinue)){
            Write-PatchLog "PRC is not installed...installing now..."
            Invoke-WebRequest -URI "https://raw.githubusercontent.com/RonRunnerElowSum/PendingRebootChecker/Prod-Branch/PRC%20Installer.ps1" -UseBasicParsing | Invoke-Expression; PunchIt | Out-Null
        }
        [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
        [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") | Out-Null
        $ToastNotification = New-Object System.Windows.Forms.NotifyIcon
        $ToastNotification.Icon = [System.Drawing.SystemIcons]::Information
        $ToastNotification.BalloonTipText = "Your computer needs to restart in order to finishing installing updates. Please restart at your earliest convenience."
        $ToastNotification.BalloonTipTitle = "Reboot Required"
        $ToastNotification.BalloonTipIcon = "Warning"
        $ToastNotification.Visible = $True
        $ToastNotification.ShowBalloonTip(50000)
        $ToastNotification_MouseOver = [System.Windows.Forms.MouseEventHandler]{$ToastNotification.ShowBalloonTip(50000)}
        $ToastNotification.add_MouseClick($ToastNotification_MouseOver)
        Unregister-Event -SourceIdentifier click_event -ErrorAction SilentlyContinue
		Register-ObjectEvent $ToastNotification BalloonTipClicked -SourceIdentifier click_event -Action {
            Write-PatchLog "Executing PRC..."
            Start-ScheduledTask -TaskName '(MSP) Pending Reboot Checker'
        } | Out-Null
        Wait-Event -Timeout 10 -SourceIdentifier click_event > $null
        Unregister-Event -SourceIdentifier click_event -ErrorAction SilentlyContinue
        $ToastNotification.Dispose()
    }
    else{
        Write-PatchLog "$env:ComputerName is not currently in a pending reboot state..."
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

function PunchIt () {
    Write-PatchLog "Starting..."
    $ClientOS = Get-WmiObject -class Win32_OperatingSystem | Select-Object -ExpandProperty Caption
    if(($ClientOS | Select-String "Windows 7") -or ($ClientOS | Select-String "Server 2003") -or ($ClientOS | Select-String "2008")){
        Write-PatchLog "OS: $ClientOS"
        Write-PatchLog "This operating system is no longer supported...exiting..."
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
    
    if($PatchGroupTest | Select-String $CurrentDate){
        Write-PatchLog "Computer Name: $Env:COMPUTERNAME"
        Write-PatchLog "OS: $ClientOS"
        if(!(Get-Module -Name "PSWindowsUpdate")){
            Write-PatchLog "Installing PSWindowsUpdate module..."
            InstallPSWindowsUpdate
        }
        if(((Get-Date).TimeOfDay.TotalHours -lt "5")){
            InstallUpdatesWithForcedReboot
        }
        else{
            InstallUpdatesWithNoReboot
        }
    }
    else{
        Write-PatchLog "Outside of patch window...$Env:COMPUTERNAME patches on the following days this patch cycle:`r`n`r`n$PatchGroupTest"
        EXIT
    }
}
