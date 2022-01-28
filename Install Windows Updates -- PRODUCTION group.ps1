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
        Write-MSPLog -LogSource "MSP Patch Health" -LogType "Error" -LogMessage "The module PSWindowsUpdate failed to install...exiting..."
        EXIT
    }
}

function InstallUpdatesWithNoReboot () {
    $BlacklistedPatches = (Invoke-WebRequest -URI "https://raw.githubusercontent.com/RonRunnerElowSum/WindowsUpdate/Prod-Branch/BlackListedPatches.cfg" -UseBasicParsing).Content
    Write-MSPLog -LogSource "MSP Patch Health" -LogType "Information" -LogMessage "Checking for Windows Updates..."
    $MissingUpdates = (Get-WindowsUpdate -MicrosoftUpdate -NotCategory Drivers -NotTitle "Feature update to Windows 10" -NotKBArticleID $BlacklistedPatches).KB
    if(!($Null -eq $MissingUpdates)){
        if($MissingUpdates.Count -eq "1"){
            $FormattedMissingUpdates = $MissingUpdates
        }
        else{
            $FormattedMissingUpdates = [string]::Join("`r`n",($MissingUpdates))
        }
        Write-MSPLog -LogSource "MSP Patch Health" -LogType "Warning" -LogMessage "$Env:COMPUTERNAME is missing the following ($($MissingUpdates.Count)) patches:`r`n$FormattedMissingUpdates"
        Write-MSPLog -LogSource "MSP Patch Health" -LogType "Information" -LogMessage "Installing missing updates..."
        $FormattedMissingUpdates | ForEach-Object {
            [string]$InstallUpdates = Install-WindowsUpdate -KBArticleID "$_" -IgnoreReboot -Confirm:$False | Out-String
            Write-MSPLog -LogSource "MSP Patch Health" -LogType "Information" -LogMessage $InstallUpdates
        }
        CheckPendingRebootStatus
    }
    else{
        Write-MSPLog -LogSource "MSP Patch Health" -LogType "Information" -LogMessage "Windows is up-to-date!"
    }
}

function InstallUpdatesWithForcedReboot () {
    $BlacklistedPatches = (Invoke-WebRequest -URI "https://raw.githubusercontent.com/RonRunnerElowSum/WindowsUpdate/Prod-Branch/BlackListedPatches.cfg" -UseBasicParsing).Content
    Write-MSPLog -LogSource "MSP Patch Health" -LogType "Information" -LogMessage "Checking for Windows Updates..."
    $MissingUpdates = (Get-WindowsUpdate -MicrosoftUpdate -NotCategory Drivers -NotTitle "Feature update to Windows 10" -NotKBArticleID $BlacklistedPatches).KB
    if(!($Null -eq $MissingUpdates)){
        if($MissingUpdates.Count -eq "1"){
            $FormattedMissingUpdates = $MissingUpdates
        }
        else{
            $FormattedMissingUpdates = [string]::Join("`r`n",($MissingUpdates))
        }
        Write-MSPLog -LogSource "MSP Patch Health" -LogType "Warning" -LogMessage "$Env:COMPUTERNAME is missing the following ($($MissingUpdates.Count)) patches:`r`n$FormattedMissingUpdates"
        Write-MSPLog -LogSource "MSP Patch Health" -LogType "Information" -LogMessage "Installing missing updates..."
        $FormattedMissingUpdates | ForEach-Object {
            [string]$InstallUpdates = Install-WindowsUpdate -KBArticleID "$_" -AutoReboot -Confirm:$False | Out-String
            Write-MSPLog -LogSource "MSP Patch Health" -LogType "Information" -LogMessage $InstallUpdates
        }
    }
    else{
        Write-MSPLog -LogSource "MSP Patch Health" -LogType "Information" -LogMessage "Windows is up-to-date!"
        EXIT
    }
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

function PunchIt () {
    Write-MSPLog -LogSource "MSP Patch Health" -LogType "Information" -LogMessage "Starting..."
    $ClientOS = Get-WmiObject -class Win32_OperatingSystem | Select-Object -ExpandProperty Caption
    if(($ClientOS | Select-String "Windows 7") -or ($ClientOS | Select-String "2003") -or ($ClientOS | Select-String "2008")){
        Write-MSPLog -LogSource "MSP Patch Health" -LogType "Warning" -LogMessage "OS: $ClientOS`r`n`This operating system is no longer supported...exiting..."
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
        Write-MSPLog -LogSource "MSP Patch Health" -LogType "Information" -LogMessage "Computer Name: $Env:COMPUTERNAME`r`nOS: $ClientOS"
        if(!(Get-Module -Name "PSWindowsUpdate")){
            Write-MSPLog -LogSource "MSP Patch Health" -LogType "Information" -LogMessage "Installing PSWindowsUpdate module..."
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
        Write-MSPLog -LogSource "MSP Patch Health" -LogType "Information" -LogMessage "Outside of patch window...$Env:COMPUTERNAME patches on the following days this patch cycle:`r`n`r`n$PatchGroupProduction"
        EXIT
    }
}
