$TaskName = "(MSP) Install Missing Windows Updates (PatchGroup -- PILOT)"
$PS1URL = "https://raw.githubusercontent.com/RonRunnerElowSum/WindowsUpdate/Prod-Branch/Install%20Windows%20Updates%20(PILOT%20group).ps1"
$LaunchString = "Invoke-WebRequest -URI $PS1URL -UseBasicParsing | Invoke-Expression; PunchIt"

function CreateSchedTask () {
    $Action = New-ScheduledTaskAction -Execute 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe' -Argument "-WindowStyle Hidden $LaunchString"
    $Trigger = @()
    $Trigger += New-ScheduledTaskTrigger -Daily -At 12am
    $Trigger += New-ScheduledTaskTrigger -Daily -At 10am
    $Trigger += New-ScheduledTaskTrigger -Daily -At 3pm
    $Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -WakeToRun
    Register-ScheduledTask -Action $Action -Trigger $Trigger -Settings $Settings -User SYSTEM -TaskName $TaskName -RunLevel Highest -Description "Keeps Windows up-to-date." | Out-Null
    if(!(Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue)){
        Write-Warning "Failed to create schedule task...exiting..."
        EXIT
    }
    else{
        Write-Host "Successfully created scheduled task!"
        EXIT
    }
}

if(Get-ScheduledTask -TaskName "(MSP) Install Missing Windows Updates (PatchGroup -- TEST)" -ErrorAction SilentlyContinue){
    Unregister-ScheduledTask -TaskName "(MSP) Install Missing Windows Updates (PatchGroup -- TEST)" -Confirm:$False | Out-Null
}
if(Get-ScheduledTask -TaskName "(MSP) Install Missing Windows Updates (PatchGroup -- NO REBOOT)" -ErrorAction SilentlyContinue){
    Unregister-ScheduledTask -TaskName "(MSP) Install Missing Windows Updates (PatchGroup -- NO REBOOT)" -Confirm:$False | Out-Null
}
if(Get-ScheduledTask -TaskName "(MSP) Install Missing Windows Updates (PatchGroup -- PRODUCTION)" -ErrorAction SilentlyContinue){
    Unregister-ScheduledTask -TaskName "(MSP) Install Missing Windows Updates (PatchGroup -- PRODUCTION)" -Confirm:$False | Out-Null
}
if(Get-ScheduledTask -TaskName "(MSP) Install Missing Windows Updates (PatchGroup -- PILOT)" -ErrorAction SilentlyContinue){
    Write-Host "The scheduled task [(MSP) Install Missing Windows Updates (PatchGroup -- NO PILOT)] already exists...exiting..."
    EXIT
}
CreateSchedTask
