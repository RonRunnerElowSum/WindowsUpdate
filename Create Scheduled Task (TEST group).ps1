$TaskName = "(MSP) Install Missing Windows Updates (PatchGroup -- TEST)"
$PS1URL = '"https://raw.githubusercontent.com/RonRunnerElowSum/WindowsUpdate/Prod-Branch/Install%20Windows%20Updates%20(TEST%20group).ps1"'
$LaunchString = "Invoke-WebRequest -URI $PS1URL -UseBasicParsing | Invoke-Expression; PunchIt"

function CreateSchedTask () {
    try{
        $Action = New-ScheduledTaskAction -Execute "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" -Argument "-WindowStyle Hidden -Command & {$LaunchString} -Verb RunAs"
        $Trigger = @()
        $Trigger += New-ScheduledTaskTrigger -AtStartup
        $Trigger += New-ScheduledTaskTrigger -Daily -At 12am
        $Trigger += New-ScheduledTaskTrigger -Daily -At 10am
        $Trigger += New-ScheduledTaskTrigger -Daily -At 3pm
        $Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -WakeToRun -StartWhenAvailable
        Register-ScheduledTask -Action $Action -Trigger $Trigger -Settings $Settings -User "NT AUTHORITY\SYSTEM" -TaskName $TaskName -RunLevel Highest -Description "Keeps Windows up-to-date." | Out-Null
    }
    catch{
        Write-Warning "An error occured. Failed to create schedule task.`r`n`r`nDetails:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())"
    }
    if(Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue){
        Write-Host "Successfully created scheduled task!"
    }
}

if(Get-ScheduledTask -TaskName "(MSP) Install Missing Windows Updates (PatchGroup -- PRODUCTION)" -ErrorAction SilentlyContinue){
    Unregister-ScheduledTask -TaskName "(MSP) Install Missing Windows Updates (PatchGroup -- PRODUCTION)" -Confirm:$False | Out-Null
}
if(Get-ScheduledTask -TaskName "(MSP) Install Missing Windows Updates (PatchGroup -- NO REBOOT)" -ErrorAction SilentlyContinue){
    Unregister-ScheduledTask -TaskName "(MSP) Install Missing Windows Updates (PatchGroup -- NO REBOOT)" -Confirm:$False | Out-Null
}
if(Get-ScheduledTask -TaskName "(MSP) Install Missing Windows Updates (PatchGroup -- PILOT)" -ErrorAction SilentlyContinue){
    Unregister-ScheduledTask -TaskName "(MSP) Install Missing Windows Updates (PatchGroup -- PILOT)" -Confirm:$False | Out-Null
}
if(Get-ScheduledTask -TaskName "(MSP) Install Missing Windows Updates (PatchGroup -- TEST)" -ErrorAction SilentlyContinue){
    Write-Host "The scheduled task [(MSP) Install Missing Windows Updates (PatchGroup -- TEST)] already exists..."
}
CreateSchedTask
