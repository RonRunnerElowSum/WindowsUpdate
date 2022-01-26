$TaskName = "(MSP) Install Missing Windows Updates (PatchGroup -- PRODUCTION)"
$PSFileURL = "'https://raw.githubusercontent.com/RonRunnerElowSum/WindowsUpdate/Prod-Branch/Install%20Windows%20Updates%20--%20PRODUCTION%20group.ps1'"

$ScheduledTaskCmd = "Invoke-WebRequest -URI $PSFileURL -UseBasicParsing | Invoke-Expression; PunchIt"
$ScheduledTaskArg = "-WindowStyle Hidden -Command `"& {$ScheduledTaskCmd}`" -Verb RunAs"

function CreateSchedTask () {
    $Action = New-ScheduledTaskAction -Execute "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" -Argument $ScheduledTaskArg
    $Trigger = @()
    $Trigger += New-ScheduledTaskTrigger -AtStartup
    $Trigger += New-ScheduledTaskTrigger -Daily -At 12am
    $Trigger += New-ScheduledTaskTrigger -Daily -At 10am
    $Trigger += New-ScheduledTaskTrigger -Daily -At 3pm
    $Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -WakeToRun -StartWhenAvailable
    Register-ScheduledTask -Action $Action -Trigger $Trigger -Settings $Settings -User "NT AUTHORITY\SYSTEM" -TaskName $TaskName -RunLevel Highest -Description "Keeps Windows up-to-date." | Out-Null
    if(Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue){
        Write-Host "Successfully created scheduled task!"
    }
    else{
        Write-Warning "An error occured. Failed to create schedule task.`r`n`r`nDetails:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())"
    }
}

function PunchIt () {
    $ClientOS = (Get-WmiObject -class Win32_OperatingSystem).Caption
    if(($ClientOS | Select-String "Windows 7") -or ($ClientOS | Select-String "Server 2003") -or ($ClientOS | Select-String "2008")){
        Write-Host "This operating system ($ClientOS) is no longer supported...exiting..."
        EXIT
    }
    if((Get-WindowsOptionalFeature -Online | Where-Object {$_.FeatureName -eq "Microsoft-Hyper-V"} | Where-Object {$_.State -eq "Enabled"}) -and ($EndpointOS | Select-String "Server")){
        Write-Host "$Env:ComputerName is a Hyper-V server and should use the 'Windows Update Hyper-V' scheduled task...exiting..."
        EXIT
    }
    if(Get-ScheduledTask -TaskName "(MSP) Install Missing Windows Updates (PatchGroup -- TEST)" -ErrorAction SilentlyContinue){
        Unregister-ScheduledTask -TaskName "(MSP) Install Missing Windows Updates (PatchGroup -- TEST)" -Confirm:$False | Out-Null
    }
    if(Get-ScheduledTask -TaskName "(MSP) Install Missing Windows Updates (PatchGroup -- NO REBOOT)" -ErrorAction SilentlyContinue){
        #Unregister-ScheduledTask -TaskName "(MSP) Install Missing Windows Updates (PatchGroup -- NO REBOOT)" -Confirm:$False | Out-Null
        Write-Host "The scheduled task named ((MSP) Install Missing Windows Updates (PatchGroup -- NO REBOOT)) already exists...exiting..."
        EXIT
    }
    if(Get-ScheduledTask -TaskName "(MSP) Install Missing Windows Updates (PatchGroup -- PILOT)" -ErrorAction SilentlyContinue){
        Unregister-ScheduledTask -TaskName "(MSP) Install Missing Windows Updates (PatchGroup -- PILOT)" -Confirm:$False | Out-Null
    }
    if(Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue){
        Write-Host "The scheduled task [$TaskName] already exists..."
    }
    else{
        CreateSchedTask
    }
}
