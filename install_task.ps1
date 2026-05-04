$pythonw = "C:\Python313\pythonw.exe"
$launcher = "C:\Users\dani1\pc-agent\launcher.py"
$workdir = "C:\Users\dani1\pc-agent"
$action = New-ScheduledTaskAction -Execute $pythonw -Argument $launcher -WorkingDirectory $workdir
$trigger = New-ScheduledTaskTrigger -AtLogOn -User "dani1"
$settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit 0 -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1) -StartWhenAvailable -DontStopIfGoingOnBatteries -AllowStartIfOnBatteries
Unregister-ScheduledTask -TaskName "AdamBot" -Confirm:$false -ErrorAction SilentlyContinue
Register-ScheduledTask -TaskName "AdamBot" -Action $action -Trigger $trigger -Settings $settings -RunLevel Highest -Force
Start-ScheduledTask -TaskName "AdamBot"
Write-Host "Done"