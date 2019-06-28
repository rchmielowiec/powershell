$task_name = 'run script as system'
$command_to_run = 'powershell.exe' 
$command_arguments = '-NoProfile -WindowStyle Hidden -command "& {whoami >> c:\whoami.txt}"'


$action = New-ScheduledTaskAction -Execute $command_to_run -Argument $command_arguments

$trigger = New-ScheduledTaskTrigger -At $(get-date).AddMinutes(1) -Once

Register-ScheduledTask -TaskName $task_name -User 'NT AUTHORITY\SYSTEM' -Description 'running script as system account' -Trigger $trigger -Action $action

Start-ScheduledTask -TaskName $task_name

Start-Sleep -Seconds 30 

$task_info = Get-ScheduledTaskInfo -TaskName $task_name

Unregister-ScheduledTask -TaskName $task_name -Confirm:$false

if ($task_info.LastRunTime -and $($task_info.LastTaskResult -eq 0))
    {
        Write-Host "Done"
    }
else
    {
        Write-Error "Running task as system failed"
    }

