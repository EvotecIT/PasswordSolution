$Tasks = Get-ScheduledTask -TaskPath "\" -TaskName "Automated-PasswordSolution"

# Fix all tasks to use proper account
foreach ($Task in $Tasks) {
    schtasks /Change /TN $Task.TaskName /RU "GMSA$" /RP ""
}
foreach ($Task in $Tasks) {
    #Start-ScheduledTask -TaskName $Task.TaskName -Verbose
}