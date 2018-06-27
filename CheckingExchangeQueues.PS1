#Making sure the SOURCE exists so that Write-EventLog will not throw error
if ([System.Diagnostics.EventLog]::SourceExists("MS Exchange Queues Monitoring") -eq $false) {New-Eventlog -LogName "Application" -Source "MS Exchange Queues Monitoring"}

#Retrieving all queues in SUSPENDED state or with queue size greater than 1000 messages (including shadow queues)
$queues = Get-TransportService | ?{$_.Name -imatch "dee|fre"} | sort name | %{Get-Queue -Server $_.Name | ?{$_.MessageCount -gt 1000 -or $_.Status -imatch "susp"}} | ft -a | Out-String

#Writing eventLog entry so that ITM can pick it up and open a ticket
if (!$queues) {Write-EventLog –LogName Application –Source “MS Exchange Queues Monitoring” –EntryType Information –EventID 11111 –Message "MS Exchange Queues are OK."}
else {Write-EventLog –LogName Application –Source “MS Exchange Queues Monitoring” –EntryType Warning –EventID 12345 –Message "Please check following queues $queues"}

