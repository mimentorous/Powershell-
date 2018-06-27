##################################################################################
# Skript startet die von MS gestellten Test-CMD-Lets und überprüft sie auf Fehler. 
# Eine "geht / geht nicht"  - Anzeige erfolgt in der Console 
# für detailinformationen wird ein Log auf dem YOUR_SERVER_NAME_HERE\d$\Exchange_check$ geschrieben.
#
# folgende Tests werden durchgeführt:
# test-mailflow 
# Get-StorageGroupCopyStatus
# test-replicationhealth
# test-mapiconnectivity
# test-servicehealth
# Test-ImapConnectivity 
# Test-PopConnectivity 
# Test-OwaConnectivity 
# Test-OutlookWebServices
#
# Version 1.3
# Log: 1.3 - ZB added uptime 
# Log: 1.4 - Leo, added a random generator for CAS server name, to make script independent from hardcoded server names.
#
#############################################################################



# don´t change below this line
#-----------------------------------------------------------------------




# Anpassung Fenstergröße für besseren output
$pswindow = (get-host).ui.rawui
#$newsize = $pswindow.buffersize
#$newsize.width = 900
#$pswindow.buffersize = $newsize

# Befüllung grundlegender Variablen
$date = get-date -Format "yyyy.MM.dd-HH-mm-ss"
$date_log = Get-Date -Format yyyy.MM.dd
$whoami = whoami
$local_computer = hostname

#################  Obtaining RANDOM CAS name to use in test commands  ########################
$AllCAS = Get-ClientAccessServer | Where-Object { $_.name -ilike "sw*"}
$CAS = Get-Random -InputObject $Allcas
#########################################

# Auslesen des Softwarestandes
Write-Host "`nCurrently installed Exchange Programms / Updates:" -ForegroundColor Yellow
Get-ChildItem HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall |foreach-object {Get-ItemProperty $_.PsPath}  | where {$_.displayname -like "*Exchange Server*"} | Sort-Object displayname | select-object displayname | foreach {$_.displayname}
Write-Host "`n`n"
$resRaw = gwmi -class Win32_OperatingSystem -ComputerName $local_computer
$resRaw | select  @{n="lastBootUpTime";e={[Management.ManagementDateTimeConverter]::ToDateTime($_.LastBootUpTime)}}, @{n="datediff";e={$datestart = [Management.ManagementDateTimeConverter]::ToDateTime($_.LastBootUpTime);(New-TimeSpan  -Start $datestart).Days}}

$ErrorActionPreference = "silentlycontinue"
$cmss = Get-ClusteredMailboxServerStatus
$ErrorActionPreference = "continue"
$clustername = (($cmss).identity.name)  
if (test-path \\YOUR_SERVER_NAME_HERE\Exchange_check$) {$local_log_path = "\\YOUR_SERVER_NAME_HERE\Exchange_check$\$date_log" }
else {$local_log_path = "c:" }
if($local_log_path -like "*YOUR_SERVER_NAME_HERE*")
{
	if ($cmss) {$logfile = "$local_log_path\$clustername\"+$clustername+"_"+$local_computer+"_"+$Date+".txt"}
	ELSE {  $logfile = "$local_log_path\$local_computer\"+$local_computer+"_"+$Date+".txt"  }
}
else {
	if ($cmss) {$logfile = "$local_log_path\"+$clustername+"_"+$local_computer+"_"+$Date+".txt"}
	ELSE {  $logfile = "$local_log_path\"+$local_computer+"_"+$Date+".txt"  }
}



# prüfen ob Server aktiv ist, oder kein Cluster
if ((!($cmss)) -or ($(($cmss.OperationalMachines | where {$_ -like "*active*"}).split(" "))[0] -like "$local_computer")) 
{
	
	# Anlage Folder für Log auf YOUR_SERVER_NAME_HERE
	$ErrorActionPreference = "silentlycontinue"
	if (!(Get-Item \\YOUR_SERVER_NAME_HERE\exchange_check$\$date_log)) { New-Item \\YOUR_SERVER_NAME_HERE\Exchange_check$\$date_log -ItemType directory | Out-Null }
	if ($cmss) {if (!(Get-Item \\YOUR_SERVER_NAME_HERE\exchange_check$\$date_log\$clustername)) { New-Item \\YOUR_SERVER_NAME_HERE\Exchange_check$\$date_log\$clustername -ItemType directory | Out-Null }}
	else {if (!(Get-Item \\YOUR_SERVER_NAME_HERE\exchange_check$\$date_log\$local_computer)) { New-Item \\YOUR_SERVER_NAME_HERE\Exchange_check$\$date_log\$local_computer -ItemType directory | Out-Null }}
	$ErrorActionPreference = "continue"
	
	Write-Host "#############################################"
	Write-Host "#  Start Exchange-Check                     #"
	if ($cmss) { Write-Host "#  logged in as $whoami on $local_computer on Cluster $clustername  #"     }
	else {  Write-Host "#  logged in as $whoami on $local_computer  #"  }
	Write-Host "#  $date        $cas.name               #"
	Write-Host "#############################################"
	
	"#############################################" | out-file $logfile -Append
	"#  Start Exchange-Check                     #" | out-file $logfile -Append
	if ($cmss) { "#  logged in as $whoami on $local_computer on Cluster $clustername  #" | out-file $logfile -Append    }
	else { "#  logged in as $whoami on $local_computer  #" | out-file $logfile -Append  }
	"#  $date                       #" | out-file $logfile -Append
	"#############################################" | out-file $logfile -Append
	
	
	
	
	# Starte Abfrage für Mailbox
	if ($cmss -or ((Get-exchangeServer $local_computer).serverrole -like "*mailbox*"))
	{
	
		####################################################################
		
		$test_mailflow = test-mailflow -AutoDiscoverTargetMailboxServer
		"test-mailflow:" | out-file $logfile -Append
		$test_mailflow | ft * -AutoSize  | out-file $logfile -Append
		$TestMailflowResult = $test_mailflow.TestMailflowResult
		if ($TestMailflowResult -like "success") { Write-Host "test-mailflow - OK" -ForegroundColor green}
		else { write-host "test-mailflow - ERROR" -ForegroundColor red; $errorlogid=1 } 
		
		########################################################################
		
		if ($cmss)
		{
			Do
			{
				
				$sgcs = Get-StorageGroupCopyStatus
				"get-storageGroupCopyStatus"  | out-file $logfile -Append
				$sgcs | ft * -AutoSize  | out-file $logfile -Append
				$errorid=0
				foreach ($line in $sgcs)
				{
					if ($line.summarycopystatus -notlike "Healthy")
					{
						$dbname = $line.storagegroupname
						Write-Host "Copystatus of store $dbname is not healthy" -ForegroundColor Red
						$errorid=1
						$errorlogid=1
					}
				}
				$i++ 
				if ($errorid -eq 1) {write-host "start waiting Time"; Start-Sleep -Seconds 10}
			}
			Until ($errorid -eq 0 -or $i -gt 6)
			if ($errorid -eq 1)
			{
				Write-Host "Get-StorageGroupCopyStatus is not Healthy" -ForegroundColor Red
			}
			else
			{
				Write-Host "Get-StorageGroupCopyStatus is Healthy" -ForegroundColor Green
			}
			#################################################
			
			$trh = Test-ReplicationHealth
			"test-ReplicationHealth"  | out-file $logfile -Append
			$trh | ft * -AutoSize  | out-file $logfile -Append
			$errorid = 0
			foreach ($line in $trh)
			{
				if ($line.result.value -notlike "Passed")
				{
					$errorid = 1
				}
			}
			
			if ($errorid -eq 1)
			{
				Write-Host "test-replicationhealth is not OK" -ForegroundColor Red 
				$errorlogid=1
			}			
			
			if ($errorid -eq 0)
			{
				Write-Host "test-replicationhealth is OK" -ForegroundColor Green
			}
		}
		#####################################################
		
		$tmc = Test-MAPIConnectivity
		"Test-MAPIConnectivity"  | out-file $logfile -Append
		$tmc | ft * -AutoSize  | out-file $logfile -Append
		$errorid = 0
		foreach ($line in $tmc)
		{
			if ($line.result.value -notlike "success")
			{
				$errorid = 1
				$errorlogid=1
			}
		}
		
		if ($errorid -eq 1)
		{
			Write-Host "test-mapiconnectivity is not OK" -ForegroundColor Red 
			
		}			
		
		if ($errorid -eq 0)
		{
			Write-Host "test-mapiconnectivity is OK" -ForegroundColor Green 
		}
		##############################################################
		
		if ($cmss) { $tic = Test-ImapConnectivity -ClientAccessServer $cas.Name.ToString()  -MailboxServer $clustername }
		else {$tic = Test-ImapConnectivity -ClientAccessServer $cas.Name.ToString() -MailboxServer $local_computer }
		"Test-ImapConnectivity" | out-file $logfile -Append
		$tic | ft * -AutoSize  | out-file $logfile -Append
		$errorid = 0
		foreach ($line in $tic)
		{
			if ($line.result.value -notlike "success")
			{
				$errorid = 1
				$errorlogid=1
			}
		}
		
		if ($errorid -eq 1)
		{
			Write-Host "Test-ImapConnectivity is not OK" -ForegroundColor Red 
			
		}			
		
		if ($errorid -eq 0)
		{
			Write-Host "Test-ImapConnectivity is OK" -ForegroundColor Green 
		}
		
		
		if ($cmss) { $tpc = Test-popConnectivity -ClientAccessServer $cas.Name.ToString() -MailboxServer $clustername }
		else {$tpc = Test-popConnectivity -ClientAccessServer $cas.Name.ToString() -MailboxServer $local_computer }
		"Test-PopConnectivity" | out-file $logfile -Append
		$tpc | ft * -AutoSize  | out-file $logfile -Append
		$errorid = 0
		foreach ($line in $tpc)
		{
			if ($line.result.value -notlike "success")
			{
				$errorid = 1
				$errorlogid=1
			}
		}
		
		if ($errorid -eq 1)
		{
			Write-Host "Test-PopConnectivity is not OK" -ForegroundColor Red 
			
		}			
		
		if ($errorid -eq 0)
		{
			Write-Host "Test-PopConnectivity is OK" -ForegroundColor Green 
		}
		
		
		if ($cmss) { $toc = Test-owaConnectivity -ClientAccessServer $cas.Name.ToString() -MailboxServer $clustername }
		else {$toc = Test-owaConnectivity -ClientAccessServer $cas.Name.ToString() -MailboxServer $local_computer }
		"Test-Owaconnectivity"  | out-file $logfile -Append
		$toc | ft * -AutoSize  | out-file $logfile -Append
		$errorid = 0
		foreach ($line in $toc)
		{
			if ($line.result.value -notlike "success")
			{
				$errorid = 1
				$errorlogid=1
			}
		}
		
		if ($errorid -eq 1)
		{
			Write-Host "Test-OwaConnectivity is not OK" -ForegroundColor Red 
			
		}			
		
		if ($errorid -eq 0)
		{
			Write-Host "Test-OwaConnectivity is OK" -ForegroundColor Green 
		}
		
		
		
		
		
	}
	
	# Starte Abfrage für CAS
	if ((!($cmss)) -and ((Get-exchangeServer $local_computer).serverrole -like "*Client*"))
	{
		$tic = Test-ImapConnectivity
		"Test-ImapConnectivity" | out-file $logfile -Append
		$tic | ft * -AutoSize  | out-file $logfile -Append
		$errorid = 0
		foreach ($line in $tic)
		{
			if ($line.result.value -notlike "success")
			{
				$errorid = 1
				$errorlogid=1
			}
		}
		
		if ($errorid -eq 1)
		{
			Write-Host "Test-ImapConnectivity is not OK" -ForegroundColor Red 
			
		}			
		
		if ($errorid -eq 0)
		{
			Write-Host "Test-ImapConnectivity is OK" -ForegroundColor Green 
		}
		
		
		$tpc = Test-PopConnectivity
		"Test-PopConnectivity" | out-file $logfile -Append
		$tpc | ft * -AutoSize  | out-file $logfile -Append
		$errorid = 0
		foreach ($line in $tpc)
		{
			if ($line.result.value -notlike "success")
			{
				$errorid = 1
				$errorlogid=1
			}
		}
		
		if ($errorid -eq 1)
		{
			Write-Host "Test-PopConnectivity is not OK" -ForegroundColor Red 
			
		}			
		
		if ($errorid -eq 0)
		{
			Write-Host "Test-PopConnectivity is OK" -ForegroundColor Green 
		}
		
		
		$toc = Test-OwaConnectivity
		"Test-Owaconnectivity"  | out-file $logfile -Append
		$toc | ft * -AutoSize  | out-file $logfile -Append
		$errorid = 0
		foreach ($line in $toc)
		{
			if ($line.result.value -notlike "success")
			{
				$errorid = 1
				$errorlogid=1
			}
		}
		
		if ($errorid -eq 1)
		{
			Write-Host "Test-OwaConnectivity is not OK" -ForegroundColor Red 
			
		}			
		
		if ($errorid -eq 0)
		{
			Write-Host "Test-OwaConnectivity is OK" -ForegroundColor Green 
		}
		
		
		$tows = Test-OutlookWebServices
		"Test-OutlookWebServices"  | out-file $logfile -Append
		$tows | ft * -AutoSize  | out-file $logfile -Append
		$errorid = 1
		foreach ($line in $tows)
		{
			if ($line.type -like "Information" -or $line.type -like "success")
			{
				$errorid = 0
			}
		}
		
		if ($errorid -eq 1)
		{
			Write-Host "Test-OutlookWebServices is not OK" -ForegroundColor Red 
			$errorlogid=1
		}			
		
		if ($errorid -eq 0)
		{
			Write-Host "Test-OutlookWebServices is OK" -ForegroundColor Green 
		}
		
		

        $t = $cas.name + ".wwg00m.rootdom.net"
        $rpcping = rpcping -t ncacn_http -s wmucv465.wwg00m.rootdom.net -o RpcProxy=$t -H 2 -u 10 -a connect -F 3 -v 3 -E -R none
		$msg = $rpcping[-1]
        "Test-RPCPing"  | out-file $logfile -Append
		$rpcping | ft * -AutoSize  | out-file $logfile -Append
		$errorid = 1
       
        foreach ($ping in $rpcping){
         
		if ($ping -ilike "*successfully*") 
		{
			$errorid = 0
		}
        
        }
		
		if ($errorid -eq 1)
		{
			Write-Host "Test-RPC-Ping is not OK" -ForegroundColor Red 
			$errorlogid=1
		}			
		
		if ($errorid -eq 0)
		{
			Write-Host "Test-RPC-Ping is OK" -ForegroundColor Green 
		}
		
	
	}
	
	# all
	$Tsh = test-servicehealth
	"Test-ServiceHealth"  | out-file $logfile -Append
	$tsh | ft * -AutoSize  | out-file $logfile -Append
	$errorid = 0
	foreach ($line in $tsh)
	{
		if ($line.requiredservicesrunning -eq $false)
		{
			$errorid = 1
			$errorlogid=1
		}
	}
	
	if ($errorid -eq 1)
	{
		Write-Host "test-servicehealth is not OK" -ForegroundColor Red 
		
	}			
	
	if ($errorid -eq 0)
	{
		Write-Host "test-servicehealth is OK" -ForegroundColor Green 
	}
	

	
	
	$enddate = get-date -Format "yyyy.MM.dd-HH-mm-ss"
	Write-Host "`nDetailled log: $logfile"
	Write-Host ""
	Write-Host "##########################"
	Write-Host "#  End Exchange-Check   #"
	Write-Host "#  $enddate                 #"
	Write-Host "##########################"
	
	"##########################" | out-file $logfile -Append
	"#  End Exchange-Check   #" | out-file $logfile -Append
	"#  $enddate    #" | out-file $logfile -Append
	"##########################" | out-file $logfile -Append
	
	
	if ($cmss) {$logfile = "\\YOUR_SERVER_NAME_HERE\Exchange_Check$\$date_log\$clustername\"+$clustername+"_"+$local_computer+"_"+$Date+".txt"}
	ELSE {  $logfile = "\\YOUR_SERVER_NAME_HERE\Exchange_Check$\$date_log\$local_computer\"+$local_computer+"_"+$Date+".txt"  }
	if ($errorlogid -eq 1) 
	{ 
		if ($cmss) {$newname = "ERR_"+$clustername+"_"+$local_computer+"_"+$Date+".txt"; Rename-Item $logfile -NewName $newname }
		else {$newname = "ERR_"+$local_computer+"_"+$Date+".txt"; Rename-Item $logfile -NewName $newname  }
	}

}

else 
{
	write-host "Local Computer is a passive node. Please re-run script on a active node." -ForegroundColor Yellow
    Get-StorageGroupCopyStatus | where {$_.CopyQueueLength -gt 3 -or $_.ReplayQueueLength -gt 3}| ft -a
    
}
