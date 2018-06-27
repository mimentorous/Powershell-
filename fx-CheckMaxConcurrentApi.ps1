PARAM ([Switch]$CheckMaxConcurrentApi, [switch]$GetNetlogonInstances, [string]$Computer = "Localhost", [string]$Instance = "_Total", [bool]$CalcMCA = $False)
#************************************************
# CheckMaxConcurrentApiScript.ps1
# Version 1.0
# Date: 3/22/2013
# Author: Tim Springston [MSFT]
# Description:Uses .Net methods and WMI to query
#  			  a remote or local computer for MaxConcurrentApi
# 			  problems and return details back in a PSObject.
#************************************************
function CheckMaxConcurrentApi ([string]$InstanceName = "_Total", [string]$ComputerName = "localhost", [bool]$Calc = $false)
{   
	#This function takes three optional parameters to select Netlogon Instance (can be obtained by using
	#sister function GetNetlogonInstances, computer to run against and whether to run
	#MaxConcurrentApi calculation-which takes longer.
    #It returns details about the computer, whether the problem is detected, and suggested MaxConcurrentApi value.
	
	$ProblemDetected = $false
	$Date = Get-Date

	#Get role, OSVer, hotfix data.
	$cs =  gwmi -Namespace "root\cimv2" -class win32_computersystem -Impersonation 3 -ComputerName $ComputerName
	$DomainRole = $cs.domainrole
	$OSVersion = gwmi -Namespace "root\cimv2" -Class Win32_OperatingSystem -Impersonation 3 -ComputerName $ComputerName
	$bn = $OSVersion.BuildNumber
	$150Hotfix = $false

	#Determine how long the computer has been running since last reboot.
	$wmi = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $ComputerName
	$LocalDateTime = $wmi.LocalDateTime
	$Uptime = $wmi.ConvertToDateTime($wmi.LocalDateTime) – $wmi.ConvertToDateTime($wmi.LastBootUpTime)
	$Days = $Uptime.Days.ToString()
	$Hours = $Uptime.Hours.ToString()
	$UpTimeStatement = $Days + " days " + $Hours + " hours"

	#Get SystemRoot so that we can map the right drive for checking file versions.
	$objReg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$ComputerName)
	$ObjRegKey = $ObjReg.OpenSubKey("SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion")
	$SystemRoot = $ObjRegKey.GetValue("SystemRoot")

	#Parse the drive letter for the remote systemroot and then map the network drive. First though
	#make sure the drive you will map to doesnt exist already.
	$Drives = gwmi -Namespace "root\cimv2"  -ComputerName $ComputerName -Impersonation Impersonate -Class Win32_LogicalDisk
	$DriveLetters = @()
	ForEach ($Drive in $Drives)
		{
		  $Caption = $Drive.Caption
		  $DriveLetters += $Caption
		  $Caption = $null
		}

	if ($ComputerName -ne "localhost")
		{
			$PossibleLetters = [char[]]"DEFGHIJKLMNOPQRTUVWXY" 
			$devices = get-wmiobject win32_logicaldisk | select -expand DeviceID
			$PossibleLetters | where {$DriveLetters -notcontains "$($_):"} | Select -First 1 | % {
			$AvailableDriveLetter = $_
			}
			$DriveToMap = $SystemRoot
			$DriveToMap = $DriveToMap.Replace(":\Windows","")
			$DriveToMap = "\\" + $ComputerName + "\" + $DriveToMap + "$"
			$AvailableDriveLetter = $AvailableDriveLetter  + ":"
			$NETUSEReturn = net use $AvailableDriveLetter $DriveToMap
			$RemoteSystem32Folder = $AvailableDriveLetter + "\Windows\System32"
			$NetlogonDll = $RemoteSystem32Folder + "\Netlogon.dll"
	
		}
		else
			{
			 $NetlogonDll = $SystemRoot + "\System32\Netlogon.dll"
			}

	#Check the file versions for the hotfixes.
	$FileVer = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($NetlogonDll).FileVersion
	switch -exact ($bn)
		{"6002" {#Hotfix Check for MCA to 150 KB975363 http://support.microsoft.com/kb/975363
				 $6002HotfixVer = "6.0.6002.22289"
				 if ($6002HotfixVer -eq $FileVer)
				 	{$150Hotfix = $true}
				}
		"7600" {#Hotfix Check for MCA to 150 KB975363 http://support.microsoft.com/kb/975363
				$7600HotfixVer = "6.1.7600.20576"
				if ($7600HotfixVer -eq $FileVer)
					{$150Hotfix = $true}
				}
		"7601" {#Hotfix Check for MCA to 150 KB975363 http://support.microsoft.com/kb/975363
				$150Hotfix = $true
				}
		}
	if ($ComputerName -ne "localhost")
		{
		$NETUSEReturn = net use $AvailableDriveLetter /d
		}

	#Determine effective MaxConcurrentApi setting based on OS, hotfix presence, role and registry setting.
	$objReg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $ComputerName)
	$objRegKey= $objReg.OpenSubKey("SYSTEM\\CurrentControlSet\\services\\Netlogon\\Parameters")
	$MCARegVal = $objRegKey.GetValue('MaxConcurrentApi')
	$CurrentMCA = 0
	if  ($DomainRole -gt 1)
			{	
			if ($bn -ge 9200)
				{
				If (($MCARegVal -lt 10) -or ($MCARegVal -eq $null) -or ($MCARegVal -gt 9999))
					{$CurrentMCA = 10}
					elseif (($MCARegVal -gt 10)	-and ($MCARegVal -lt 9999))
						{$CurrentMCA = $MCARegVal}
				}
				else
					{		
					If (($MCARegVal -gt 10) -and ($150Hotfix -eq $true))
						{		
						$CurrentMCA = $MCARegVal}
						elseif (($MCARegVal -gt 10) -and ($150Hotfix -eq $false))	
							{$CurrentMCA = 2}
							elseif (( $MCARegVal -gt 2 ) -and ($MCARegVal -le 10))
								{$CurrentMCA = $MCARegVal}
								elseif ($MCARegVal -lt 2)
									{$CurrentMCA = 2}
									elseif ($MCARegVal -eq $null)
										{$CurrentMCA = 2}
					}
				}

	#Get a sample of the counters.
	$Category = "Netlogon"
	$CounterASHT = "Average Semaphore Hold Time"
	$CounterST = "Semaphore Timeouts"
	$CounterSA = "Semaphore Acquires"
	$CounterSH = "Semaphore Holders"
	$CounterSW = "Semaphore Waiters"
	#Query remote computer for counters.
	$NetlogonRemoteASHT = New-Object System.Diagnostics.PerformanceCounter($Category,$CounterASHT,$InstanceName,$ComputerName)
	$NetlogonRemoteST = New-Object System.Diagnostics.PerformanceCounter($Category,$CounterST,$InstanceName,$ComputerName)
	$NetlogonRemoteSA = New-Object System.Diagnostics.PerformanceCounter($Category,$CounterSA,$InstanceName,$ComputerName)
	$NetlogonRemoteSW = New-Object System.Diagnostics.PerformanceCounter($Category,$CounterSW,$InstanceName,$ComputerName)
	$NetlogonRemoteSH = New-Object System.Diagnostics.PerformanceCounter($Category,$CounterSH,$InstanceName,$ComputerName)
	#Cook values
	$CookedASHT = $NetlogonRemoteASHT.NextValue()
	$CookedST = $NetlogonRemoteST.NextValue()
	$CookedSA = $NetlogonRemoteSA.NextValue()
	$CookedSW = $NetlogonRemoteSW.NextValue()
	$CookedSH = $NetlogonRemoteSH.NextValue()
	if ((($CookedSW -gt 0) -and (-not($CookedSW -gt 4GB))) -or ($CookedSH -eq $CurrentMCA) -or ((($CookedST -gt 0) -and (-not($CookedST -gt 4GB))) -and (($CookedSW -gt 0) -and (-not($CookedSW -gt 4GB)))))
				{$ProblemDetected = $true}

	#Do a second data sample and compare results in order to run the "suggested MCA" math.
	if (($ProblemDetected -eq $true) -and ($Calc -eq $true))
		{
		Start-Sleep -Seconds 60
		$NetlogonRemoteASHT = New-Object System.Diagnostics.PerformanceCounter($Category,$CounterASHT,$InstanceName,$ComputerName)
		$NetlogonRemoteST = New-Object System.Diagnostics.PerformanceCounter($Category,$CounterST,$InstanceName,$ComputerName)
		$NetlogonRemoteSA = New-Object System.Diagnostics.PerformanceCounter($Category,$CounterSA,$InstanceName,$ComputerName)
		$NetlogonRemoteSW = New-Object System.Diagnostics.PerformanceCounter($Category,$CounterSW,$InstanceName,$ComputerName)
		$NetlogonRemoteSH = New-Object System.Diagnostics.PerformanceCounter($Category,$CounterSH,$InstanceName,$ComputerName)
		#Cook values
		$SecondCookedASHT = $NetlogonRemoteASHT.NextValue()
		$SecondCookedST = $NetlogonRemoteST.NextValue()
		$SecondCookedSA = $NetlogonRemoteSA.NextValue()
		$SecondCookedSW = $NetlogonRemoteSW.NextValue()
		$SecondCookedSH = $NetlogonRemoteSH.NextValue()
		#Next, calculate the suggested MCA 
	       					 #using formula from http://support.microsoft.com/kb/2688798
	       					 #(semaphore_acquires + semaphore_timeouts) * average_semaphore_hold_time / time_collection_length =< New_MaxConcurrentApi_setting
							 #subtract Sample1SA from Sample2SA = SampleSADelta
							 $SampleSADelta = ($SecondCookedSA - $CookedSA)
							 $SampleSTDelta = ($SecondCookedST - $CookedST)
							 $ASHT = ($SecondCookedASHT + $CookedASHT)
							 $SampleASHTDelta = ($ASHT / 2 )
							 $SamplesDeltaSAST = ($SampleSADelta + $SampleSTDelta)
							 $AllSampleDeltas = ($SampleASHTDelta * $SamplesDeltaSAST)
						     $AllSampleDeltas /= 90
							 $SuggestedMCA = $AllSampleDeltas
							 $SuggestedMCA = "{0:N0}" -f $SuggestedMCA
							 if ($SuggestedMCA -le 2)
								{$SuggestedMCA = $CurrentMCA}
		
		}
	#Create PSObject for returned data.
	$ReturnedData = New-Object PSObject
	add-member -inputobject $ReturnedData  -membertype noteproperty -name "Detection Time" -value $Date
	add-member -inputobject $ReturnedData  -membertype noteproperty -name "Problem Detected" -value $ProblemDetected
	add-member -inputobject $ReturnedData  -membertype noteproperty -name "Server Name" -value $cs.Name
	if ($cs.DomainRole -le 1)
		{add-member -inputobject $ReturnedData  -membertype noteproperty -name "Server Role" -value "Client"}
	if (($cs.DomainRole -eq 3) -or ($cs.DomainRole -eq 2))
		{add-member -inputobject $ReturnedData  -membertype noteproperty -name "Server Role" -value "Member Server"}
	if ($cs.DomainRole -ge 4)
		{add-member -inputobject $ReturnedData  -membertype noteproperty -name "Server Role" -value "Domain Controller"}
	add-member -inputobject $ReturnedData  -membertype noteproperty -name "Domain Name" -value $cs.Domain
	add-member -inputobject $ReturnedData  -membertype noteproperty -name "Operating System" -value $OSVersion.Caption
	add-member -inputobject $ReturnedData  -membertype noteproperty -name "Time Since Last Reboot" -value $UpTimeStatement
	add-member -inputobject $ReturnedData  -membertype noteproperty -name "Current Effective MaxConcurrentApi Setting" -value $CurrentMCA
	if ($SuggestedMCA -eq $null)
		{add-member -inputobject $ReturnedData  -membertype noteproperty -name "Suggested MaxConcurrentApi Setting (may be same as current)" -value $CurrentMCA}
		else
		{add-member -inputobject $ReturnedData  -membertype noteproperty -name "Suggested MaxConcurrentApi Setting (may be same as current)" -value $SuggestedMCA}
	add-member -inputobject $ReturnedData  -membertype noteproperty -name "Current Threads in Use (Semaphore Holders)" -value $CookedSH
	add-member -inputobject $ReturnedData  -membertype noteproperty -name "Clients Currently Waiting (Semaphore Waiters)" -value $CookedSW
	add-member -inputobject $ReturnedData  -membertype noteproperty -name "Cumulative Client Timeouts (Semaphore Timeouts) " -value $CookedST
	add-member -inputobject $ReturnedData  -membertype noteproperty -name "Cumulative MaxConcurrentApi Thread Uses (Semaphore Acquires)" -value $CookedSA
	add-member -inputobject $ReturnedData  -membertype noteproperty -name "Duration of Calls (Avg Semaphore Hold Time)" -value $CookedASHT
	return $ReturnedData
}

function GetNetlogonInstances ([string]$RemoteComputerName = "localhost")
	{
	 #This function takes a computer name as input (default to local computer) and returns
	 #the instances-analagous to secure channels-a computer has. 
	 #Format returned is \\hostname.domainname.com.
	 
	 if ($RemoteComputerName -eq $null)
	 	{
		$LocalNetlogon = New-Object System.Diagnostics.PerformanceCounterCategory("Netlogon",$RemoteComputerName)
		$LocalInstances = $LocalNetlogon.GetInstanceNames()
	    $AllLocalInstances = @()
		foreach ($LocalInstance in $LocalInstances)
			{	
	 		 if ($LocalInstance -ne "_total")
	 			{
				 $AllLocalInstances += $LocalInstance
				}		 
			} 
			if ($AllLocalInstances -eq $null)
				{
				WriteTo-StdOut "The local computer was missing its DC perf instance so getting DC name from WMI." -shortformat

				$Query = "select * from win32_ntdomain where description = '" + $env:userdomain + "'"
				$v2 = get-wmiobject -query $Query
				$DCName = $v2.DomainControllerName
				$AllLocalInstances += $DCName
				WriteTo-StdOut "DCName is $AllLocalInstances" -shortformat
				}
		return $AllLocalInstances
		}
		else
		{
		$RemoteNetlogon = New-Object System.Diagnostics.PerformanceCounterCategory("Netlogon",$RemoteComputerName)
		$RemoteInstances = $RemoteNetlogon.GetInstanceNames()
		$AllRemoteInstances = @()
		foreach ($RemoteInstance in $RemoteInstances)
			{	
	 		 if ($RemoteInstance -ne "_Total")
	 			{
				 $AllRemoteInstances += $RemoteInstance
				}
			}
			if ($AllRemoteInstances -eq $null)
				{
				#If the local computer was missing its DC perf instance so getting DC name from WMI.
				$Query = "select * from win32_ntdomain where description = '" + $env:userdomain + "'"
				$v2 = get-wmiobject -query $Query
				$DCName = $v2.DomainControllerName
				$AllRemoteInstances += $DCName
				}
		return $AllRemoteInstances
		}
	}


if (($CheckMaxConcurrentApi) -and ($Instance -ne "_Total") -and ($Computer -ne "Localhost") -and ($CalcMCA -eq $true))
	{CheckMaxConcurrentApi -instancename $Instance -ComputerName $Computer -Calc $CalcMCA  | FL }
elseif (($CheckMaxConcurrentApi) -and ($Instance -ne "_Total") -and ($Computer -ne "Localhost"))
	{CheckMaxConcurrentApi -instancename $Instance -ComputerName $Computer | FL }
elseif (($CheckMaxConcurrentApi) -and ($Instance -ne "_Total"))
	{CheckMaxConcurrentApi -instancename $Instance  | FL}
elseif (($CheckMaxConcurrentApi) -and ($Computer -ne "Localhost"))
	{CheckMaxConcurrentApi -ComputerName $Computer  | FL}
elseif (($CheckMaxConcurrentApi) -and ($CalcMCA -eq $true))
	{CheckMaxConcurrentApi -Calc $calcmca  | FL}
elseif ($CheckMaxConcurrentApi)
	{CheckMaxConcurrentApi | FL}
if (($GetNetlogonInstances) -and ($Computer -ne "Localhost"))
	{GetNetlogonInstances  | FL}
elseif ($GetNetlogonInstances) 
	{GetNetlogonInstances -RemoteComputerName $Computer | FL}
