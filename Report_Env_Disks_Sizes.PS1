<#
    This scripts reads the availabled / used disk-space from the Exchange-Servers as well as the Database sizes
    all values are in megabyte if not described other
    Out-files contains only the disk capacity / used capacity from one Server per DAG. This is sufficcient since all Databases should have the same size

#>


# input parameters
$threshold_information = 50    # value in percent -> reflects % of used disk-space
$threshold_warning = 70    # value in percent -> reflects % of used disk-space

# don't change below this line
#----------------------------------------------------
$timestamp = get-date -Format "dd.MM.yyyy HH:mm"
$date = get-date -Format "yyyyMMdd"
$outpath = "\\Server\e$\ExchAdm\Export\Capacity-reporting_Disk-size"
[array]$alldags = $null
$h = hostname
$scriptname = $MyInvocation.MyCommand.Definition

Set-ADServerSettings -ViewEntireForest $true

New-Item -name "$date" -Path $outpath -ItemType directory -ErrorAction SilentlyContinue

$DAGs = get-databaseavailabilitygroup EU*
foreach ($DAG in ($DAGs | sort-object name)) {
    "reading DAG: " + $dag.name
    $dag_object = new-object PsObject
    
    # variables / counters for diskspace
    $dag_out = ""
    $dag_databasesize = 0
    $dag_databaseavailablenewmailboxspace = 0
    $server_disksize=0
    $server_freespace=0
    $server_used = 0
    $dagquota = 0
    $dagname = $dag.name

    $databases = get-mailboxdatabase -server $dag.Servers[0].name -status | sort-object name # check only databases from one server per DAG since all servers should have the same size

    [array]$databases_object = $null
    foreach ($database in $databases) {
        $DBObject = New-Object PSObject
        $databasename = $database.name 
        write-host "Read Database: $databasename ..." -NoNewline    
        $databasesize = $database.databasesize.tomb()
        $databaseAvailableNewMailboxspace = $database.AvailableNewMailboxSpace.tomb()
        $dag_databasesize+=$databasesize
        $dag_databaseavailablenewmailboxspace+=$databaseAvailableNewMailboxspace
        $mailboxes = get-mailbox -Database $databasename -resultsize unlimited # read mailboxes to get the quota value. This is required to get an estimation how much disk-space is "planded" for each mailbox
        $quota = 0
        write-host "read Mailboxes..." -NoNewline
        $ErrorActionPreference="silentlycontinue"
        $mailboxcount_Database = $mailboxes.count
        foreach ($mailbox in $mailboxes) {
            $quota+=$mailbox.prohibitsendquota.value.tomb()
            $quota+=$mailbox.ArchiveQuota.value.tomb()
        }
        $dagquota+=$quota
        $ErrorActionPreference="continue"
        write-host "done"
    
        $DBObject | add-member Noteproperty "DB-NAME" $databasename
        $DBObject | add-member Noteproperty "DB-SIZE (GB)"$([System.Math]::Round($databasesize/1024, 1))
        $DBObject | add-member Noteproperty "WHITESPACE (GB)"$([System.Math]::Round($databaseavailablenewmailboxspace/1024, 1)) 
        $DBObject | add-member Noteproperty "QUOTA (GB)"$([System.Math]::Round($quota/1024, 1))  
        $databases_object+=$DBObject
    }

    # checking disk-size only on one server of DAG since Databases should have more-or less the same size on all hosts
    $disks = get-WmiObject win32_logicaldisk -Computername $dag.Servers[0].name | where {$_.deviceid -like "G:" -or $_.deviceid -like "H:" -or $_.deviceid -like "I:"}
    [array]$lunobject = $null
    foreach ($disk in $disks) {
        $DiskObject = New-Object PSObject
        $server_disksize+=$disk.size/1024/1024
        $server_freespace+= $disk.freespace/1024/1024
        $server_used = $server_disksize-$server_freespace
        $DiskObject | add-member Noteproperty "DeviceID" $disk.deviceid
        $DiskObject | add-member Noteproperty "LUN-SIZE (GB)" $([System.Math]::Round($disk.size/1024/1024/1024, 1)) 
        $DiskObject | add-member Noteproperty "LUN-Free (GB)" $([System.Math]::Round($disk.freespace/1024/1024/1024, 1)) 
        $DiskObject | add-member Noteproperty "LUN-Used (GB)" $([System.Math]::Round(($disk.size-$disk.freespace)/1024/1024/1024, 1))
        $lunobject+=$diskobject
    }
    write-host "Creating objects"
    $dag_Object | add-member Noteproperty DAG-NAME $dagname
    $dag_Object | add-member Noteproperty "TOTAL-SIZE (GB)" $([System.Math]::Round($server_disksize/1024, 1)) 
    $dag_Object | add-member Noteproperty "USED (GB)" $([System.Math]::Round($server_used/1024, 1))  
    $dag_Object | add-member Noteproperty "FREE (GB)" $([System.Math]::Round($Server_freespace/1024, 1))  
    $dag_Object | add-member Noteproperty "DATABASE-SIZES (GB)" $([System.Math]::Round($dag_databasesize/1024, 1))  
    $dag_Object | add-member Noteproperty "WHITESPACE (GB)" $([System.Math]::Round($dag_databaseavailablenewmailboxspace/1024, 1))  
    $dag_Object | add-member Noteproperty "Quota (GB)" $([System.Math]::Round($dagquota/1024, 1))
    $dag_Object | add-member Noteproperty "DATABASEDETAIL" $databases_object
    $dag_Object | add-member Noteproperty "LUNDETAIL" $lunobject
    $alldags+=$dag_object 
    

}


# create Detailed Report 
$html_out=""
$null  | out-file "$outpath\$date\Disk-Size_detail_$Date.csv" 
"`"DAG-NAME`",`"TOTAL-SIZE (GB)`",`"USED (GB)`",`"FREE (GB)`",`"DATABASE-SIZES (GB)`",`"WHITESPACE (GB)`",`"Quota (GB)`"" | out-file "$outpath\$date\Disk-Size_detail_$Date.csv" -Append


$html_out+='<head><meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"/><title>Storage Capacity Overview (Detailled)</title><style type="text/css" media="screen,projection">body{font-size:16px;}div{padding: 1em .5em;}p{margin:.5em;}a{color:#000000;}a span{display: none;}a:focus span,a:active span{display: block;border:0 none;padding:1em;text-decoration: none;}</style></head><body><div>'
$html_out+="<a href='Disk-Size.html'>Go to Simple-Report</a><br>Created: $timestamp by $h, $scriptname<br><br>"

$html_out+="<table border='1'>"
$html_out+="<tr><th>Name</th><th>Total-Disk-Size (GB)</th><th>Disk-Size Used (GB)</th><th>Disk-Size Free (GB)</th><th>% Used</th><th>Database(s)-Size (GB)</th><th>Whitespace (GB)</th><th>Quota (GB)</th></tr>"
foreach ($entry in $alldags) {    
    $freepercent = [System.Math]::Round($entry.'USED (GB)' * 100 / $entry.'TOTAL-SIZE (GB)',1)
    $html_out+="<tr bgcolor = `"#D3D3D3`"><td>" + $entry.'DAG-name' + "</td><td>" + $entry.'TOTAL-SIZE (GB)'+ "</td><td>" + $entry.'USED (GB)'+ "</td><td>" + $entry.'FREE (GB)'+ "</td><td>" + $freepercent + "</td><td>" + $entry.'DATABASE-SIZES (GB)'+ "</td><td>" + $entry.'WHITESPACE (GB)'+ "</td><td>" + $entry.'Quota (GB)'+ "</td></tr>"
    "`"" + $entry.'DAG-name' + "`",`"" + $entry.'TOTAL-SIZE (GB)'+ "`",`"" + $entry.'USED (GB)'+ "`",`"" + $entry.'FREE (GB)'+ "`",`"" + $freepercent + "`",`"" + $entry.'DATABASE-SIZES (GB)'+ "`",`"" + $entry.'WHITESPACE (GB)'+ "`",`"" + $entry.'Quota (GB)'+ "`"" | out-file "$outpath\$date\Disk-Size_detail_$Date.csv" -append

    $rowcount = 0   # counter for each item (Database) in $entry.databasedetail
    $luncount = 0   # counter for each item (LUN) in $entry.lundetail
    $entry.databasedetail | foreach {
        if ($rowcount%6 -eq 0) {  # every 6 databases are on one LUN
            $freepercent_lun = [System.Math]::Round(($entry.lundetail)[$luncount].'LUN-Used (GB)' * 100 / ($entry.lundetail)[$luncount].'LUN-SIZE (GB)',1)
            $html_out+="<tr><td>--" + ($entry.lundetail)[$luncount].'Deviceid'+"</td><td>"+ ($entry.lundetail)[$luncount].'LUN-SIZE (GB)'+"</td><td>"+ ($entry.lundetail)[$luncount].'LUN-Used (GB)'+"</td><td>"+ ($entry.lundetail)[$luncount].'LUN-FREE (GB)'+"</td>"
            if ($freepercent_lun -lt $threshold_information) { $html_out+="<td bgcolor=`"green`">$Freepercent_LUN</td>" }
            elseif ($freepercent_lun -gt $threshold_warning) { $html_out+="<td bgcolor=`"red`">$Freepercent_LUN</td>" }
            else { $html_out+="<td bgcolor=`"yellow`">$Freepercent_LUN</td>"}
            $html_out+="<td>-</td><td>-</td><td>-</td></tr>"
            $luncount+=1
        }
        $html_out+="<tr><td>----" + $_.'DB-Name'+"</td><td>-</td><td>-</td><td>-</td><td>-</td><td>"+$_.'DB-SIZE (GB)'+"</td><td>"+$_.'Whitespace (GB)'+"</td><td>"+$_.'Quota (GB)'+"</td></tr>"
        "`"--" + $_.'DB-Name'+"`",`"`",`"`",`"`",`"`",`""+$_.'DB-SIZE (GB)'+"`",`""+$_.'Whitespace (GB)'+"`",`""+$_.'Quota (GB)'+"`"" | out-file "$outpath\$date\Disk-Size_detail_$Date.csv" -Append
        $rowcount+=1
        
    }
}
$html_out+="</table>"
$html_out | out-file "$outpath\Disk-Size_detail.html"

# create Standard Report
$html_out=""
$html_out+='<head><meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"/><title>Storage Capacity Overview</title><style type="text/css" media="screen,projection">body{font-size:16px;}div{padding: 1em .5em;}p{margin:.5em;}a{color:#000000;}a span{display: none;}a:focus span,a:active span{display: block;border:0 none;padding:1em;text-decoration: none;}</style></head><body><div>'
$html_out+="<a href='Disk-Size_detail.html'>Go to Detail-Report</a><br>Created: $timestamp by $h, $scriptname<br><br>"

$html_out+="<table border='1'>"
$html_out+="<tr><td>Name</td><td>Total-Disk-Size (GB)</td><td>Disk-Size Used (GB)</td><td>Disk-Size Free (GB)</td><td>% Used</td><td>Databases-Size (GB)</td><td>Whitespace (GB)</td><td>Quota (GB)</td></tr>"
foreach ($entry in $alldags) {    
    $freepercent = [System.Math]::Round($entry.'USED (GB)' * 100 / $entry.'TOTAL-SIZE (GB)',1)
    $html_out+="<tr><td>" + $entry.'DAG-name' + "</td><td>" + $entry.'TOTAL-SIZE (GB)'+ "</td><td>" + $entry.'USED (GB)'+ "</td><td>" + $entry.'FREE (GB)'+ "</td><td>" + $freepercent + "</td><td>" + $entry.'DATABASE-SIZES (GB)'+ "</td><td>" + $entry.'WHITESPACE (GB)'+ "</td><td>" + $entry.'Quota (GB)'+ "</td></tr>"
}
$html_out+="</table>"

$html_out | out-file "$outpath\Disk-Size.html"

$alldags | select-object 'DAG-NAME', 'TOTAL-SIZE (GB)', 'USED (GB)', 'FREE (GB)', 'DATABASE-SIZES (GB)', 'WHITESPACE (GB)', 'Quota (GB)' | ConvertTo-Csv -NoTypeInformation | out-file "$outpath\$date\Disk-Size_$Date.csv"



# create Report for all disks

$servers = get-mailboxserver | where {$_.name -like "*ash*"} | sort-object name
[array]$alldisks = $null
foreach ($server in $servers) {
    $servername = $server.name

    $disks = Get-WmiObject Win32_Volume -ComputerName $servername| select-object Name, Label, FreeSpace, capacity | sort-object name
    
    foreach ($disk in $disks) {
       $DiskObject = New-Object PSObject
       $DiskObject | add-member Noteproperty "Servername" $servername
       $DiskObject | add-member Noteproperty "DeviceID" $disk.name
       $DiskObject | add-member Noteproperty "Name" $disk.Label
       $DiskObject | add-member Noteproperty "Free Space" $([System.Math]::Round($disk.freespace/1024/1024/1024, 1)) 
       $DiskObject | add-member Noteproperty "Capacity" $([System.Math]::Round($disk.capacity/1024/1024/1024, 1))
       $DiskObject | add-member Noteproperty "% used" $([System.Math]::Round(($disk.capacity-$disk.freespace)*100/$disk.capacity, 1))
   
       $alldisks+=$DiskObject
    }
}

$alldisks | ConvertTo-Csv -NoTypeInformation | out-file "$outpath\$date\Disk-Size_Server_$Date.csv"
$alldisks | ConvertTo-Html | out-file "$outpath\Disk-Size_Server.html"

