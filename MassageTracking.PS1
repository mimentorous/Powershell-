$messageID="700464CF28836543ABF16F6A74480F6A84D9651"
$sender=""
$recipient=""
$start="11.06.2018 01:00:13"
$end="14.06.2018   23:59:26"

$s=get-transportserver  |where {$_.name -ilike 'Nai*' -or $_.name -ilike 'SW*'}| foreach {Get-MessageTrackingLog  -server $_    -messageid $messageid  -start  $start -end $end -resultsize unlimited } 

$s | select-Object  Timestamp,  eventid,source,*Sender, @{Name=’recipients‘;Expression={[string]::join(“;”, ($_.recipients))}},messagesubject,*byte* | Sort-Object Timestamp | FT -a




$Addresses = Gc C:\temp\export_msgtracking_logs.txt

Foreach ($address in $addresses) {

$sender=$Address;
$recipient=$address;
$start="01.05.2017 02:00:13";
$end="30.06.2017   19:59:26";
"starting to collcet Sent mails logs for address "
$sender

$export_sent = "c:\temp\" + $sender + "_sent.txt"
$export_rec = "c:\temp\" + $sender + "_receive.txt"


$sent  = get-transportserver  |where {$_.name -ilike 'Nai*' -or $_.name -ilike 'SW*'}| foreach {Get-MessageTrackingLog -server $_  -sender $sender  -start $start -end $end -resultsize unlimited } ;
$sent | select-Object  Timestamp,  eventid,source,*Sender, @{Name=’recipients‘;Expression={[string]::join(“;”, ($_.recipients))}},messagesubject,*size*,*byte* | Sort-Object Timestamp | where { $_.eventid -eq "DELIVER"} | export-csv $export_sent  -notypeinformation ;
$sent.count
$rec  = get-transportserver  |where {$_.name -ilike 'Nai*' -or $_.name -ilike 'SW*'}| foreach {Get-MessageTrackingLog -server $_  -recipient $recipient  -start $start -end $end -resultsize unlimited } ;
$rec | select-Object  Timestamp,  eventid,source,*Sender, @{Name=’recipients‘;Expression={[string]::join(“;”, ($_.recipients))}},messagesubject,*byte* | Sort-Object Timestamp | where { $_.eventid -eq "DELIVER"} | export-csv $export_rec  -notypeinformation ;
$rec.count

}
