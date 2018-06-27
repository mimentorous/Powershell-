add-pssnapin Microsoft.Exchange.Management.Powershell.Admin  -ErrorAction silentlycontinue
$VersionScript = "2016-03-23ZB" 
Start-Transcript 
$VerbosePreference="Continue"
$MyInvocation
## ZB end headers


# WICHTIG!!! In diesem Block ÄNDERN ++++++++++++++++++++++++++++++++++++++++++++
$backup_dir = "C:\Users\user\Desktop\ADS\Backup" # string darf nicht mit "\" enden
$logpath = "C:\Users\user\Desktop\ADS\Backup" # string darf nicht mit "\" enden

#+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
# Ab hier Änderungen nur in Absprache mit Editor (siehe oben)!!!!!!!!!!!!      +
#+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


# Übergabe der Parameter vom ISIM
if ($args.count -eq 0 -or $args.count -gt 2)
{
	Write-Host "Zu viele/wenige User übergeben"
	exit 20
}
$user = $args[0]
$dc = $args[1]

Write-Host "Suche Mailbox..." -NoNewline
$mbx_to_convert = Get-Mailbox $user
$dn = $mbx_to_convert.distinguishedname

$date = get-date -Format yyyyMMdd.HHmmss
$lpath = "$logpath\"+$date+"_"+$dn+".log"
Start-Transcript -Path $lpath

if (!$mbx_to_convert) {Write-Host "Mailbox nicht gefunden"; exit 10}
Write-Host "Mailbox gefunden"

Write-Host "Work on Mailbox $user... sleep"
Start-Sleep -Seconds 10

# Suche Domaincontroller, falls keiner übergeben wurde
if (!($dc))
{
	Write-Host "Suche Domaincontroller"
	$ctx = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext([System.DirectoryServices.ActiveDirectory.DirectoryContextType]::Domain, "wwg00m")
	$SiteName = [System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite().Name
	$dc = [System.DirectoryServices.ActiveDirectory.DomainController]::FindOne($ctx, $SiteName).Name
	if (!($dc))
	{
		write-host "Kein Domaincontroller gefunden"
		exit 12
	}
	Write-Host "Domaincontroller gefunden: $dc"
}


# Übergabe an Variable, mit welcher innerhalb des Skriptes gearbeitet wird
$mbx = Get-Mailbox -Identity $mbx_to_convert -DomainController $dc


# Abbruch, wenn Mailbox keine Linked Mailbox ist
Write-Host "Prüfe ob Mailbox Linked ist" -NoNewline
if ($mbx.RecipientTypeDetails -notlike "LinkedMailbox")
{
	"Mailbox ist keine Linked Mailbox - Skript wird abgebrochen"
	exit 14 
}
Write-Host "...fertig"


# Abfrage ob schonmal ein logon auf der Mailbox war.
if (!(Get-MailboxStatistics $mbx))
{
	# initialer Mailboxzugriff
	Get-MailboxFolderStatistics $mbx | Out-Null
	Start-Sleep -Seconds 5
}

Write-Host "Mailboxstatistics vorhanden"

# Prüfen auf doppelte primäre SMTP-Adresse in allianzde
Write-Host "Prüfen doppelte SMTP-Adresse in allianzde"
$PrimarySmtpAddress = $mbx.PrimarySmtpAddress
$path = "LDAP://DC.NAME/DC=allianzde,DC=rootdom,DC=NET"
$Domain = $path
$root = New-Object system.DirectoryServices.DirectoryEntry $Domain
$objFinder = New-Object system.DirectoryServices.DirectorySearcher
$objFinder.SearchRoot = $root
$objFinder.pagesize = 90000
$objFinder.filter = "(mail=$PrimarySmtpAddress)"
$objFinder.PropertiesToLoad.Add("DistinguishedName") | Out-Null
$colResults = $objFinder.FindAll()
$colResults_count_allianzde = $colResults.count


# Prüfen auf doppelte primäre SMTP-Adresse in wwg00m
Write-Host "prüfe doppelte SMTP-Adresse in wwg00m"
$path = "LDAP://DC.NAME/DC=wwg00m,DC=rootdom,DC=NET"
$Domain = $path
$root = New-Object system.DirectoryServices.DirectoryEntry $Domain
$objFinder = New-Object system.DirectoryServices.DirectorySearcher
$objFinder.SearchRoot = $root
$objFinder.pagesize = 90000
$objFinder.filter = "(mail=$PrimarySmtpAddress)"
$objFinder.PropertiesToLoad.Add("DistinguishedName") | Out-Null
$colResults = $objFinder.FindAll()
$colResults_count_wwg00m = $colResults.count

if ($colResults_count_allianzde -gt 1 -or $colResults_count_wwg00m -gt 1)
{
	# Primäre Adresse des Users bereits vergeben. Zuerst Adresse wechseln
	Write-Host "Doppelte SMTP-Adresse gefunden. Skript wird abgebrochen"
	exit 16
}


# Aktivierung AD-Konto
Write-Host "Aktivierung AD-Konto" -NoNewline
$account = [adsi]"LDAP://$dc/$dn"
$account.psbase.invokeset("AccountDisabled", "False")
$account.setinfo()
Write-Host "...fertig"


# Auslesen der fest definierten Variableninhalte vom jeweiligen Benutzer
Write-Host "Beginne Auslesen der MBX-Konfiguration" -NoNewline
$homedb=$mbx.database
$legacydn=$mbx.LegacyExchangeDN
$samaccount=$mbx.samaccountname
$alias=$mbx.alias
$emailaddresses=$mbx.EmailAddresses
$language=$mbx.languages
$ExternalOoO=$mbx.ExternalOofOptions
$prohibitsendquota=$mbx.ProhibitSendQuota
$prohibitsendreceivequota=$mbx.ProhibitSendReceiveQuota
$issuewarningquota=$mbx.IssueWarningQuota
$useDBdefault=$mbx.UseDatabaseQuotaDefaults
$addresslistmember=$mbx.AddressListMembership
$PoliciesIncluded=$mbx.PoliciesIncluded
$identity=$mbx.identity
$PrimarySmtpAddress = $mbx.PrimarySmtpAddress
$CustomAttribute1 = $mbx.CustomAttribute1
$CustomAttribute2 = $mbx.CustomAttribute2
$CustomAttribute3 = $mbx.CustomAttribute3
$CustomAttribute4 = $mbx.CustomAttribute4
$CustomAttribute5 = $mbx.CustomAttribute5
$CustomAttribute6 = $mbx.CustomAttribute6
$CustomAttribute7 = $mbx.CustomAttribute7
$CustomAttribute8 = $mbx.CustomAttribute8
$CustomAttribute9 = $mbx.CustomAttribute9
$CustomAttribute10 = $mbx.CustomAttribute10
$CustomAttribute11 = $mbx.CustomAttribute11
$CustomAttribute12 = $mbx.CustomAttribute12
$CustomAttribute13 = $mbx.CustomAttribute13
$CustomAttribute14 = $mbx.CustomAttribute14
$CustomAttribute15 = $mbx.CustomAttribute15
$sendonbehalf = $mbx.grantsendonbehalfto
$HiddenFromGal = $mbx.HiddenFromAddressListsEnabled
$AltRecipient = $mbx.ForwardingAddress
$AltRecipientDeliver = $mbx.DeliverToMailboxAndForward
$lma = $mbx.linkedmasteraccount
$EndDateForRetentionHold = $mbx.EndDateForRetentionHold
$StartDateForRetentionHold = $mbx.StartDateForRetentionHold
$RetentionHoldEnabled = $mbx.RetentionHoldEnabled
$ManagedFolderMailboxPolicy = $mbx.ManagedFolderMailboxPolicy
$RulesQuota = $mbx.RulesQuota 
$EmailAddressPolicyEnabled = $mbx.EmailAddressPolicyEnabled 
$RecipientLimits = $mbx.RecipientLimits
$RequireSenderAuthenticationEnabled = $mbx.RequireSenderAuthenticationEnabled 
$OfflineAddressBook = $mbx.OfflineAddressBook 
$RetainDeletedItemsUntilBackup = $mbx.RetainDeletedItemsUntilBackup 
$MaxReceiveSize = $mbx.MaxReceiveSize 
$SimpleDisplayName = $mbx.SimpleDisplayName 
$RetainDeletedItemsFor = $mbx.RetainDeletedItemsFor
$RejectMessagesFrom = $mbx.RejectMessagesFrom 
$AcceptMessagesOnlyFrom = $mbx.AcceptMessagesOnlyFrom 
Write-Host "...fertig" 


# Auslesen CAS_MBX-Konfiguration
Write-Host "Beginne Auslesen der CAS_MBX-Konfiguration" -NoNewline
$Exch_casMbx = Get-CasMailbox -Identity $mbx -DomainController $dc -ErrorAction SilentlyContinue
$Exch_casMbx_ActSync = $Exch_casMbx.ActiveSyncEnabled
$Exch_casMbx_POP = $Exch_casMbx.PopEnabled
$Exch_casMbx_IMAP = $Exch_casMbx.ImapEnabled
$Exch_casMbx_MAPI = $Exch_casMbx.MAPIEnabled
Write-Host "...fertig"


# Auslesen AD-Attribute
Write-Host "Beginne Auslesen AD-Attribute" -NoNewline
$dn = $mbx.distinguishedname
$targetaddress = ([adsi]"LDAP://$dc/$dn").targetaddress
$textEncodedORAddress = ([adsi]"LDAP://$dc/$dn").textEncodedORAddress
$allianzOEdNSdomainName = ([adsi]"LDAP://$dc/$dn").'allianz-OEdNSdomainName'
$msExchmasteraccountsid = ([adsi]"LDAP://$dn").msExchmasteraccountsid
Write-Host "...fertig"


# Auslesen Proxyadressen
Write-Host "Beginne Auslesen Proxyadressen" -NoNewline
$emailaddresses | foreach {$proxy = $_.ProxyAddressString;  [Array]$proxyaddresses+=$proxy}
Write-Host "...fertig" 


# Auslesen RejectMessagesFrom
Write-Host "Beginne Auslesen RejectMessagesFrom" -NoNewline
$RejectMessagesFrom | foreach {$tmp_rej = $_.distinguishedname;  [Array]$reject+=$tmp_rej}
Write-Host "...fertig" 


# Auslesen AcceptMessagesOnlyFrom
Write-Host "Beginne Auslesen AcceptMessagesOnlyFrom" -NoNewline
$AcceptMessagesOnlyFrom | foreach {$temp_acc = $_.distinguishedname;  [Array]$accept+=$temp_acc}
Write-Host "...fertig" 


# Auslesen SendOnBehalf
Write-Host "Beginne Auslesen SendOnBehalf" -NoNewline
foreach ($granted in $mbx.GrantSendOnBehalfTo) 
{
	[Array]$SOB+=$granted.distinguishedname
}
Write-Host "...fertig" 


# Auslesen Addresslistmembership
Write-Host "Beginne Auslesen Addresslistmembership" -NoNewline
$addresslistmember | foreach {$temp_alm = $_.distinguishedname;  [Array]$almembership+=$temp_alm}
Write-Host "...fertig" 


# Fullmailboxaccess
# FullMailboxAccess bleibt erhalten


# SendAsPermission
# SendAsPermission bleibt erhalten


# Erstellung dump-File:
Write-Host "Beginne Erstellung Dump File" -NoNewline
$foldername = $dn+"_"+$date
New-Item -Path $backup_dir -Name $foldername -ItemType directory | out-null
$mbx | Select-Object * | out-file -FilePath "$backup_dir\$foldername\MBX_$dn.txt"
$Exch_casMbx | Select-Object * | out-file -FilePath "$backup_dir\$foldername\MBXCAS_$dn.txt"
$account = [adsi]"LDAP://$dc/$dn" | Select-Object * | out-file -FilePath "$backup_dir\$foldername\AD_$dn.txt"
Write-Host "...fertig" 


# Erstellung Input-File für Archiv-Aktivierung
$adquery = [adsi]"LDAP://$dc/$dn"

if ($adquery.memberof | where {$_ -like "*APPL-MARCH*"}) {
	
	$archive_output = "\\server\ISIM$\" + $mbx.distinguishedname + ".txt"
	$displayn = $mbx.displayname
	
	"ArchiveName = $displayn " | out-file $archive_output 
	"GrantAccess = read write delete, $lma" | out-file $archive_output -append

}


# Warte Bis AD-Konto Aktiv ist
$err_AD=1
do 
{ 
	Write-Host "Warte bis AD-Konto Aktiv ist" -NoNewline
	start-sleep -Seconds 2
	$uac = ([adsi]"LDAP://$dc/$dn").userAccountControl
	if($uac -eq "544" -or $uac -eq "512" -or $uac -eq "66048" -or $uac -eq "66080" -or $uac -eq "262656" -or $uac -eq "262688" -or $uac -eq "328192" -or $uac -eq "328224") 
	{$err_AD=0}
} 
while ($err_AD -eq 1)
Write-Host "...fertig" 
	

#Disconnecten der Mailbox
write-host "Disable Mailbox $PrimarySmtpAddress" -NoNewline
disable-mailbox -Identity $mbx -DomainController $dc -Confirm:$false
Write-Host "...fertig" 

		
# kurze Pause um Replizierungsprobleme zu vermeiden
write-host "Starte pause für Replizierung" -nonewline
Start-Sleep -s 10
write-host "...fertig"


#reconnect der Mailbox
write-host "Connect Mailbox $PrimarySmtpAddress "
$trycount = 1
write-host "Versuch $trycount"
$trycount++
Connect-Mailbox -Identity $legacydn -Database $homedb -User $identity -DomainController $dc -Confirm:$false 


# Warten bis Mailbox-Connect Repliziert ist und Mailbox verfügbar ist.
$err=0
do 
{ 
	if(get-mailbox $legacydn -DomainController $dc) {$err=1}
	else
	{
		write-host "Versuch $trycount"
		$trycount++	
		Connect-Mailbox -Identity $legacydn -Database $homedb -User $identity -DomainController $dc -Confirm:$false 
	}
	start-sleep -Seconds 2
	
	
	# Abbruchkriterium, dass nur max. 50x versucht wird die Mailbox zu reconnecten
	if ($trycount -gt 50)
	{ $err = 2  }
	
} 
while ($err -lt 1)

if ($err -eq 2)
{
	Write-Host "Mailbox konnte nicht reconnected werden"
	exit 18
}

Write-Host "...fertig" 


#setzen der Mailboxproperties
Write-Host "Beginne Setzen Mailboxproperties" -NoNewline
Set-Mailbox -Identity $legacydn -PrimarySmtpAddress $PrimarySmtpAddress -EmailAddressPolicyEnabled $false -DomainController $dc
Set-Mailbox -Identity $legacydn -Alias $alias -languages $language -ExternalOofOptions $ExternalOoO -prohibitsendquota $prohibitsendquota -prohibitsendreceivequota $prohibitsendreceivequota -issuewarningquota $issuewarningquota -usedatabasequotadefaults $useDBdefault -ForwardingAddress $AltRecipient -DeliverToMailboxAndForward $AltRecipientDeliver -HiddenFromAddressListsEnabled $HiddenFromGal -DomainController $dc
Set-Mailbox -Identity $legacydn -CustomAttribute1 $CustomAttribute1 -CustomAttribute10 $CustomAttribute10 -CustomAttribute11 $CustomAttribute11 -CustomAttribute12 $CustomAttribute12 -CustomAttribute13 $CustomAttribute13 -CustomAttribute14 $CustomAttribute14 -CustomAttribute15 $CustomAttribute15 -CustomAttribute2 $CustomAttribute2 -CustomAttribute3 $CustomAttribute3 -CustomAttribute4 $CustomAttribute4 -CustomAttribute5 $CustomAttribute5 -CustomAttribute6 $CustomAttribute6 -CustomAttribute7 $CustomAttribute7 -CustomAttribute8 $CustomAttribute8 -CustomAttribute9 $customattribute9  -DomainController $dc
Set-Mailbox -Identity $legacydn -EndDateForRetentionHold $EndDateForRetentionHold -StartDateForRetentionHold $StartDateForRetentionHold -RetentionHoldEnabled $RetentionHoldEnabled -ManagedFolderMailboxPolicy $ManagedFolderMailboxPolicy  -DomainController $dc
Set-Mailbox -Identity $legacydn -RulesQuota $RulesQuota -EmailAddressPolicyEnabled $EmailAddressPolicyEnabled -RecipientLimits $RecipientLimits -OfflineAddressBook $OfflineAddressBook  -RequireSenderAuthenticationEnabled $RequireSenderAuthenticationEnabled  -DomainController $dc
Set-Mailbox -Identity $legacydn -RetainDeletedItemsUntilBackup $RetainDeletedItemsUntilBackup -MaxReceiveSize $MaxReceiveSize -SimpleDisplayName $SimpleDisplayName -RetainDeletedItemsFor $RetainDeletedItemsFor  -DomainController $dc
Set-CASMailbox -Identity $legacydn -PopEnabled $Exch_casMbx_POP -ImapEnabled $Exch_casMbx_IMAP -ActiveSyncEnabled $Exch_casMbx_ActSync -MAPIEnabled $Exch_casMbx_MAPI -DomainController $dc
Write-Host "...fertig" 


# Setze SendOnBehalf
Write-Host "Setze SendOnBehalf" -NoNewline
foreach ($SendOnBehalfEntry in $SOB | where {$_.length -gt 1})
{
	$tmp_mbx = Get-Mailbox -Identity $legacydn -DomainController $dc
	$tmp_mbx.GrantSendOnBehalfTo.Add("$SendOnBehalfEntry")
	Set-Mailbox -Instance $tmp_mbx -DomainController $dc
}
Write-Host "...fertig" 


# Setzen der AD-Attribute
if ($targetaddress)
{
	Write-Host "Setze AD-Attribute - Targetaddress" -NoNewline
	$ad_user= [adsi]"LDAP://$dc/$dn"
	$ad_user.targetAddress="$targetaddress"
	$ad_user.textEncodedORAddress = "$textEncodedORAddress"
    $ad_user.msexchmasteraccountsid=$msExchmasteraccountsid
	$ad_user.Setinfo()
	Write-Host "...fertig" 
}
Write-Host "Setze AD-Attribute - textEncodedORAddress" -NoNewline
$ad_user= [adsi]"LDAP://$dc/$dn"
$ad_user.textEncodedORAddress = "$textEncodedORAddress"
$ad_user.Setinfo()
Write-Host "...fertig" 


# Setzen der Proxyadressen
Write-Host "Setze Proxyadressen" -NoNewline
foreach ($address in $proxyaddresses | where {$_.length -gt 1 -and $_ -ne "SMTP:$PrimarySmtpAddress"})
{
	$tmp_mbx = Get-Mailbox -Identity $legacydn -DomainController $dc
	$tmp_mbx.EmailAddresses.Add("$address")
	Set-Mailbox -Instance $tmp_mbx -DomainController $dc -ErrorAction SilentlyContinue
}
Write-Host "...fertig" 


# Lösche während des Reconnects entstandene Proxy-Adresse @local
Write-Host "Lösche @local - Proxyadresse" -NoNewline
foreach ($WrongAddress in $tmp_mbx.EmailAddresses | where {$_.SmtpAddress -like "*@local"})  ############## Änderung für Produktion!!!!!
{
	$tmp_mbx.EmailAddresses.Remove("$WrongAddress") | Out-Null
	Set-Mailbox -Instance $tmp_mbx -DomainController $dc
}
Write-Host "...fertig" 


# Setzen RejectmessagesFrom
Write-Host "Setze RejectmessagesFrom" -NoNewline
foreach ($RejectEntry in $reject | where {$_.length -gt 1})
{
	$tmp_mbx = Get-Mailbox -Identity $legacydn -DomainController $dc
	$tmp_mbx.RejectmessagesFrom.Add("$RejectEntry")
	Set-Mailbox -Instance $tmp_mbx -DomainController $dc 
}
Write-Host "...fertig" 


# Setzen AcceptMessagesOnlyFrom 
Write-Host "Setze AcceptMessagesOnlyFrom" -NoNewline
foreach ($acceptEntry in $accept | where {$_.length -gt 1})
{
	$tmp_mbx = Get-Mailbox -Identity $legacydn -DomainController $dc
	$tmp_mbx.AcceptMessagesOnlyFrom.Add("$acceptEntry")
	Set-Mailbox -Instance $tmp_mbx -DomainController $dc 
}
Write-Host "...fertig" 



# Entferne alte Berechtigungen
Write-Host "Entferne Berechtigungen für alten User" -NoNewline
if ($lma -like "wwg00m\*"){
	# determine correct linkedmasteraccount-name
	$lma_account = ($lma.split("\"))[1]
	$lma = "$allianzoedNSdomainname\$lma_account"
}
get-mailbox $legacydn -DomainController $dc | remove-mailboxpermission -user $lma -accessrights fullaccess, externalaccount -confirm:$false  -DomainController $dc
$adpermission = get-mailbox $legacydn -DomainController $dc  | Get-adPermission -user $lma -DomainController $dc
if ($adpermission)
{
	$adpermission | Remove-ADPermission -DomainController $dc -Confirm:$false
}
Write-Host "...fertig" 


# Deaktivierung AD-Konto
Write-Host "Deaktivierung AD-Konto" -NoNewline
$account = [adsi]"LDAP://$dc/$dn"
$account.psbase.invokeset("AccountDisabled", "True")
$account.setinfo()
Write-Host "...fertig" 


# Finale Prüfung ob Mailbox eine User-Mailbox ist
if ((get-mailbox $legacydn -DomainController $dc).RecipientTypeDetails -eq "Usermailbox")
{
	Write-Host "Mailbox ist vom Typ User"
	Stop-Transcript
	exit 0
}
else
{
	Write-Host "Umstellung fehlgeschlagen"
	Stop-Transcript
	exit 1
}

