Set-Mailbox Email@addres -GrantSendOnBehalfTo ((Get-Mailbox EFM-PSO-Accounting.Team@allianz.com).GrantSendOnBehalfTo + "vvvvvvvv")

Set-DistributionGroup "Email@addres" -AcceptMessagesOnlyFrom ((Get-DistributionGroup -identity "Email@addres").AcceptMessagesOnlyFrom + "Email@addres")

Set-PublicFolder "\test"  -Replicas  ((PublicFolder "\test").Replicas  + "Database_Identity")
