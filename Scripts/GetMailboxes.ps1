# configurable parameters
$outputTargetAccountsCSV = "C:\targets.csv"
$targetADGroup = "Domain Users"

# Import exchange server management cmdlets
. 'C:\Program Files\Microsoft\Exchange Server\V14\bin\RemoteExchange.ps1'
. 'C:\TEMP\Linagora\Join-Object.ps1'
Connect-ExchangeServer -auto

# Available exchange mailboxes
$excAccounts = Get-Mailbox -ResultSize Unlimited -Filter { RecipientTypeDetails -ne "DiscoveryMailbox" } | Select DistinguishedName,DisplayName,PrimarySmtpAddress, @{Name=“EmailAddresses”;Expression={$_.EmailAddresses | ? {$_.PrefixString -ceq “smtp”} | % {$_.SmtpAddress}}}
# List of Users from a given AD group
$adUsers = Get-ADGroupMember $targetADGroup | Select DistinguishedName
# Join users and mailboxes on DistinguishedName
$targetAccounts = Join-Object -Left $adUsers -Right $excAccounts  –LeftJoinProperty DistinguishedName –RightJoinProperty DistinguishedName –RightProperties PrimarySmtpAddress,EmailAddresses -Type OnlyIfInBoth
# Dump mailboxes in a CSV file
$targetAccounts | Export-Csv -NoTypeInformation $outputTargetAccountsCSV
