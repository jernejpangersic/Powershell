# Counts all users with domain
Import-Module MsOnline
$credentials = Get-Credential
Connect-MsolService -Credential $credentials

$domain = Read-Host -Prompt 'Count users with domain: '

Write-Host "Counting Users ..."
(Get-MsolUser -DomainName $domain -All).Count