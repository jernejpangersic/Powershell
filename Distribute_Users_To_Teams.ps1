ImportModule MicrosoftTeams

$credentials = Get-Credential

$excel = Import-Excel -Path "ExistingUsers.xlsx"

Connect-MicrosoftTeams -Credential $credentials

foreach($item in $excel) {

    $index = [array]::IndexOf($excel, $item)
    $percent = (($index/$excel.Count)*100)
    Write-Progress -Activity "Processing users $($index) of $($excel.Count)" -Status "Progress: $percent%" -PercentComplete $percent;
    
    Write-Output "Adding $($item.UPN) to team $($item.Team)"
    Get-Team -DisplayName $item.Team | Add-TeamUser -User $item.UPN
}