# This script lets you create an user and add them to a team

# Function to remove special characters
function RmvSpc
{
    PARAM ([string]$String)
    [Text.Encoding]::ASCII.GetString([Text.Encoding]::GetEncoding("Cyrillic").GetBytes($String))
}

$domain = "@mydomain.com"

$firstName = Read-Host -Prompt "First name"
$lastName = Read-Host -Prompt "Last name"

$firstName = RmvSpc(($firstName).Trim()).Replace(" ", "-")
$lastName = RmvSpc(($lastName).Trim()).Replace(" ", "-")
$uname = $firstName + "." + $lastName + $domain
$staff = Read-Host -Prompt "Staff"
$team = Read-Host -Prompt "Team"

# Based on user input, we define the license for users - Faculty or Student. A1 license will be assigned
# Important - replace XXXXX with your tenant name (XXXX.onmicrosoft.com)
if($staff) {
    $LicenseAssignment = "XXXXX:STANDARDWOFFPACK_FACULTY"
} else {
    $LicenseAssignment = "XXXXX:STANDARDWOFFPACK_STUDENT"
}

# Fill with username and password
$UserName = "username"
$PWord = ConvertTo-SecureString -String "password" -AsPlainText -Force
$credentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $UserName, $PWord

# Connect to MSOL Service and Microsoft Teams
Connect-MsolService -Credential $credentials
Connect-MicrosoftTeams -Credential $credentials

# Parameters for new user
$params = @{
    DisplayName = (Get-Culture).TextInfo.ToTitleCase($firstName) + " " + (Get-Culture).TextInfo.ToTitleCase($lastName)
    UserPrincipalName = $uname
    UsageLocation = "US"
    PreferredLanguage  = "en-US"  # Preferred language for users - https://docs.microsoft.com/en-us/previous-versions/commerce-server/ee825488(v=cs.20)?redirectedfrom=MSDN
    ForceChangePassword = $false
    FirstName = (Get-Culture).TextInfo.ToTitleCase($firstName.ToLower())
    LastName = (Get-Culture).TextInfo.ToTitleCase($lastName.ToLower())
}

# Create new user based on parameters, and add license
New-MsolUser @params -LicenseAssignment $LicenseAssignment
Write-Output "Created $($firstName) $($lastName) with username $($uname)"

# Wait for 1 second, to make sure the user is created
Start-Sleep -s 1

$user = Get-MsolUser -UserPrincipalName $uname

# Find a team, based on input and add user to it
Get-Team -DisplayName $team | Add-TeamUser -User $user.ObjectId
Write-Output "User $($user.UserPrincipalName) added to $($team)"