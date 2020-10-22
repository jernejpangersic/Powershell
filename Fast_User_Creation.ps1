# This script lets you create an user and add them to a team

Import-Module MsOnline
Import-Module MicrosoftTeams

# Function to remove special characters
function RmvSpc
{
    PARAM ([string]$String)
    [Text.Encoding]::ASCII.GetString([Text.Encoding]::GetEncoding("Cyrillic").GetBytes($String))
}
$credentials = Get-Credential
$domain = "@mydomain.com"

$firstName = Read-Host -Prompt "First name"
$lastName = Read-Host -Prompt "Last name"

$firstName = RmvSpc(($firstName).Trim()).Replace(" ", "-") # Removing special characters and replacing space with -
$lastName = RmvSpc(($lastName).Trim()).Replace(" ", "-") # Removing special characters and replacing space with -
$uname = $firstName + "." + $lastName + $domain # Constructing username
$staff = Read-Host -Prompt "Staff" # Is the user staff of student? 1 = Staff, 0 = Student
$team = Read-Host -Prompt "Team" # Not required, but you can enter the name of the Team, where you want to add the new user

# Based on user input, we define the license for users - Faculty or Student. A1 license will be assigned
# Important - replace XXXXX with your tenant name (XXXX.onmicrosoft.com)
if($staff) {
    $LicenseAssignment = "XXXXX:STANDARDWOFFPACK_FACULTY"
} else {
    $LicenseAssignment = "XXXXX:STANDARDWOFFPACK_STUDENT"
}


# Connect to MSOL Service and Microsoft Teams
Connect-MsolService -Credential $credentials
Connect-MicrosoftTeams -Credential $credentials

# Parameters for new user
$params = @{
    FirstName = (Get-Culture).TextInfo.ToTitleCase($firstName.ToLower())
    LastName = (Get-Culture).TextInfo.ToTitleCase($lastName.ToLower())
    DisplayName = (Get-Culture).TextInfo.ToTitleCase($firstName) + " " + (Get-Culture).TextInfo.ToTitleCase($lastName)
    UserPrincipalName = $uname
    UsageLocation = "US" # Usage location - fill with two letter code of your country
    PreferredLanguage  = "en-US"  # Preferred language for users - https://docs.microsoft.com/en-us/previous-versions/commerce-server/ee825488(v=cs.20)?redirectedfrom=MSDN
    Password = "MyPassword1"    # In this case, the password will be hard-coded. Remove this line if you want to have it generated automatically
    ForceChangePassword = $false  # Set to true if you want to force the user to change password at first login
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