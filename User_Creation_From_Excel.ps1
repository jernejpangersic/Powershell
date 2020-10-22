## Configuration
# Function to remove special characters from string
function RmvSpc
{
    PARAM ([string]$String)
    [Text.Encoding]::ASCII.GetString([Text.Encoding]::GetEncoding("Cyrillic").GetBytes($String))
}

# Login information for administrator. Change the username and password below
$domain = "@mydomain.com"
$credentials = Get-Credential

## End configuration

# Input Excel file with data
$excel = Import-Excel -Path "Users.xlsx"

# Output Excel file - new accounts created
$output = "Output.xlsx"

Import-Module MicrosoftTeams
Import-Module MSOnline

# Connect to Office 365 Tenant
Connect-MsolService -Credential $credentials

foreach($item in $excel) {

    $FirstName = RmvSpc(($item.FirstName).Trim()).Replace(" ", "-")     # Using RmvSpc function to remove special characters and replace space with - (in case of two names)
    $LastName = RmvSpc(($item.LastName).Trim()).Replace(" ", "-")       # Using RmvSpc function to remove special characters and replace space with - (in case of two names)
    $logonName = $FirstName + "." + $LastName   # Building the logonName
    $UPN = $FirstName + "." + $LastName + $domain

    # Prepare loading bar
    $index = [array]::IndexOf($excel, $item)
    $percent = (($index/$excel.Count)*100)
    Write-Progress -Activity "Processing users $($index) of $($excel.Count)" -Status "Progress: $percent%" -PercentComplete $percent;
    
    # If user exists, create a different username (if user@domain.com exists, it tries user1@domain.com, user2@domain.com, ...)
    if(Get-MsolUser -UserPrincipalName $UPN -erroraction SilentlyContinue) {
        Write-Output "User with UPN $($UPN) exists"
        $i = 0
        while (Get-MsolUser -UserPrincipalName $UPN -ErrorAction SilentlyContinue) {
            $i++
            $UPN = $logonName + $i + $domain
            Write-Output "`t Trying $($UPN)"
	    }
    }

    # Putting together parameters for the new user
    $params = @{
        FirstName = $item.FirstName   # First name as taken from the Excel file
        LastName = $item.LastName    # Last name as taken from the Excel file
        DisplayName = $item.FirstName + " " + $item.LastName   # Building the DisplayName
        UserPrincipalName = $UPN    # Building the UserPrincipalName
        AlternateEmailAddresses  = $item.Email    # If user has an alternative email, we can add it here
        UsageLocation = "US"    # Usage location must be set 
        PreferredLanguage  = "en-US"    # Preferred language for users - https://docs.microsoft.com/en-us/previous-versions/commerce-server/ee825488(v=cs.20)?redirectedfrom=MSDN
        ForceChangePassword = $true     # Force password reset on first login
        LicenseAssignment = "TENANT:STANDARDWOFFPACK_STUDENT" # The format here is TENANT:LICENSE. STANDARDWOFFPACK_STUDENT for A1 Student, STANDARDWOFFPACK_Faculty for A1 Faculty. Replace TENANT with your tenant name (XXXX.onmicrosoft.com)
    }

    Write-Output "Creating user with UPN: $($UPN)"

    # Create new user based on the parameters and export FirstName, LastName, DisplayName, UserPrincipalName, Password and AlternateEmailAddress and export it to excel
    New-MsolUser @params | Select-Object -Property `
        FirstName, `
        LastName, `
        DisplayName, `
        UserPrincipalName, `
        Password, `
        @{Name=“AlternateEmailAddresses”;Expression={$_.AlternateEmailAddresses}}, `
        ObjectId | Export-Excel -Path $output -Append -AutoSize -TableName "TableName"

}