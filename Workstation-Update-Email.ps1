# This script sends a custom email to users on a vulnerability report spreadsheet, requesting contact to update a device.
# The script requires the ImportExcel module, which can be installed with the following command:
# Install-Module ImportExcel -Scope CurrentUser

# Install ImportExcel module for current user
Install-Module ImportExcel -Scope CurrentUser

# Prompt user for the path to the vulnerability report spreadsheet
$path = Read-Host -Prompt 'Path to spreadsheet'

# Initialize Outlook application
$outlook = new-object -comobject outlook.application

# Function to enumerate Global Address List (GAL)
function enumerate-GAL {
    [Microsoft.Office.Interop.Outlook.Application] $outlook = New-Object -ComObject Outlook.Application
    $entries = $outlook.Session.GetGlobalAddressList().AddressEntries
    foreach ($entry in $entries) {
        $entry2 = $entry.getExchangeUser()
        write-output $entry2
    }
}

# Get the users from the Global Address List
$users = enumerate-GAL 

# Prompt user for the site to send the emails to
$site = Read-Host -Prompt 'Site to send'

# Import device information from the specified worksheet
$devices = Import-Excel -Path $path -WorkSheetname $site
$tags = $devices."Asset Tag"
$types = $devices."Type"
$updates = $devices."Windows Updated"
$devices = $devices."Assigned User"

# Set the domain for email addresses
$domain = "@domain.org"

# Iterate through the devices and send the email reminders
foreach ($device in $devices) {
    if ([string]::IsNullOrEmpty($device)) {
        continue
    }
    if (!($users | where { $_.name -like $device })) {
        Write-Host $device "not found in address book. Check Name"
        continue
    }
    $tag = $tags[$devices.IndexOf($device)]
    $type = $types[$devices.IndexOf($device)]
    $updated = $updates[$devices.IndexOf($device)]
    
    # Check if necessary information is available
    if ([string]::IsNullOrEmpty($tag)) {
        Write-Host $device "Tag empty"
        continue
    }
    if ([string]::IsNullOrEmpty($type)) {
        Write-Host $device "Type empty"
        continue
    }
    if (-not [string]::IsNullOrEmpty($updated)) {
        Write-Host $device "Listed as updated"
        continue
    }
    if ($type -eq "Not assigned") {
        Write-Host $device "Type listed as Not assigned"
        continue
    }
    
    # Prepare email content
    $type = $type.replace("Tower - ", "").replace("AIO - ", "")
    $device = $device.replace(" ", ".").split(":")[0]
    $first = $device.split(".")[0]
    $first = $first.substring(0, 1).toupper() + $first.substring(1)
    Write-Host $first ", Success! Email sent to" $device "with tag" $tag
    
    # Create and send the email
    $email = $outlook.CreateItem(0)
    $email.To = $device + $domain
    $email.Subject = $first + ", Your " + $type + " Is In Need Of An Update"
    $email.Body = @"
Hey $first,

Your $type with the property tag $tag has been identified by the I.T department as in need of a urgent Windows update. Please reply to this email to schedule an update with I.T. If you do not have
Thanks,

Vincent Spagnola
Job Title
Cell: Phone Number

This email is automated.
"@
$email.Send()
}

Prompt user when the process is completed
$input = Read-Host -Prompt 'Completed'
