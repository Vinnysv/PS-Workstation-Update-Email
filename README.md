# PS-Workstation-Update-Email
This was written over one year ago. The method of sending mail and enumerating users has been replaced with the graph API.

# Device Update Reminder Script

This PowerShell script automates the process of sending custom email reminders to users, notifying them about the need for a Windows(or other software) update on their devices. It is particularly useful for organizations that need to ensure their devices are always up-to-date and secure, reducing the risk of security breaches and improving overall system performance.

## Features

- Retrieves device information from a vulnerability report spreadsheet.
- Enumerates users from the organization's Global Address List (GAL) in Microsoft Outlook.
- Filters devices that are already updated or have missing information.
- Sends personalized emails to users, reminding them to update their devices or contact IT support for assistance.

## Use Case Scenarios

- **IT Support and Maintenance**: Proactively notify users about pending updates, reducing the workload on IT support teams and minimizing the risk of security vulnerabilities.
- **Large Organizations**: Ensure devices are updated and secure across the organization, especially when managing a large number of devices and users.
- **Automated Device Management**: Integrate this script into an existing device management system to automate the update reminder process and keep track of device updates more efficiently.

## Prerequisites

- The script requires the `ImportExcel` PowerShell module. Install it using the following command:
Install-Module ImportExcel -Scope CurrentUser
- A vulnerability report spreadsheet containing device information (Asset Tag, Type, Windows Updated, Assigned User) for each user in the organization.
- Microsoft Outlook configured with access to the organization's Global Address List (GAL).

## Usage

1. Customize the `$domain` variable in the script with your organization's email domain.
2. Execute the script in PowerShell.
3. Provide the path to the vulnerability report spreadsheet when prompted.
4. Enter the site name (worksheet name in the spreadsheet) to send the email reminders.
5. The script will iterate through the devices, sending personalized update reminder emails to the respective users.
6. When the process is completed, you will be prompted with a 'Completed' message.
