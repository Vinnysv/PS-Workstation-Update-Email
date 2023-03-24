# PS-Workstation-Update-Email
This was written over one year ago. The method of sending mail has been replaced with the graph API becoming mandatory. 

Device Update Reminder Script
This PowerShell script automates the process of sending custom email reminders to users, notifying them about the need for a Windows update on their devices. It is particularly useful for organizations that need to ensure their devices are always up-to-date and secure, reducing the risk of security breaches and improving overall system performance.

Features
Retrieves device information from a vulnerability report spreadsheet.
Enumerates users from the organization's Global Address List (GAL) in Microsoft Outlook.
Filters devices that are already updated or have missing information.
Sends personalized emails to users, reminding them to update their devices or contact IT support for assistance.
Use Case Scenarios
IT Support and Maintenance: Proactively notify users about pending updates, reducing the workload on IT support teams and minimizing the risk of security vulnerabilities.
Large Organizations: Ensure devices are updated and secure across the organization, especially when managing a large number of devices and users.
Automated Device Management: Integrate this script into an existing device management system to automate the update reminder process and keep track of device updates more efficiently.
Prerequisites
The script requires the ImportExcel PowerShell module. Install it using the following command:

powershell
Copy code
Install-Module ImportExcel -Scope CurrentUser
A vulnerability report spreadsheet containing device information (Asset Tag, Type, Windows Updated, Assigned User) for each user in the organization.

Microsoft Outlook configured with access to the organization's Global Address List (GAL).

Usage
Customize the $domain variable in the script with your organization's email domain.
Execute the script in PowerShell.
Provide the path to the vulnerability report spreadsheet when prompted.
Enter the site name (worksheet name in the spreadsheet) to send the email reminders.
The script will iterate through the devices, sending personalized update reminder emails to the respective users.
When the process is completed, you will be prompted with a 'Completed' message.
By implementing this script, organizations can improve the efficiency of their device management process and reduce the time and effort spent on manually reminding users to update their devices.
