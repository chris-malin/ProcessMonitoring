# Introduction
Monitoring of remote machines' processes, with alert eMail/SMS and automated remote machine restart.


# Prerequisites
1. Dedicated machine to host the solution (MonitoringServer), I'm using Windows Server 2019, but Win10 would work, too.
2. MS Office with Excel installed and licensed for the user running the solution on MonitoringServer.
3. For SMS functionality: Twilio Account (free should be more than sufficient)

# Installation
1. Create a directory on your MonitoringServers' C drive named "Monitoring" and copy the provided files/folders there.
2. Update the Variables listed below with your values.
3. Run trustedhosts.ps1 on your MonitoringServer once, and everytime you add a new Machine/IP to listofmachines.txt, Accept All. 
4. Run ProcessStatus_eMail-and-Twilio-SMS.ps1 as Admin manually, or use ScheduledTask_toImport.xml to import scheduled task.

# Variables
## listofmachines.txt
Enter your list of machines, either as hostname or IP, each in new line.

## ProcessStatus_eMail-and-Twilio-SMS.ps1
### Main Variables
`$processname = "processname"` Update with the process you want to monitor, e.g. **chrome**, **explorer**.

`$adminpass = ConvertTo-SecureString “PaintextPassword” -AsPlainText -Force` Enter your password for the user running the script.

`$adminCred = New-Object System.Management.Automation.PSCredential (“Administrator”, $adminpass)` User running the script.

### eMail variables
`$PSEmailServer = "smtp.gmail.com"`

`$SmtpUser = "SmtpUser"`

`$smtpPassword = "eMailpassword"`  Note: for e.g. gmail you need to create an app password in your google account

`$MailTo = "emailaddress"`

`$MailFrom = 'emailaddress@gmail.com'`

### SMS variables
C:\Monitoring\ServerDown-Twilio_v1\ServerDown_v2_Twilio-SMS-Helper.ps1

`-AccountSid "AccountSid"` SID from your Twilio account.

`-authToken "authToken"` Token from your Twilio account.

`-fromNumber "+15555555555"` Twilio number you set up (~$1/month at time of writing, but twilio free account gives you $15 credit).

`-toNumber  "+15555555555"` Target number you want the alert sent to.

-message $SMSText


## ScheduledTask_toImport.xml
`<Author>YourMonitoringServer\Administrator</Author>` Update with your MonitoringServerName or Domain

`<UserId>S-1-5-21-UPDATEWITHYOURS</UserId>` --> run cmd and execute **wmic getname,sid** to get SID

`<Arguments>-ExecutionPolicy Bypass -File C:\Monitoring\ProcessStatus_eMail-and-Twilio-SMS.ps1</Arguments>` Ensure the path is correct.


# Tweaking
eMail, SMS and Restart can be commented out in the .ps1, in case you don't want and/or need it. See comments starting with #CommentOutIf..
