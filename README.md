# Send-Communication

Send targetable and repeatable email communication

# Background

Working as an IT professional it is necessary to be able to send targetable and repeatable communication in the event of maintenance or security events within any environment. These types of communications are often predetermined and require review before being sent. I developed this solution as a means of creating communications which could be prepared, tested, reviewed prior to being sent to their target audience.

# Use Cases:

- Bulk Password Change Notifications
- Introductions to new platforms & provide starting credentials
- Release Notes
- Maintenance Windows
- Status Messages

# Getting Started

1. Download the entirety of the code to a location your computer
1. Make a copy of the Example_Communication folder in the Communications folder
1. Modify the CommunicationTemplate.html file to create the desired message contents
1. Modify the subject.txt file to include the subject line of the email you wish to send
1. Run `$Credentials = get-credentials` and enter your Office 365 Account Username and Password
1. Run `.\Send-Communcation.ps1 -CommunicationTemplatePath .\Communications\MyCommunication -FromEmailAddress "ITSystems@fingerhut.us" -TestEmailAddress "recepient@test.com" -ContainsAttachments -Credentials $Credentials`

# Known Issues

- Currently uses Send-MailMessage to Microsoft Office 365 using a user account without MFA. Send-MailMessage has been marked as obsolete and will need to be replaced, my goal is to replace this with MailKit so more secure multi-factor secured accounts can be used and services other than Microsoft Office 365.

# Issues

Please report any issues or need for clarification and I will do what I can to address.

# Support Me

I write these solutions for my own learning opportunities, often in my free time. If this solution has helped you in some way, please support me so I can do more.
<a href="https://www.buymeacoffee.com/FingerhutAsCode"><img src="https://img.buymeacoffee.com/button-api/?text=Fund my caffeine intake&emoji=🥤&slug=FingerhutAsCode&button_colour=FFDD00&font_colour=000000&font_family=Cookie&outline_colour=000000&coffee_colour=ffffff" /></a>
