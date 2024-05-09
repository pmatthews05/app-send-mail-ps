# app-send-mail-ps
Setup and code to send email using an App Registration and Microsoft Graph

# Set up
## App Registration
- Create an App Registration in Azure AD with Certificate
- Add the following API Permissions:
#  - Office 365 Exchange Online
#    - Exchange.ManageAsApp
  - Microsoft Graph
    - Mail.Send
    - Mail.ReadWrite
# - Grant Role Assignment to the App Registration
#  - Exchange Administrator Role
#  - Compliance Administrator Role

The following powershell script will create the App Registration for you. You will need to update the variables at the top of the script to match your environment.
[create-app-reg.ps1](create-app-reg.ps1)

## Create 