# PowerShell Auto Emailer

A PowerShell script designed to be ran as a scheduled task for automating the process of sending emails on a consistent basis. This script was specifically written for the purpose of sending emails to a support inbox that the RMM platform NinjaOne auto-creates into tickets. These emails/tickets are sent on a weekly basis to remind technicians to verify backup services are running. 

With some customization you can use this script to send virtually any email to anyone. Over time I will be working on this script, making it more customizable, with more features. As of right now the script is designed for use in MY specific enviroment, you will need to spend some time editing the code to fit your exact enviroment and get the script working. Particularly with storing/obtaining credentials needed for authorization in SQL and the Microsoft Graph API.

## Azure & Microsoft Graph API

This script uses the Microsoft Graph API to send emails. To get access to this you must first register an application in Azure and give it the correct permissions.

If you need assistance setting up your application please refer to the following article...
``https://learn.microsoft.com/en-us/azure/active-directory-b2c/microsoft-graph-get-started?tabs=app-reg-ga``

## Installation / Setup

Step 1. ``git clone https://github.com/marsacom/NinjaOneToolKit.git``

Step 2. ``Install-Module -Name dbatools``

Step 3. ``Modify variables as needed in verifybackups.ps1``

## Usage

Can either be run straight from CLI using ``.\verifybackups.ps1`` or set up as a scheduled task to run at a scheduled time...

## Notes

***This script is currently in early stages of development and is being actively updated.*** 
***Getting this script working out of the box for your own needs may take some work as it was designed originally for one specific application, in one specific enviroment.***

Author : Brayden Kukla - 2024
