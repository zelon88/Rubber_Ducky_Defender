NAME: Rubber_Ducky_Defender

TYPE: VBS Script

PRIMARY LANGUAGE: VBScript
 
AUTHOR: Justin Grimes

ORIGINAL VERSION DATE: 1/3/2020

CURRENT VERSION DATE: 1/7/2020

VERSION: v1.2


DESCRIPTION: An application to detect the use of "Bad USB" devices, such as a HAK5 "Rubber Ducky" payload deployment device.





PURPOSE: To detect and disable potential "Bad USB" devices before they can affect or compromise company infrastructure.




INSTALLATION INSTRUCTIONS: 
1. Install Rubber_Ducky_Defender into a subdirectory of your Network-wide scripts folder.
2. Open Rubber_Ducky_Defender.vbs with a text editor and configure the variables at the start of the script to match your environment.
3. Open sendmail.ini with a text editor and configure your email server settings.
4. Run the script automatically on domain workstations at machine startup or user logon with a GPO. Or both!


USAGE / ARGUMENTS / SWITCHES:
 No Email Notifications:     Rubber_Ducky_Defender.vbs -ne
                             Rubber_Ducky_Defender.vbs --nemail
 No Logging:                 Rubber_Ducky_Defender.vbs -nl
                             Rubber_Ducky_Defender.vbs --nlog
 No UI (Message Boxes):      Rubber_Ducky_Defender.vbs -ng
                             Rubber_Ducky_Defender.vbs --ngui
 

NOTES: 
1. "Fake Sendmail for Windows" is required for this application to send notification emails. Per the "Fake Sendmail" license, the required binaries are provided. To reinstall "Fake Sendmail for Windows" please visit  https://www.glob.com.au/sendmail/
2. Use absolute UNC paths for network addresses. DO NOT run this from a network drive letter.
