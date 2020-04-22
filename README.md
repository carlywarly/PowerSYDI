# PowerSYDI
PowerSYDI repository

.SYNOPSIS
Extracts information about the operating system, applications installed and 
hardware for a windows system and creating a html or XML file.

.DESCRIPTION
Extracts information about the Windows server or workstation, for example 
operating system, applications installed and hardware etc and generates a report
in html file or an XML file.  The files generated will have the 
same name as the server in question and the corresponding extension for 
the relevant output type.  

.PARAMETER Output 
What format should the output be generated in, either html or XML. 
Default: html

.PARAMETER Computername
The name of the computer you wish to run against, IP addresses can also be
user. Defaults to the local computer.

.PARAMETER LoadDoc
Display the output once the script has completed. Default: False

.PARAMETER NoPing
Do not ping servers first, useful for servers that have the Windows Firewall 
enabled. Default: False

.PARAMETER Background
Colour of the header, link bar and footer. Default: Red

.PARAMETER Path
Path to where you wish to save the output. Default: ".\", current folder

.PARAMETER username 
Username the script should run under, such as "User01", "Domain01\User01", or 
User@domain.com.   If you have not used -password, you will be prompted for a 
password.

.PARAMETER password 
Used in conjunction with the above -username

.PARAMETER Credential
Powershell credential object you wish to run the script under, cannot be used 
with above username and password

.PARAMETER reportSoftware, reportHardware, reportStorage, reportNetwork, 
               reportMisc, reportExtra, ReportUser
Extract corresponding details and add to report. Default: True

.EXAMPLE 
PowerSYDI.ps1
Reports on current system and save to current folder, using your logged on credentials.

.EXAMPLE 
PowerSYDI.ps1 -computer server2 -username domain\admininstrator -noping -Background "#ff0000"
Runs the script against host server2, using the domain administrator account. You will be prompted for a password and the server will not be pinged. The output will be called server2.html

.EXAMPLE 
PowerSYDI.ps1 -computer server3 -path d:\documents -LoadDoc -Background Blue
Runs the script against host server3, saving the output, with a blue background to d:\documents and display the html webpage. The output will be called server3.html

.EXAMPLE 
PowerSYDI.ps1 -computer server4 -LoadDoc -reportMisc:$false -reportExtra:$false
Runs the script against host server4, display the output and don't report on miscellaneous or extra information. The output will be called server4.html

.LINK
http://www.powerforge.net

.ToDo
Win32_ServerFeature
