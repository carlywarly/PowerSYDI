# PowerSYDI

## SYNOPSIS

Extracts information about the operating system, applications installed and hardware for a windows system and creating a html or XML file.
 
## DESCRIPTION

Extracts information about the Windows server or workstation, for example operating system, applications installed and hardware etc and generates a report in html file or an XML file. The files generated will have the same name as the server in question and the corresponding extension for the relevant output type.

## PARAMETERS

| Name | Type | Details | Default |
| ---- | ------- | ---| ------- |
| Output | String | What format should the output be generated in, either html or XML. | html |
| Computername | String  | The name of the computer you wish to run against, IP addresses can also be user. | To the local computer |
| LoadDoc | Switch | Display the output once the script has completed. |False |
| NoPing | Switch | Do not ping servers first, useful for servers that have the Windows Firewall enabled. | False|
| Background | String | Colour of the header, link bar and footer. | Blue |
| Path | String | Path to where you wish to save the output. | ".\\"  current folder |
| username | String | Username the script should run under, such as "User01", "Domain01\User01", or User@domain.com. If you have not used -password, you will be prompted for a password. | |
| password | String | Used in conjunction with the above -username | | 
| Credential | Credential Object | Powershell credential object you wish to run the script under, cannot be used with above username and password ||
| reportSoftware<BR>reportHardware<BR>reportStorage<BR>reportNetwork<BR>reportMisc<BR>reportExtra<BR> reportUser | Switches | Extract corresponding details and add to report. | True |

  
## EXAMPLES

    PowerSYDI.ps1

>Reports on current system and save to current folder, using your logged on credentials.

    PowerSYDI.ps1 -computer server2 -username domain\admininstrator -noping -Background "#ff0000"

> Runs the script against host server2, using the domain administrator account. You will be prompted for a password and the server will not be pinged. The output will be called server2.html
 
    PowerSYDI.ps1 -computer server3 -path d:\documents -LoadDoc -Background Blue

> Runs the script against host server3, saving the output, with a blue background to d:\documents and display the html webpage. The output will be called server3.html

    PowerSYDI.ps1 -computer server4 -LoadDoc -reportMisc:$$false -reportExtra:$$false

> Runs the script against host server4, display the output and don't report on miscellaneous or extra information. The output will be called server4.html

### LINK

[https://github.com/carlywarly/PowerSYDI](https://github.com/carlywarly/PowerSYDI)

##### ToDo

Win32_ServerFeature
