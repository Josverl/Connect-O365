# Connect-O365
   Connect to Office 365 and most related services and get ready to admin with Powershell.

I created this script because I found it too confusing to remember the different options to connect, as well as the different  installation methods for the different workloads, to prepare a workstation to connect to and administer office 365 services.

What initially started as a few lines, has grown over time to a script that is quite usefull, especially in combination with the ability of powershell 5 ( and earlier versions) to easily add and update scripts via the Powershell gallery.

The main use cases are :  

* Connect to multiple Office 365 services without the need to re-enter your credentials multiple times. 
* Close speccific connections ( -Close) 
* Install all relevant components and modules ( -Install)
* Tests the installed Modules ( -Test) 


How to Install
--------------
Run the below commands from an Admin elevated Powershell :

`Install-Script -Name Connect-O365`

To download and install the supporting modules :

`Connect-O365 -Install`


