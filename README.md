#Connect-O365
-------------
Connect to Office 365 and most related services and get ready to admin with Powershell.

I created this script because I found it too confusing to remember the different options to install  and connect connect to the different Office 365 workloads. In my line of work this is also something I frequently need to explain to others, so they can repeat this on other workstations.
Also I found it tedious to need to enter the credentials for the different test accounts over and over, as well as the inability to use this from other scripts withouth re-writing nearly the same code over and over.

What initially started as a few lines, has grown over time to a script that is quite usefull, especially in combination with the ability of PowerShell 5 (and earlier versions) to easily add and update scripts via the Powershell gallery.

The main use cases are :  

* Connect to multiple Office 365 services without the need to re-enter your credentials multiple times. 
* Close all or specific connections ( -Close ) 
* Install all relevant components and modules ( -Install )
* Tests the installed Modules ( -Test ) 

##How to Use
--------------
Connect to AAD 
>*`Connect-O365 -Account admin@contoso.com`*

Note that the script support autocomplete for the -Account parameter using [Tab] or [Ctrl-Space] to choose from any of the persisted credentials on disk or in the Credential Manager.

Connect to SharePoint and Exchange online ( AAD automaticaly included)  
>`Connect-O365 -Account admin@contoso.com -SPO - EXO` 

##How to Install
----------------
Run the below commands from an Admin elevated Powershell :
>`Install-Script -Name Connect-O365`

or install for a single user only :
>`Install-Script -Name Connect-O365 -scope CurrentUser`

To download and install the supporting modules :
>`Connect-O365 -Install`

The install option will determine the modules to install. The exact modules to install are stored in a psd1 document that is stored on Github. 
Currently this contains the following modules :
* Microsoft Online Services Sign-In Assistant for IT Professionals
* Windows Azure Active Directory Module for Windows PowerShell
* Skype for Business Online, Windows PowerShell Module
* SharePoint Online Management Shell
* Azure Rights Management Administration Tool

This document contains a list of the moduldes to install, the required versions, the download locations and any specific installation options that are needed for a smooth installation. The required powershell module files are downloaded to the systems download folder. Before installation the digital certificate of all downloaded files is verified to make sure the installers are signed by Microsoft, before the installation of each module is started.

Dependend Modules that are published on the PowerShell Gallery are not subject to certificate verification.
* OfficeDevPnP.PowerShell, Module OfficeDevPnP.PowerShell.V16.Commands
* CredentialManager

##How to Update
---------------
Periodically update this and other scripts by running 
`update-script`

##Credential Management 
---------------
Credentials can be used from 2 locations: 
* A folder in the userprofile ($env:userProfile\Creds)
* Generic credentials stored in the Windows credential manager    

Store or update the credentials for admin@acontoso.com
>`Connect-O365 -Account admin@contoso.com -persist`
 
Use the credential stored in the credential manager using the 
>`Connect-O365 -Account Production`

Both credential stores can be udated (new password) by specifying the -Persist  parameter
 
New credentails that are created are stored in the userprofile folder.
Ro create credentials in the windows credential manager use : 
Control Panel > Credential Manager > Windows Credentials > Generic Credential > Add a Generic Credential

The network address can be used as an Alias

####*Sample :*

```
Internet or network address : Production
Username                    : serviceadmin@contoso.com
Password                    : pass@word1
```
When looking up the credentials from the credential manager matches can be made either on 
* The Username (`Connect-O365 -Account serviceadmin@contoso.com`)
* The target network address (`Connect-O365 -Account Production`)



