# Connect-O365
   Connect to Office 365 and most related services and get ready to admin with Powershell.

I created this script because I found it too confusing to remember the different options to connect, as well as the different  installation methods for the different workloads, to prepare a workstation to connect to and administer office 365 services.

What initially started as a few lines, has grown over time to a script that is quite usefull, especially in combination with the ability of powershell 5 ( and earlier versions) to easily add and update scripts via the Powershell gallery.

The main use cases are :  

* Connect to multiple Office 365 services without the need to re-enter your credentials multiple times. 
* Close all or specific connections ( -Close ) 
* Install all relevant components and modules ( -Install )
* Tests the installed Modules ( -Test ) 

How to Use
--------------
Connect to AAD 
`Connect-O365 -Account admin@contoso.com`
Note that there ia an autocomplete option using [Tab] or [Ctrl-Space] to choose from any of the persisted credentials

Connect to SharePoint and Exchange online ( AAD automaticaly included)  
`Connect-O365 -Account admin@contoso.com -SPO - EXO` 

How to Install
--------------
Run the below commands from an Admin elevated Powershell :
`Install-Script -Name Connect-O365`
To download and install the supporting modules :
`Connect-O365 -Install`

The install option will determine the modules to install. the exact modules to install on the source download locations on m download.microsoft.com are contained 
a xml formatted PSD1 document also stored on this github location.

The required powershell module  files are downloaded to the systems download folder , the digital certificate in checked to make sure the installers are signed by Mmicrosoft, before the installation 
of each module is started.


Credential Management 
---------------------
Credentials can be used from 2 locations 
* A folder in the userprofile ($env:userProfile\Creds)
* Generic credentials stored in the Windows credential manager    

Store or update the credentials for admin@acontoso.com
`Connect-O365 -Account admin@contoso.com -persist`
 
Use the credential stored in the credential manager using the 
`Connect-O365 -Account Production`

Both credential stores can be udated (new password) by specifying the -Persist  parameter
 
New credentails that are created are stored in the userprofile folder.
Ro create credentials in the windows credential manager use : 
Control Panel > Credential Manager > Windows Credentials > Generic Credential > Add a Generic Credential
the network address can be used as an Alias 
Sample 
`Internet or network address : Production`
`Username                    : serviceadmin@contoso.com`
`Password                    : pass@word1`


