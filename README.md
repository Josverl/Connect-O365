Connect-O365
==============

Connect to Office 365 and most related services and get ready to admin with Powershell.

I created this script because I found it (too) confusing to remember the different options to install and connect connect to the different Office 365 workloads. In my line of work this is also something I frequently need to explain to others, so they can repeat this on other workstations. Also I found it tedious to need to enter the credentials for the different test accounts over and over, as well as the inability to use this from other scripts withouth re-writing nearly the same code over and over.

What initially started as a few lines, has grown over time to a script that is quite usefull, especially in combination with the ability of PowerShell 5 (and earlier versions) to easily add and update scripts via the Powershell gallery.

The main use cases are :

-   Connect to multiple Office 365 services without the need to re-enter your credentials multiple times.

-   Close all or specific connections ( -Close )

-   Install all relevant components and modules ( -Install )

-   Tests the installed Modules ( -Test )

How to Use
--------------

**Connect to AAD **

&gt;`Connect-O365 -Account admin@contoso.com`

Note that the script support autocomplete for the -Account parameter using \[Tab\] or \[Ctrl-Space\] to choose from any of the persisted credentials on disk or in the Credential Manager.

**Connect to SharePoint and Exchange online ( AAD automaticaly included)

&gt;`Connect-O365 -Account admin@contoso.com -SPO - EXO`

How to Install
--------------

Install to the *machine* by installing from an Admin elevated Powershell :

&gt;`Install-Script -Name Connect-O365`

or install for the *current user* only : (No admin permissions required)

&gt;`Install-Script -Name Connect-O365 -scope CurrentUser`


Intellisense
------------

The script supports autocomplete for the -Account parameter using [Tab] or [Ctrl-Space]. This allows you to toggle between, or select from any of the persisted credentials on Disk and in the Credential Manager.

To download and install the supporting modules to connect to Office 365:
=======

Note that the script support autocomplete for the -Account parameter using [Tab] or [Ctrl-Space] to choose from any of the persisted credentials on disk or in the Credential Manager.

&gt;`Connect-O365 -Install`

The install option will determine the modules to install. The exact modules to install are stored in a psd1 document that is stored on Github. Currently this contains the following modules :

-   Microsoft Online Services Sign-In Assistant for IT Professionals

-   Windows Azure Active Directory Module for Windows PowerShell

-   Skype for Business Online, Windows PowerShell Module

-   SharePoint Online Management Shell

-   Azure Rights Management Administration Tool

This document contains a list of the moduldes to install, the required versions, the download locations and any specific installation options that are needed for a smooth installation. The required powershell module files are downloaded to the systems download folder. Before installation the digital certificate of all downloaded files is verified to make sure the installers are signed by Microsoft, before the installation of each module is started.

Dependend Modules that are published on the PowerShell Gallery are not subject to certificate verification.

-   OfficeDevPnP.PowerShell, Module OfficeDevPnP.PowerShell.V16.Commands

-   CredentialManager

How to Update
-----------------

Periodically update this and other scripts by running `update-script`

Credential Management
=====================

Connect-O365 allows you to store and re-use the accounts and passwords , so youcan focus  on the task at hand.
The credentials are defely stored in the Windows Credential Manager.

When looking up the credentials from the credential manager matches can be made either on
-   The Username (`Connect-O365 -Account <serviceadmin@contoso.com>`) [Default]  
    
-   A label you assign in hte Credential Manager  (`Connect-O365 -Account ProductionAdmin`)

-   The target network address (`Connect-O365 -Account https://consoso.sharepoint.com`)
    


Credentials can be used from 2 locations:

-   Generic credentials stored in the Windows credential manager
-   A folder in the userprofile ($env:userProfile)  [Will be Depricated]

# *Sample :* 


Store or update the credentials for admin@acontoso.com

&gt;`Connect-O365 -Account admin@contoso.com -persist`

Use the credential stored in the credential manager using the &gt;`Connect-O365 -Account Production`

Both credential stores can be udated with a new account or new password by specifying the -Persist parameter

New credentails that are created are stored in the Windows Credential Manager, as this method is prefered over the file based store.

To manually create credentials in the windows credential manager use :

Control Panel &gt; Credential Manager &gt; Windows Credentials &gt; Generic Credential &gt; Add a Generic Credential

The network address can be used as an Alias

# *Sample : Label*

    Internet or network address : Production
    Username                    : serviceadmin@contoso.com
    Password                    : pass@word1

# *Sample : sharepoint*

    Internet or network address : https://contoso.sharepoint.com
    Username                    : serviceadmin@contoso.com
    Password                    : pass@word1 

Note that this type of credential is also used by the PnP Powershell cmdlets to retrieve the credentials when connecting to SharePoint Online. 

Code of Conduct
===============

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
