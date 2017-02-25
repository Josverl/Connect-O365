BUG : UCC error 

PS C:\Users\josverl\onedrive\powershell> Connect-O365.ps1  -Account admin@atticware.onmicrosoft.com -SharePoint -AAD -UCC
UserName admin@atticware.onmicrosoft.com
Azure AD
SharePoint Online
Import-PSSession : The attribute cannot be added because variable Credential with value  would no longer be valid.
At C:\Program Files\WindowsPowerShell\Scripts\Connect-O365.ps1:556 char:17
+ ...             Import-PSSession $PSCompliance -AllowClobber -Verbose:$fa ...
+                 ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    + CategoryInfo          : MetadataError: (:) [Import-PSSession], ValidationMetadataException
    + FullyQualifiedErrorId : ValidateSetFailure,Microsoft.PowerShell.Commands.ImportPSSessionCommand



BUG : Auto restart bij installatie van de 'Sign in Assitstant' ??
	Microsoft Online Services Sign-In Assistant for IT Professionals Version : 7.250.4556.0  is already installed
	Daar moet op z'n minst een waarschuwing voor.


Feature: Support more Proxy configurations for Remote powershell
FEATURE - allow passing additional parameters to the different services 
ie timeout , proxy configuration 
 $PSBoundParameters


BUG: $tenantname not available/filled in all cases (-UCC) ( AAD Not loaded )  

Feature: Install module on demand 
	Ask to install a module if it is used , rather than thowing an error 

Feature: Install per Module , rather than all modules \
Refactor: Break up into Script + Depending module(s) 
	
Feature: Disconnect-O365 cmdlet

Feature : install.ps1
NUGET PowerShell Scripts
When a package carries an install.ps1 file within its \tools folder, 
the script will be run after package installation. An uninstall.ps1 is executed before uninstallation. Lastly, init.ps1 is executed every time the solution is opened (assuming the NuGet PowerShell Console is open). Target framework filters apply to this folder too.

BUG: cannot use account withouth saved password  - DONE 
	> Must use -persist 
	Cause: error in logic 

BUG: fix broken dependency on pnppowershell - DONE 

Feature: Add -Credential prameter - done  


