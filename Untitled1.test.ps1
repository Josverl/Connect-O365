
Register-PSRepository -Name DevRepo -SourceLocation \\nas\DevRepo -InstallationPolicy Trusted

Uninstall-Script Connect-O365

get-installedmodule | Uninstall-Module

find-Script Connect-O365 -Repository DevRepo

Install-Script Connect-O365 -Repository DevRepo -Verbose 
