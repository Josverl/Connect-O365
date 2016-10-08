$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.Tests\.', '.'
#Import-Module .\Connect-O365.ps1 
#cd  C:\Users\josverl\OneDrive\PowerShell\Dev\Connect-O365
#. .\Connect-O365.ps1 

