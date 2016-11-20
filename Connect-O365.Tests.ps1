$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.Tests\.', '.'
#Import-Module .\Connect-O365.ps1 
#cd  C:\Users\josverl\OneDrive\PowerShell\Dev\Connect-O365
#. .\Connect-O365.ps1 


Describe 'Script connect-o365' {
    Context 'Script Metadata' {
        It 'should have a valid ScriptFileInfo' {
            { $Current = Test-ScriptFileInfo -Path connect-o365.ps1 } | Should not throw
        }
        $Current = Test-ScriptFileInfo -Path connect-o365.ps1
        It 'should have the correct GUID' {
            $Current.Guid | Should be 'a3515355-c4b6-4ab8-8fa4-2150bbb88c96'
        }

        It 'Current Version version should be Newer that Published version'-skip {
            $Published = find-script -Name 'connect-o365' -Repository 'psgallery' 
            $Current.Version -GT $Published.Version | Should be $true
        }
    }

    context 'actual connection test ' {
        $account = 'admin@atticware.onmicrosoft.com'

        It 'should connect to AAD' {
            { .\Connect-O365.ps1 -Account $account -AAD                              } | Should not throw
        }
        It 'should connect to SPO' {
            { .\Connect-O365.ps1 -Account $account -SPO                              } | Should not throw
        }
        It 'should connect to Exchange' {
            { .\Connect-O365.ps1 -Account $account -EXO                              } | Should not throw
        }
        It 'should connect to Skype' {
            { .\Connect-O365.ps1 -Account $account -Skype                              } | Should not throw
        }
        It 'should connect to RMS' {
            { .\Connect-O365.ps1 -Account $account -RMS                              } | Should not throw
        }
        It 'should connect to UCC' {
            { .\Connect-O365.ps1 -Account $account -UCC                              } | Should not throw
        }


    }
    
}
