$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.Tests\.', '.'
#Import-Module .\Connect-O365.ps1 
#cd  C:\Users\josverl\OneDrive\PowerShell\Dev\Connect-O365
#. .\Connect-O365.ps1 


Describe 'Script connect-o365' {
    Context 'Scrip Metadata' {
        It 'should have a valid ScriptFileInfo' {
            { $Current = Test-ScriptFileInfo -Path connect-o365.ps1 } | Should not throw
        }
        $Current = Test-ScriptFileInfo -Path connect-o365.ps1
        It 'should have the correct GUID' {
            $Current.Guid | Should be 'a3515355-c4b6-4ab8-8fa4-2150bbb88c96'
        }

        It 'Current Version version should be Newer that Published version' {
            $Published = find-script -Name 'connect-o365' -Repository 'psgallery' 
            $Current.Version -GT $Published.Version | Should be $true
        }
    }
    
}
