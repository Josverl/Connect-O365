$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.Tests\.', '.'
Import-Module .\Connect-O365.ps1 

Describe ".\Connect-O365" {
    Context "credentials" { 

        It "can store an account "  -Pending {
            $true | Should Be $false
        }

        It "can retrieve an account"  -Pending {
            $true | Should Be $false
        }

    }
}
