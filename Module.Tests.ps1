#$here = Split-Path -Parent $MyInvocation.MyCommand.Path
#$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.Tests\.', '.'
#cd  C:\Users\josverl\OneDrive\PowerShell\Dev\Connect-O365

Import-Module .\ConnectO365 -Force -Verbose
Import-Module credentialmanager

Describe ".\ConnectO365 Module" {
    Context "credentials" { 

        It "1. defines all the functions" {
            Test-Path Function:\Get-myCreds| Should Be $true
            Test-Path Function:\Store-myCreds| Should Be $true
            Test-Path Function:\Test-myCreds| Should Be $true
            Test-Path Function:\retrieve-credentials| Should Be $true
        }
        
        #test Credentials 

        $Tester = 'CredentialStoreTester@verlinde.us'
        $TestPass = "pass@word1"| ConvertTo-SecureString -AsPlainText -Force
        $TestCred = New-Object System.Management.Automation.PsCredential($Tester,$TestPass)
        $TestPass2 = "pass@word2"| ConvertTo-SecureString -AsPlainText -Force
        $TestCred2 = New-Object System.Management.Automation.PsCredential($Tester,$TestPass2)


        It "2. Test returns true for existing FILE accounts "  {
            #Create it just to be sure
            md "$env:USERPROFILE\creds" -ErrorAction SilentlyContinue
            "Test"  > "$env:USERPROFILE\creds\$Tester.txt" 
            Test-myCreds -Account $Tester | Should Be $true
        }

        It "3. Test returns false for non existing FILE accounts "  {
            #remove it just to be sure
            Remove-Item "$env:USERPROFILE\creds\$Tester.txt" -Force -ErrorAction SilentlyContinue
            Test-myCreds -Account $Tester | Should Be $false
        }

        It "4. can retrieve an existing account from CredMan -Persist | persist=false " {
            $x = New-StoredCredential -Comment "TEST Connect-O365" -Persist ENTERPRISE -Target $Tester -Type GENERIC -Credentials $TestCred 
            $r = retrieve-credentials -Account $Tester -Persist:$false 
            $r.UserName | Should Be $Tester
            
        }

        It '5. can retrieve  an existing account from CredMan , asks for input and returns a valid cred | persist=true' {
            Mock Get-Credential -Verifiable { return $TestCred } -ModuleName ConnectO365
            #remove & Create it just to be sure
            Remove-Item "$env:USERPROFILE\creds\$Tester.txt" -Force -ErrorAction SilentlyContinue
            $x = New-StoredCredential -Comment "TEST Connect-O365" -Persist ENTERPRISE -Target $Tester -Type GENERIC -Credentials $TestCred 

            $r = retrieve-credentials -Account $Tester -Persist:$true
            $r| Should not Be $null
            Assert-MockCalled Get-Credential -Exactly 1 -Scope It -ModuleName ConnectO365
        }


        It '6. can retrieve an non-stored account, asks for input and returns a valid cred | persist:$false' {
            Mock Get-Credential -Verifiable { return $TestCred } -ModuleName ConnectO365

            #remove & Create it just to be sure
            Remove-Item "$env:USERPROFILE\creds\$Tester.txt" -Force -ErrorAction SilentlyContinue
            Remove-StoredCredential -Target $Tester -ErrorAction SilentlyContinue
            #$x = New-StoredCredential -Comment "TEST Connect-O365" -Persist ENTERPRISE -Target $Tester -Type GENERIC -Credentials $TestCred 

            $r = retrieve-credentials -Account $Tester 
            $r| Should not Be $null
            Assert-MockCalled Get-Credential -Exactly 1 -Scope It -ModuleName ConnectO365
        }

        It '7. can retrieve an non-stored account, asks for input and returns $null on cancel | persist:$false'  {
            Mock Get-Credential { return $null} -ModuleName ConnectO365

            Remove-StoredCredential -Target $Tester -ErrorAction SilentlyContinue
            $r = retrieve-credentials -Account $Tester 
            $r| Should Be $null
            Assert-MockCalled Get-Credential -Exactly 1 -Scope It -ModuleName ConnectO365
        }


         It '8. can Save retrieve an non-stored account, asks for input and returns a valid cred | persist:$true'  {
            Mock Get-Credential { return $TestCred } -ModuleName ConnectO365

            Remove-StoredCredential -Target $Tester -ErrorAction SilentlyContinue
            $r = retrieve-credentials -Account $Tester -Persist
            $r.UserName | Should Be $Tester
            Assert-MockCalled Get-Credential -Exactly 1 -Scope It -ModuleName ConnectO365
        }

         It '9. can Save retrieve an non-stored account, asks for input and returns a Different valid cred | persist:$true'  {
            Mock Get-Credential { return $TestCred2 } -ModuleName ConnectO365

            Remove-StoredCredential -Target $Tester -ErrorAction SilentlyContinue
            $r = retrieve-credentials -Account $Tester -Persist
            $r.UserName | Should Be $TestCred2.UserName

            Assert-MockCalled Get-Credential -Exactly 1 -Scope It -ModuleName ConnectO365
        }

        #Clean Up 
        Remove-StoredCredential -Target $Tester -ErrorAction SilentlyContinue
    }

}