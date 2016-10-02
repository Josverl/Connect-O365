$here = Split-Path -Parent $MyInvocation.MyCommand.Path
#$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.Tests\.', '.'
#cd C:\Users\josverl\OneDrive\PowerShell\Dev\Connect-O365


#Having the Module installed from PSGallery breaks the Mocking, or the returned values from the mocks, possibly caused by Signing
#Get-InstalledModule -Name 'ConnectO365' -ErrorAction SilentlyContinue | Uninstall-Module -Force -Verbose


Import-Module (Join-path $here ConnectO365) -Force -DisableNameChecking
Import-Module credentialmanager 

#test Credentials 
$Tester = 'CredentialStoreTester@verlinde.us'
$TestPass = "pass@word1"| ConvertTo-SecureString -AsPlainText -Force
$TestCred = New-Object System.Management.Automation.PsCredential($Tester,$TestPass)
$TestPass2 = "pass@word2"| ConvertTo-SecureString -AsPlainText -Force
$TestCred2 = New-Object System.Management.Automation.PsCredential($Tester,$TestPass2)

Describe "$here\ConnectO365 Module" {


    Context " File credentials" { 

        It "1. defines all the functions" {
            Test-Path Function:\Get-myCreds| Should Be $true
            Test-Path Function:\Store-myCreds| Should Be $true
            Test-Path Function:\Test-myCreds| Should Be $true
            Test-Path Function:\retrieve-credentials| Should Be $true
        }
        
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

    }

    Context "Stored Credentials" {
        It "4. can retrieve an existing account from CredMan -Persist | persist=false " {
            $x = New-StoredCredential -Comment "TEST Connect-O365" -Persist ENTERPRISE -Target $Tester -Type GENERIC -Credentials $TestCred 
            $r = retrieve-credentials -Account $Tester -Persist:$false 
            $r.UserName | Should Be $Tester
            
        }

        #do not use script scoped vars in mock
        Mock -ModuleName ConnectO365 Get-Credential { 
            $user = 'CredentialStoreTester1@verlinde.us'
            $Pass = "pass@word1"| ConvertTo-SecureString -AsPlainText -Force
            $Cred = New-Object System.Management.Automation.PsCredential($user,$Pass)
            return $Cred
        }

        It '5. can retrieve  an existing account from CredMan , asks for input and returns a valid cred | persist=true'  {
            
            #remove & Create it just to be sure
            Remove-Item "$env:USERPROFILE\creds\$Tester.txt" -Force -ErrorAction SilentlyContinue
            $x = New-StoredCredential -Comment "TEST Connect-O365" -Persist ENTERPRISE -Target $Tester -Type GENERIC -Credentials $TestCred 

            $r = retrieve-credentials -Account $Tester -Persist:$true
            $r| Should not Be $null
            $r.username | Should be 'CredentialStoreTester1@verlinde.us'
            $r.GetNetworkCredential().Password | should be 'pass@word1'
            Assert-MockCalled -ModuleName ConnectO365 Get-Credential -Exactly 1 -Scope It 
        }

        It '6. can retrieve an non-stored account, asks for input and returns a valid cred | persist:$false' {

            #remove & Create it just to be sure
            Remove-Item "$env:USERPROFILE\creds\$Tester.txt" -Force -ErrorAction SilentlyContinue
            Remove-StoredCredential -Target $Tester -ErrorAction SilentlyContinue
            #$x = New-StoredCredential -Comment "TEST Connect-O365" -Persist ENTERPRISE -Target $Tester -Type GENERIC -Credentials $TestCred 

            $r = retrieve-credentials -Account $Tester -Persist:$false
            $r| Should not Be $null
            $r.username | Should be 'CredentialStoreTester1@verlinde.us'
            $r.GetNetworkCredential().Password | should be 'pass@word1'
            Assert-MockCalled Get-Credential -Exactly 1 -Scope It -ModuleName ConnectO365
        }

        #do not use script scoped vars in mock
        Mock -ModuleName ConnectO365 Get-Credential { 
            return $null
        }

        It '7. can retrieve an non-stored account, asks for input and returns $null on cancel | persist:$false'  {
            Remove-StoredCredential -Target $Tester -ErrorAction SilentlyContinue
            $r = retrieve-credentials -Account $Tester 
            $r| Should Be $null
            Assert-MockCalled Get-Credential -Exactly 1 -Scope It -ModuleName ConnectO365
        }

        #do not use script scoped vars in mock
        Mock -ModuleName ConnectO365 Get-Credential { 
            $user = 'CredentialStoreTester1@verlinde.us'
            $Pass = "pass@word1"| ConvertTo-SecureString -AsPlainText -Force
            $Cred = New-Object System.Management.Automation.PsCredential($user,$Pass)
            return $Cred
        }

         It '8. can Save retrieve an non-stored account, asks for input and returns a valid cred | persist:$true'  {
            Remove-StoredCredential -Target $Tester -ErrorAction SilentlyContinue
            $r = retrieve-credentials -Account $Tester -Persist
            $r| Should not Be $null
            $r.UserName | Should Be 'CredentialStoreTester1@verlinde.us'
            $r.GetNetworkCredential().Password | should be 'pass@word1'
            Assert-MockCalled Get-Credential -Exactly 1 -Scope It -ModuleName ConnectO365
        }

        #do not use script scoped vars in mock
        Mock -ModuleName ConnectO365 Get-Credential { 
            $user = 'CredentialStoreTester2@verlinde.us'
            $Pass = 'pass@word2'| ConvertTo-SecureString -AsPlainText -Force
            $Cred = New-Object System.Management.Automation.PsCredential($user,$Pass)
            return $Cred
        }

        It '9. can Save retrieve an non-stored account, asks for input and returns a Different valid cred | persist:$true'  {
           

            Remove-StoredCredential -Target $Tester -ErrorAction SilentlyContinue
            $r = retrieve-credentials -Account $Tester -Persist
            $r| Should not Be $null
            $r.UserName | Should Be 'CredentialStoreTester2@verlinde.us'
            $r.GetNetworkCredential().Password | should be 'pass@word2'
            Assert-MockCalled Get-Credential -Exactly 1 -Scope It -ModuleName ConnectO365
        }

        #Clean Up 
        Remove-StoredCredential -Target $Tester -ErrorAction SilentlyContinue

    }
}