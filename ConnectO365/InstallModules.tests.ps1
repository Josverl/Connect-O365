$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.Tests\.', '.'

#$sut = $sut -replace '.ps1', '.psd1'
#$ModuleName = (Split-Path $sut -Leaf).Split('.')[0]

$ModuleName = 'ConnectO365'
$sut = Join-path $here -ChildPath "$ModuleName.psd1"

#Having multiple modules loaded breaks the mocking, remove all earlier versions
get-module -Name connecto365 -all  | remove-module 

#Now load the fresh Module
Import-Module (Join-path $here $ModuleName) -Force -DisableNameChecking -Verbose


Describe ".\ConnectO365" {
    Context "External Modules on which we depend"  {
       #load the required modules from a configuration file on GitHub
        $Components = Import-DataFile -url 'https://raw.githubusercontent.com/Josverl/Connect-O365/master/RequiredModuleInfo.psd1' 

        It "the module defines all the public functions" {
            $MFT = Test-ModuleManifest -Path $sut 
            $mft.ExportedCmdlets.ContainsKey("Get-O365ModuleFile") | Should Be $true
            $mft.ExportedCmdlets.ContainsKey("Import-DataFile") | Should Be $true

            #Workaround - cmdlets do not 
            #$mft.ExportedFunctions.ContainsKey("Get-O365ModuleFile") | Should Be $true
            #$mft.ExportedFunctions.ContainsKey("Import-DataFile") | Should Be $true


        }

         It "can load the file from github" {
            $Components | Should not Be $null
            $Components.AdminComponents| Should not Be $null
        }
         It "component file has correct version" {
            $Components.Version -ge "1.6.3" | should be $true
        }
        It "it Has at least 6 admin components " {
            $Components.AdminComponents.Count -ge 6 | Should Be $true
        }

        It "all MSI and EXE sources can be downloaded" {
            foreach ($c in $Components.AdminComponents ) { 
                if (($c.type).ToUpper()  -in "EXE","MSI") {
                    Write-Host $c.Source
                    #$temp = New-TemporaryFile
                    #{ Invoke-WebRequest $c.Source -OutFile $temp.FullName -Method Head }| should not throw 
                    #Downloaded Filelength should be GT 0 
                    #$temp.Length -gt 0 | should be $true
                    #$temp | Remove-Item
                    $Download = Invoke-WebRequest -uri $c.Source -Method Head -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
                    $Download.StatusCode | should be 200
                    $Download.Headers["Content-Length"] -gt 0| should be $true
                     
                }
            }
        }

        It "all Modules can be found 1 time in the PS SPGallery" {
            foreach ($c in $Components.AdminComponents ) { 
                if (($c.type).ToUpper()  -in "MODULE") {
                    Write-Host $c.Module
                    $test = @(find-module -Name $c.Module -Repository PSGallery )
                    $test.Count | should be 1
                }
            }
        }        

    }
}
