$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.Tests\.', '.'

Import-Module (Join-path $here 'installmodules') -Force -DisableNameChecking


Describe ".\Connect-O365" {
    Context "External Modules on which we depend"  {
       #load the required modules from a configuration file on GitHub
        $Components = Import-DataFile -url 'https://raw.githubusercontent.com/Josverl/Connect-O365/master/RequiredModuleInfo.psd1' 


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
                    $temp = New-TemporaryFile
                    { Invoke-WebRequest $c.Source -OutFile $temp.FullName}| should not throw 
                    #Downloaded Filelength should be GT 0 
                    $temp.Length -gt 0 | should be $true
                    $temp | Remove-Item
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
