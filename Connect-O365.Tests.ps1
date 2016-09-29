$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.Tests\.', '.'
#Import-Module .\Connect-O365.ps1 
#cd  C:\Users\josverl\OneDrive\PowerShell\Dev\Connect-O365
#. .\Connect-O365.ps1 

Describe ".\Connect-O365" {
    Context "External Modules on which we depend"  {

        <#
        .Synopsis
           import a .psd1 file from a url ,( Github) 
        .DESCRIPTION
           import a .psd1 file from a url 
           and perform a safe expansion using a number of predefined variables.
        #>
        function Import-DataFile
        {
            param (
                [Parameter(Mandatory)]
                [string] $Url
            )
            try
            {
                #setup variables to use during configuration expansion
                $CPU = $env:PROCESSOR_ARCHITECTURE
                switch ($env:PROCESSOR_ARCHITECTURE)
                {
                    'x86'   {$xcpu = 'x86' ; $bitness='32';}
                    'AMD64' {$xcpu = 'x64' ; $bitness='64'; }
                }
                $Filename = $URL.Split("/")[-1]
                try {   wget -Uri $URL -OutFile "$env:TEMP\$Filename" } 
                #failsafe if IE never been run 
                catch { wget -Uri $URL -OutFile "$env:TEMP\$Filename" -UseBasicParsing  } 

                $content = Get-Content -Path "$env:TEMP\$Filename" -Raw -ErrorAction Stop
                Remove-Item "$env:TEMP\$Filename" -Force
                $scriptBlock = [scriptblock]::Create($content)

                # This list of approved cmdlets and variables is what is used when you import a module manifest
                [string[]] $allowedCommands = @( 'ConvertFrom-Json', 'Join-Path', 'Write-Verbose', 'Write-Host' )
                #list of pedefined variables that can be used
                [string[]] $allowedVariables = @('language' ,'LangCountry', 'cpu','xcpu' , 'bitness' )
                # This is the important line; it makes sure that your file is safe to run before you invoke it.
                # This protects you from injection attacks / etc, if someone has placed malicious content into
                # the data file.
                $scriptBlock.CheckRestrictedLanguage($allowedCommands, $allowedVariables, $true)
                #
                return & $scriptBlock
            }
            catch
            {
                throw
            } 
        }

        #Specify the Language code of the modules to download ( not applicable to all modules) 
        $Language = $Host.CurrentUICulture.TwoLetterISOLanguageName # EN 

        #Specify the Language-Locale code of the modules to download ( not applicable to all modules) 
        $Culture = $Host.CurrentUICulture.Name

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
