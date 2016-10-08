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
        
        #Specify the Language code of the modules to download ( not applicable to all modules)
        if( $Language -eq $null ) {  
            $Language = $Host.CurrentUICulture.TwoLetterISOLanguageName # EN
        }

        #Specify the Language-Locale code of the modules to download ( not applicable to all modules)
        if ($Culture -eq $null) {
            $Culture = $Host.CurrentUICulture.Name
        }
        
        $Filename = Split-Path $url -Leaf # $URL.Split("/")[-1]
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
        
        #return the processed restrictedlanguage file 
        return & $scriptBlock
    }
    catch
    {
        throw "Error reading configuration file from $Url"
    } 
}


<#
.Synopsis
    Get the location of the downloads folder via Interop 
.DESCRIPTION
    [shell32]::GetKnownFolderPath([KnownFolder]::Downloads
    Ref : http://stackoverflow.com/questions/25049875/getting-any-special-folder-path-in-powershell-using-folder-guid
#>
Add-Type @"
    using System;
    using System.Runtime.InteropServices;

    public static class KnownFolder
    {
        public static readonly Guid Documents = new Guid( "FDD39AD0-238F-46AF-ADB4-6C85480369C7" );
        public static readonly Guid Downloads = new Guid( "374DE290-123F-4565-9164-39C4925E467B" );
    }
    public class shell32
    {
        [DllImport("shell32.dll")]
        private static extern int SHGetKnownFolderPath(
                [MarshalAs(UnmanagedType.LPStruct)] 
                Guid rfid,
                uint dwFlags,
                IntPtr hToken,
                out IntPtr pszPath
            );
            public static string GetKnownFolderPath(Guid rfid)
            {
            IntPtr pszPath;
            if (SHGetKnownFolderPath(rfid, 0, IntPtr.Zero, out pszPath) != 0)
                return ""; // add whatever error handling you fancy
            string path = Marshal.PtrToStringUni(pszPath);
            Marshal.FreeCoTaskMem(pszPath);
            return path;
            }
    }
"@ 

<#
function Get-O365ModuleFile  {
[CmdletBinding()]
#[OutputType([String])]
param(
)    
    $operation = "Load the required module information from the configuration file"
    write-verbose $Operation
    Write-Progress "Install External PowerShell modules to connect to Office 365" `
        -CurrentOperation $Operation -PercentComplete $script:Prog_pct ; 

    #load the required modules from a configuration file on GitHub
    $Components = Import-DataFile -url 'https://raw.githubusercontent.com/Josverl/Connect-O365/master/RequiredModuleInfo.psd1' 
}
#>

