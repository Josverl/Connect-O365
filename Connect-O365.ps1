<#PSScriptInfo
.TITLE Connect-O365
.VERSION 1.6.4
.GUID a3515355-c4b6-4ab8-8fa4-2150bbb88c96
.AUTHOR Jos Verlinde [MSFT]
.COMPANYNAME Microsoft
.COPYRIGHT 
.TAGS  O365 RMS 'Exchange Online' 'SharePoint Online' 'Skype for Business' 'PnP-Powershell' 'Office 365'
.LICENSEURI 
.PROJECTURI https://github.com/Josverl/Connect-O365
.ICONURI https://raw.githubusercontent.com/Josverl/Connect-O365/master/Connect-O365.png
.EXTERNALMODULEDEPENDENCIES MSOnline, Microsoft.Online.SharePoint.PowerShell, AADRM, OfficeDevPnP.PowerShell.V16.Commands
.REQUIREDSCRIPTS 
.EXTERNALSCRIPTDEPENDENCIES 
.RELEASENOTES
V1.6.4  Add autocompletion for saved accounts
V1.6.3  Add progress bars
V1.6.2  Resolve multiple Aliases per parameter bug on some PS flavors, 
V1.6.1  add test for Sign-in Assistant,Add pro-acive check for modules during administration.
V1.6.0  Publish to Github
v1.5.9  update install OS version match logic to use [System.Environment]::OSVersion.Version, correct DefaultParameterSetName=”Admin", Add -test option to check correct installation
V1.5.8  Seperate configuration download info from script, Retrieve Module info from github.
V1.5.7  Update to SPO shell build : 5111 1200 (March 2016)
v1.5.6  Add -close parameter and fixed parameter sets, added inline help to parameters 
v1.5.5  Fix Language for MSOnline / AAD module
v1.5    Add installation of dependent modules
v1.4    Correct bug wrt compliance search, remove prior created remote powershell sessions 
V1.3    Add dependend module information
V1.2    Add try-catch for SPO PNP Powershell, as that is less common
V1.1    Initial publication to scriptcenter
#>

<#
.Synopsis
   Connect to Office 365 and get ready to administer all services. 
   Includes installation of PowerShell modules 1/4/2016
.DESCRIPTION
   Connect to Office 365 and most related services and get ready to administer all services.
   The commandlet supports saving your administrative credentials in a safe manner so that it can be used in unattended files
   Allows Powershell administration of : O365, Azure AD , Azure RMS, Exchange Online, SharePoint Online including PNP Powershell
      
.EXAMPLE
   connect-O365 -Account 'admin@contoso.com' -SharePoint 

.EXAMPLE
   connect-O365 -Account 'admin@contoso.com' -SPO -EXO -Skype -Compliance -AADRM

.EXAMPLE
   #close any previously opened PS remote sessions (Exchange , Skype , Compliance Center) 
   connect-O365 -close

.EXAMPLE
   #Connect to MSOnline, and store securly store the credentials 
   connect-O365 -Account 'admin@contoso.com' -Persist:$false 
.EXAMPLE

   connect-O365 -Account 'admin@contoso.com'   

   #retrieve credentials for use in other cmdlets
   $Creds = Get-myCreds 'admin@contoso.com'

.EXAMPLE
   #Download and Install dependent Modules 
   connect-O365 -install
   
#>

[CmdletBinding(DefaultParameterSetName=”Admin")] 
[Alias("Connect-Office365")]
[OutputType([int])]
Param
(
    # Specify the (Admin) Account to authenticate with
    # paremeter completion is addedd in Dynamic Parameter s 
    [Parameter(ParameterSetName="Admin",Mandatory=$true,Position=0)]
    [ValidateNotNullOrEmpty()]
    [string]$Account,

    # Save the account credentials for later use        
    [Parameter(ParameterSetName="Admin",Mandatory=$false)]
    [switch]$Persist = $false, 
        
<# valid for Admin and Close #>
    #Connect to Azure AD aka MSOnline 
    [Parameter(ParameterSetName="Admin",Mandatory=$false)]
    [Parameter(ParameterSetName="Close",Mandatory=$false)]
    [Alias("AzureAD")] 
    [switch]$AAD = $false, 

    #Connect to Exchange Online
    [Parameter(ParameterSetName="Admin",Mandatory=$false)]
    [Parameter(ParameterSetName="Close",Mandatory=$false)]
    [Alias("EXO")] 
    [switch]$Exchange = $false, 

    #Connect to Skype Online
    [Parameter(ParameterSetName="Admin",Mandatory=$false)]
    [Parameter(ParameterSetName="Close",Mandatory=$false)]
    [Alias("CSO","Lync")] 
    [switch]$Skype = $false, 
    
    #Connecto to SharePoint Online
    [Parameter(ParameterSetName="Admin",Mandatory=$false)]
    [Parameter(ParameterSetName="Close",Mandatory=$false)]
    [Alias("SPO","ODFB")] 
    [switch]$SharePoint = $false, 

    #Connecto to SharePoint Online PNP
    [Parameter(ParameterSetName="Admin",Mandatory=$false)]
    [Parameter(ParameterSetName="Close",Mandatory=$false)]
    [Alias("PNP")] 
    [switch]$SharePointPNP = $false, 
        
    #Load and connecto to the O365 Compliance center
    [Parameter(ParameterSetName="Admin",Mandatory=$false)]
    [Parameter(ParameterSetName="Close",Mandatory=$false)]
    [switch]$Compliance = $false,

    #Connect to Azure Rights Management
    [Parameter(ParameterSetName="Admin",Mandatory=$false)]
    [Parameter(ParameterSetName="Close",Mandatory=$false)]
    [Alias("RMS","AzureRMS")] 
    [switch]$AADRM = $false,

    #All Services
    [Parameter(ParameterSetName="Admin",Mandatory=$false)]
    [Parameter(ParameterSetName="Close",Mandatory=$false)]
    [switch]$All = $false,

<# parameterset Close #>
    #Close all open Connections
    [Parameter(ParameterSetName="Close",Mandatory=$false)]
    [switch]$Close = $false,

<# parameterset INstall #>
    #Download and Install the supporting Modules
    [Parameter(ParameterSetName="Install",Mandatory=$true)]
    [switch]$Install,

    #Specify the Language code of the modules to download ( not applicable to all modules) 
    #Sample : -Language NL
    [Parameter(ParameterSetName="Install",Mandatory=$false)]
    [ValidatePattern("[a-zA-Z]{2}")]
    [Alias("Lang")] 
    $Language = $Host.CurrentUICulture.TwoLetterISOLanguageName, # EN 

    #Specify the Language-Locale code of the modules to download ( not applicable to all modules) 
    #Sample : -Language NL-NL  
    #Sample : -Language EN-US  
    [Parameter(ParameterSetName="Install",Mandatory=$false)]
    [ValidatePattern("[a-zA-Z]{2}-[a-zA-Z]{2}")]
    [Alias("LangCountry")]
    $Culture = $Host.CurrentUICulture.Name, # EN-US

    #Specify where to download the installable MSI and EXE modules to 
    [Parameter(ParameterSetName="Install",Mandatory=$false)]
    $Folder = $null, #'C:\Users\Jos\Downloads',

#    [Parameter(ParameterSetName="Install",Mandatory=$false)]
#    $InstallPreview = $true,

#Mixed parameterset

    # Save the account credentials for later use        
    [Parameter(ParameterSetName="Test",Mandatory=$false)]    
    [Parameter(ParameterSetName="Install",Mandatory=$false)]
    [switch]$Test = $false, 

    #Force asking for, and optionally force the Perstistance of the credentials.
    [Parameter(ParameterSetName="Admin",Mandatory=$false)]
    [Parameter(ParameterSetName="Install",Mandatory=$false)]
    [switch]$Force = $false
)

DynamicParam {
    if ($PSVersionTable.PSVersion -ge 5.0.0.0) {
        Register-ArgumentCompleter  -ParameterName "account" -ScriptBlock { 
             # Generate and set the CompletionSet
             #Types : DynamicKeyword / Property story / ProviderItem
             $arrSet = Get-ChildItem -Path "$env:USERPROFILE\Creds" 
             $arrSet | ForEach-Object {
                New-Object System.Management.Automation.CompletionResult $_.BaseName, $_.BaseName, 'DynamicKeyword', $_.FullName
             }
        }
    }
}

begin {

    function global:Store-myCreds ($username){
        $Credential = Get-Credential -Credential $username
        $Store = "$env:USERPROFILE\creds\$USERNAME.txt"
        MkDir "$env:USERPROFILE\Creds" -ea 0 | Out-Null
        $Credential.Password | ConvertFrom-SecureString | Set-Content $store
        Write-Verbose "Saved credentials to $store"
        return $Credential 
     }

    function global:Get-myCreds ($UserName , [switch]$Persist, [switch]$Force=$false){
        $Store = "$env:USERPROFILE\creds\$USERNAME.txt"
        if ( (Test-Path $store) -AND $Force -eq $false ) {
            #use a stored password if found , unless -force is used to ask for and store a new password
            Write-Verbose "Retrieved credentials from $store"
            $Password = Get-Content $store | ConvertTo-SecureString
            $Credential = New-Object System.Management.Automation.PsCredential($UserName,$Password)
            return $Credential
        } else {
            if ($persist -and -not [string]::IsNullOrEmpty($UserName)) {
                $admincredentials  = Store-myCreds $UserName
                return $admincredentials
            } else {
                return Get-Credential -Credential $username
            }
        }
     }

    # Verbose log of the input parameters
    Write-Verbose -Message 'Connect-O365 Parameters :'
    $PSBoundParameters.GetEnumerator() | ForEach-Object { Write-Verbose -Message "$($PSItem)" }

    #Parameter logic for explicit ans implicit -All 

    If ( $PsCmdlet.ParameterSetName -ieq "Close" ) 
    {
        if ( $all -eq $false -and  $exchange -eq $false -and $skype -eq $false -and $Compliance -eq $false  -and $SharePoint -eq $false -and $AADRM -eq $false) {
            Write-Verbose "Online Workload specified, assume all workloads"
            $all = $true
        }
    }
    If ( $PsCmdlet.ParameterSetName -iin "Close","Admin" ) 
    {
        if ( -not ( $Exchange -or $Skype -or $Compliance -or $SharePoint -or $SharePointPNP -or $AADRM ))
        { $AAD = $true } # default to AAD Only

        if ($SharePoint -or $SharePointPNP){
            ## AAD is used to find the URLS and the tenant name for SharePoint to connect to 
            $AAD = $true
        }
    
        if ($all) {
            $AAD = $true
            $Exchange = $true
            $Skype= $true
            $Compliance = $true
            $SharePoint= $true
            $SharePointPNP = $true
            $AADRM = $TRUE
        } 
    
    }
    # Start and step size for the progress baks 
    $script:Prog_pct = 0 
    $script:Prog_step = 12
}

Process{ 
    # Optionally close any prior sessions
    If ( $PsCmdlet.ParameterSetName -eq "Close") {
         Try {
            write-verbose "Closing open session(s) for :"
            Write-Progress "Connect-O365" -CurrentOperation "Closing" -PercentComplete $script:Prog_pct ; 
            $script:Prog_pct += $prog_step
            #Close Existing (remote Powershell Sessions) 
            if ($Exchange)   { 
                Write-Progress "Connect-O365" -CurrentOperation "Closing - Exchange Online" -PercentComplete $script:Prog_pct ; 
                $script:Prog_pct += $prog_step
                write-verbose "- Exchange Online"
                Get-PSSession -Name "Exchange Online" -ea SilentlyContinue | Remove-PSSession  } 
            if ($Compliance) { 
                Write-Progress "Connect-O365" -CurrentOperation "Closing - Compliance Center" -PercentComplete $script:Prog_pct ; 
                $script:Prog_pct += $prog_step
                write-verbose "- Compliance Center"
                Get-PSSession -Name "Compliance Center"  -ea SilentlyContinue | Remove-PSSession }
            if ($Skype)      { 
                Write-Progress "Connect-O365" -CurrentOperation "Closing - Skype Online" -PercentComplete $script:Prog_pct ; 
                $script:Prog_pct += $prog_step
                write-verbose "- Skype Online"
                Get-PSSession -Name "Skype Online" -ea SilentlyContinue| Remove-PSSession }
    
            if ($SharePoint) { 
                if ( get-module Microsoft.Online.SharePoint.Powershell )
                {
                    Write-Progress "Connect-O365" -CurrentOperation "Closing - SharePoint Online" -PercentComplete $script:Prog_pct ; 
                    $script:Prog_pct += $prog_step
                    write-verbose "- SharePoint Online"
                    Try {Disconnect-SPOService -ErrorAction Ignore } catch{}
                    
                }
            }
            if ($SharePointPNP) { 
                if ( get-module OfficeDevPnP.PowerShell.V16.Commands )
                {
                    #Also Disconnect PNPPowershell
                    write-verbose "- Disconnect PNP Powershell"
                    Write-Progress "Connect-O365" -CurrentOperation "Closing - SharePoint Online PnP Powershell" -PercentComplete $script:Prog_pct ; 
                    $script:Prog_pct += $prog_step
                    Try { Disconnect-SPOnline -ErrorAction Ignore } catch{}
                }
            } 
            If($AADRM) { 
                if ( get-module AADRM )
                {
                    write-verbose "- Azure RMS"
                    Write-Progress "Connect-O365" -CurrentOperation "Closing - Azure RMS" -PercentComplete $script:Prog_pct ; 
                    $script:Prog_pct += $prog_step
                    Disconnect-AadrmService 
                }
            } 
            return $True
        } catch {
            return $false
        }
        Finally {
            Write-Progress "Connect-O365" -Completed  
        }
    }

    # Admin , the main part and purpose 
    If ( $PsCmdlet.ParameterSetName -eq "Admin") {
        $operation = "Retrieve Credentials"
        write-verbose $Operation
        Write-Progress "Connect-O365" -CurrentOperation $Operation -PercentComplete $script:Prog_pct ; 
        $script:Prog_pct += $prog_step        
    
        $admincredentials = Get-myCreds $account -Persist:$Persist -Force:$Force
        if ($admincredentials -eq $null){ throw "A valid Tenant Admin Account is required." } 

        if ( $AAD) {
            $operation = "Connecting to Azure AD"
            write-verbose $Operation
            Write-Progress "Connect-O365" -CurrentOperation $Operation -PercentComplete $script:Prog_pct ; 
            $script:Prog_pct += $prog_step        

            $mod = 'MSOnline'
            if ( (get-module -Name $mod -ListAvailable) -eq $null ) {
                Write-warning "Required module: $mod is not installed or cannot be located."
                Write-Host "Install the missing module using the -Install parameter." -ForegroundColor Yellow
                return $false
            }
            #Imports the installed Azure Active Directory module.
            Import-Module MSOnline -Verbose:$false 
            #Establishes Online Services connection to Office 365 Management Layer.
            Connect-MsolService -Credential $admincredentials
        }

        if ($Skype ){
            $operation = "Connecting to Skype Online"
            write-verbose $Operation
            Write-Progress "Connect-O365" -CurrentOperation $Operation -PercentComplete $script:Prog_pct ; 
            $script:Prog_pct += $prog_step
                
            $mod = 'SkypeOnlineConnector'
            if ( (get-module -Name $mod -ListAvailable) -eq $null ) {
                Write-warning "Required module: $mod is not installed or cannot be located. "
                Write-Host "Install the missing module using the -Install parameter; or restart PowerShell" -ForegroundColor Yellow
                return $false
            }
            #Imports the installed Skype for Business Online services module.
            Import-Module SkypeOnlineConnector -Verbose:$false  -Force 

            #Remove prior  Session 
            Get-PSSession -Name "Skype Online" -ea SilentlyContinue| Remove-PSSession 

            #Create a Skype for Business Powershell session using defined credential.
            $SkypeSession = New-CsOnlineSession -Credential $admincredentials -Verbose:$false
            $SkypeSession.Name="Skype Online"

            #Imports Skype for Business session commands into your local Windows PowerShell session.
            Import-PSSession -Session  $SkypeSession -AllowClobber -Verbose:$false | Out-Null

        }


        If ($SharePoint) {
            $operation = "Connecting to SharePoint Online"
            write-verbose $Operation
            Write-Progress "Connect-O365" -CurrentOperation $Operation -PercentComplete $script:Prog_pct ; 
            $script:Prog_pct += $prog_step        

            if (!$AAD) {
                Throw "AAD Connection required"
            } else {
                #get tenant name for AAD Connection
                $tname= (Get-MsolDomain | ?{ $_.IsInitial -eq $true}).Name.Split(".")[0]
            }
            try {
                $mod = 'Microsoft.Online.Sharepoint.PowerShell'
                if ( (get-module -Name $mod -ListAvailable) -eq $null ) {
                    Write-warning "Required module: $mod is not installed or cannot be located. "
                    Write-Host "Install the missing module using the -Install parameter." -ForegroundColor Yellow
                    return $false
                }
                #Imports SharePoint Online session commands into your local Windows PowerShell session.
                Import-Module Microsoft.Online.Sharepoint.PowerShell -DisableNameChecking -Verbose:$false
                #lookup the tenant name based on the intial domain for the tenant
                Connect-SPOService -url https://$tname-admin.sharepoint.com -Credential $admincredentials          
            }
            catch {
                Write-Warning "Unable to connect to SharePoint Online."
            }
        }

        If ($SharePointPNP) {

            $operation = "Connecting to SharePoint Online PnP"
            write-verbose $Operation
            Write-Progress "Connect-O365" -CurrentOperation $Operation -PercentComplete $script:Prog_pct ; 
            $script:Prog_pct += $prog_step        

            if (!$AAD) {
                Throw "AAD Connection required"
            } else {
                #get tenant name for AAD Connection
                $tname= (Get-MsolDomain | ?{ $_.IsInitial -eq $true}).Name.Split(".")[0]
            }

            try { 
                write-verbose "Connecting to SharePoint Online PNP"
                $mod = 'OfficeDevPnP.PowerShell.V16.Commands'
                if ( (get-module -Name $mod -ListAvailable) -eq $null ) {
                    Write-Warning "Required module: $mod is not installed or cannot be located. "
                    Write-Host "Install the missing module using the -Install parameter." -ForegroundColor Yellow
                    #return $false
                }
                import-Module OfficeDevPnP.PowerShell.V16.Commands -DisableNameChecking -Verbose:$false
                Connect-SPOnline -Credential $admincredentials -url "https://$tname.sharepoint.com"
            } catch {
                Write-Warning "Unable to connect to SharePoint Online using the PnP PowerShell module"
            }
        }

        if ($Exchange ) {
            $operation = "Connecting to Exchange Online"
            write-verbose $Operation
            Write-Progress "Connect-O365" -CurrentOperation $Operation -PercentComplete $script:Prog_pct ; 
            $script:Prog_pct += $prog_step        

            #Remove prior  Session 
            Get-PSSession -Name "Exchange Online" -ea SilentlyContinue| Remove-PSSession 

            #Creates an Exchange Online session using defined credential.
            $ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $admincredentials -Authentication "Basic" -AllowRedirection
            $ExchangeSession.Name = "Exchange Online"
            #This imports the Office 365 session into your active Shell.
            Import-PSSession $ExchangeSession -AllowClobber -Verbose:$false -DisableNameChecking | Out-Null
        }

        if ($Compliance) {
            $operation = "Connecting to the Unified Compliance Center"
            write-verbose $Operation
            Write-Progress "Connect-O365" -CurrentOperation $Operation -PercentComplete $script:Prog_pct ; 
            $script:Prog_pct += $prog_step        
            #Remove prior  Session 
            Get-PSSession -Name "Compliance Center" -ea SilentlyContinue| Remove-PSSession 

            $PSCompliance = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $AdminCredentials -Authentication Basic -AllowRedirection -WarningAction Ignore
            $PSCompliance.Name = "Compliance Center"
            Import-PSSession $PSCompliance -AllowClobber -Verbose:$false -DisableNameChecking | Out-Null
        }
        If ($AADRM) {
            $operation = "Connecting to Azure Rights Management"
            write-verbose $Operation
            Write-Progress "Connect-O365" -CurrentOperation $Operation -PercentComplete $script:Prog_pct ; 
            $script:Prog_pct += $prog_step        
            #Azure RMS 
            $mod = 'AADRM'
            if ( (get-module -Name $mod -ListAvailable) -eq $null ) {
                Write-Warning "Required module: $mod is not installed or cannot be located. "
                Write-Host "Install the missing module using the -Install parameter." -ForegroundColor Yellow
                return $false
            }
            import-module AADRM -Verbose:$false
            Connect-AadrmService -Credential $admincredentials 
        }
    }

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


    If ( $PsCmdlet.ParameterSetName -eq "Install") {
        Write-Host -f Yellow "Starting Installation"
    
        #always perform module test after installation 
        $test = $true
    # Get the location of the downloads folder 
    # Ref : http://stackoverflow.com/questions/25049875/getting-any-special-folder-path-in-powershell-using-folder-guid 
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
        #Lookup downloads location
        if ($Folder -eq $null ) {$folder = [shell32]::GetKnownFolderPath([KnownFolder]::Downloads) }
        write-verbose "Download folder : $folder"

        $operation = "Load the required module information from the configuration file"
        write-verbose $Operation
        Write-Progress "Install External PowerShell modules to connect to Office 365" `
            -CurrentOperation $Operation -PercentComplete $script:Prog_pct ; 

        #load the required modules from a configuration file on GitHub
        $Components = Import-DataFile -url 'https://raw.githubusercontent.com/Josverl/Connect-O365/master/RequiredModuleInfo.psd1' 
    
        #figure out the progress rate 
        $script:Prog_step = 100 / $components.AdminComponents.Count
        $script:Prog_pct = $script:Prog_step

    #>
        # (Get-Module aadrm -ListAvailable).Version
        <# use the below in order to update relevant information 
        Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | sort -Property DisplayName | select PSChildName, DisplayName, Publisher, DisplayVersion
        #>
        foreach ($c  in $Components.AdminComponents) {
            $operation = "Install $($C.Name)"
            write-verbose $Operation
            Write-Progress "Install External PowerShell modules to connect to Office 365" `
                -CurrentOperation $Operation -PercentComplete $script:Prog_pct ; 
            $script:Prog_pct += $prog_step 

            if ($c.Preview -ieq "Yes" -and $InstallPreview -eq $false) {
                    write-host -f Gray "Skip Preview component : $($c.Name)"
                    continue; 
            }
            #IF OS Major Specified , and if the current OS Matches the specified OS
            if ($c.OS) { 
                if ( ($c.OS).Split(",") -notcontains [System.Environment]::OSVersion.Version.Major) { 
                    write-host -f Gray "OS mismatch, Skip component : $($c.Name)"
                    continue; 
                } 
            }

            switch ($c.Type.ToUpper() )
            {
                {$_ -in "EXE","MSI"} {
                    Write-Verbose "EXE or MSI package"
                    $AppInfo = Get-ItemProperty "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\$($c.ID)" -ErrorAction SilentlyContinue
                    $Installed = $false;
                    if ($AppInfo ) { 
                        $Installed = $appInfo.Displayname -ne $null
                    }
                
                    if (-not $Installed -or $force) {
                        # Filename
                        $file = Split-Path $c.Source -Leaf
                        #path in downloads 
                        $msi = join-path $folder $file
                        #remove existing item prior to Downloading 
                        Remove-Item $msi -Force -ea SilentlyContinue
                        try { 
                            #download it 
                            Write-Verbose "Download package"
                            Invoke-WebRequest $c.Source -OutFile $msi
                            if ( Test-Path $msi ) { 
                                $Sign = Get-AuthenticodeSignature $msi
                                if ( $Sign.Status -eq 'Valid' -and $sign.SignerCertificate.DnsNameList[0].Unicode -eq 'Microsoft Corporation' ) {
                                    if ($force) { 
                                        #de-install before re-install
                                        write-host -ForegroundColor Yellow "Removing current $($c.Type) package"
                                        Start-Process -FilePath "msiexec" -ArgumentList "/uninstall $($c.ID) " -Wait
                                    }
                                    try { 
                                        if ($c.Type -ieq "MSI" ) {
                                            Write-Verbose "Install MSI package : $msi"
                                            Start-Process -FilePath "msiexec" -ArgumentList "/package $msi /passive" -Wait
                                        } else {
                                            $Options = "/Passive"
                                            if ($c.Setup ) { $Options = $c.SetupOptions }
                                            Write-Verbose "Install EXE package : $msi $options"                                        
                                            Start-Process -FilePath $MSI -ArgumentList $options -Wait
                                        }

                                            $AppInfo = Get-ItemProperty "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\$($c.ID)" -ErrorAction SilentlyContinue
                                            Write-Host $c.Name "Version :" $AppInfo.DisplayVersion " was installed" -f Green
                                    } catch {
                                        Write-warning "$($c.Name) could not be installed"
                                        #Open in browser
                                        if (-not [string]::IsNullOrEmpty($c.Web)){
                                            Start-Process $c.Web
                                        }
                                    } 
                                }
                            }
                        } catch {
                            Write-Warning "could not install: $($c.Name)"
                        }
                    } 
                    else { 
                        Write-Host $c.Name "Version :" $AppInfo.DisplayVersion " is already installed"
                    }
                }
                "MODULE" {
                    # Add check for PS5 / WMF 5 
                    if (Get-Command install-module) { 
                        #check for installed version of this module 
                        $Current  = Get-Module -Name $c.Module -ListAvailable | sort -Property Version | select -First 1
                        $Source = Find-Module -Name $c.Module -Repository $c.Source -Verbose:$false
                        if ( $Current -eq $null ) {
                            write-verbose "Preparing to install module $($c.Module)"
                            #Not yet installed , find source 
                            #if not installed or newer version avaialble
                            if( $Source -and $Current -eq $null -or  ($Source.Version -GT $Current.Version) ) {
                                #install it 
                                $IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator") 
                                #Check Admin Perms 
                                if (-not $IsAdmin) {
                                    Write-Verbose "Start PS elevated to admin to run the install"
                                    Start-Process "$psHome\powershell.exe" -Verb Runas -ArgumentList "-Command Install-Module -Name $($c.Module) -Repository $($c.Source)" 

                                } else {
                                    Install-Module -InputObject $source
                                }
                                $Now = Get-Module $c.Module -ListAvailable -Verbose:$false
                                if ($now){
                                    Write-Host "Installed $($now.Name) version : $($Now.Version)"
                                }
                            } else {
                                #Could not  be found 
                                Write-warning  "The Module $($c.Name) Could not be located"
                                #Open in browser
                                if (-not [string]::IsNullOrEmpty($c.Web)){
                                    Start-Process $c.Web
                                }
                            } 
                        } else {
                            #version already installed
                            if ($Source.Version -gt $Current.Version -or $force ) {
                                write-verbose "Updating Module $($c.Module)"
                                if (-not $IsAdmin) {
                                    Write-Verbose "Start PS elevated to admin to run the install"
                                    Start-Process "$psHome\powershell.exe" -Verb Runas -ArgumentList "-Command update-Module -Name $($c.Module)" 
                                } else {
                                    update-Module -InputObject $source -Force:$force
                                }
                            }
                            else 
                            {
                                Write-verbose "Version : $($Current.Version) of $($c.Name) is already installed"
                            }
                            $NOW  = Get-Module $c.Module -ListAvailable
                            Write-Host "Version : $($NOW.Version) of $($c.Name) was newly installed"
                                       
                        }
                    } 
                    else 
                    { 
                        #No PS5 / WMF 5
                        Write-warning  "The Module $($c.Name) cannot be installed automatically. Please install manually or install WMF 5 (preview)"
                        #Open in browser
                        if (-not [string]::IsNullOrEmpty($c.Web)){
                            Start-Process $c.Web
                        }
                    }
                }
            default { Write-Warning "Unknown component type"}
            }
        }
    }

    #test is both a parameterset as well as an option for installation
    if ($test )  {
        Write-Host -f Yellow "Starting Test"    
        $script:Prog_pct = 0
        $script:Prog_step = 100 /6

        $ServiceName ='Microsoft Online Services Sign-in Assistant'

        $operation = $ServiceName
        write-verbose $Operation
        Write-Progress "Test and validate External powershell modules to connect to Office 365" `
            -CurrentOperation $Operation -PercentComplete $script:Prog_pct ; 
        $script:Prog_pct += $prog_step 

        Write-Host "Validating Service: $ServiceName" -NoNewline
        $SignInAssistant = Get-Service -Name msoidsvc
        if ( $SignInAssistant -eq $null )
        {
            Write-Host 
            Write-Warning "Service : '$ServiceName' is not installed"
        } else {
            if ($SignInAssistant.Status -ine "Running" ) {
                Write-Host 
                Write-Warning "Service '$ServiceName' is not running"
                Write-Host "Install the missing module using the -Install parameter." -ForegroundColor Yellow
            }
            else {
                Write-Host " - OK" -ForegroundColor Green
            }
        }

        #test if all Local modules are installed correctly 
        foreach ($module in @( "MSonline","SkypeOnlineConnector","Microsoft.Online.Sharepoint.PowerShell","OfficeDevPnP.PowerShell.V16.Commands","AADRM" ) ) {
            $operation = $Module
            write-verbose $Operation
            Write-Progress "Test and validate External powershell modules to connect to Office 365" `
                -CurrentOperation $Operation -PercentComplete $script:Prog_pct ; 
            $script:Prog_pct += $prog_step 

            Write-Host "Validating Module : $Module" -NoNewline
            $M = Get-Module -Name $module -ListAvailable
            if ($m -eq $null) {
                Write-Host 
                Write-warning "Module '$Module' is not installed or cannot be located."
                Write-Host "Install the missing module using the -Install parameter, or restart PowerShell." -ForegroundColor Yellow

            } else {
                Try {
                    Import-Module -Name $module -Force -DisableNameChecking
                    remove-module -Name $module -Force
                    Write-Host " - OK" -ForegroundColor Green
                } catch   {
                    Write-Host 
                    Write-warning "Module '$Module' could not be Imported."
                    Write-Host "Install the missing module using the -Install parameter, or restart PowerShell." -ForegroundColor Yellow
                }
            }
        }
    }
}
