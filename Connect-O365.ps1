<#PSScriptInfo
.TITLE Connect-O365
.VERSION 1.6.0
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
V1.6.0  move to Github
v1.5.9  update install OS version match logic to use [System.Environment]::OSVersion.Version
        correct DefaultParameterSetName=”Admin"
        Add -test option to check correct installation
V1.5.8  Seperate configuration download info from script, 
        Retrieve Module info from github.
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
    [switch]$AAD = $true, 

    #Connect to Exchange Online
    [Parameter(ParameterSetName="Admin",Mandatory=$false)]
    [Parameter(ParameterSetName="Close",Mandatory=$false)]
    [Alias("EXO")] 
    [switch]$Exchange = $false, 

    #Connect to Skype Online
    [Parameter(ParameterSetName="Admin",Mandatory=$false)]
    [Parameter(ParameterSetName="Close",Mandatory=$false)]
    [Alias("CSO")] 
    [Alias("Lync")] 
    [switch]$Skype = $false, 
    
    #Connecto to SharePoint Online
    [Parameter(ParameterSetName="Admin",Mandatory=$false)]
    [Parameter(ParameterSetName="Close",Mandatory=$false)]
    [Alias("SPO")] 
    [switch]$SharePoint = $false, 
        
    #Load and connecto to the O365 Compliance center
    [Parameter(ParameterSetName="Admin",Mandatory=$false)]
    [Parameter(ParameterSetName="Close",Mandatory=$false)]
    [switch]$Compliance = $false,

    #Connect to Azure Rights Management
    [Parameter(ParameterSetName="Admin",Mandatory=$false)]
    [Parameter(ParameterSetName="Close",Mandatory=$false)]
    [Alias("AZRMS")] 
    [Alias("RMS")]
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
    $Language = 'EN',

    #Specify the Language-Locale code of the modules to download ( not applicable to all modules) 
    #Sample : -Language NL-NL  
    #Sample : -Language EN-US  
    [Parameter(ParameterSetName="Install",Mandatory=$false)]
    [ValidatePattern("[a-zA-Z]{2}-[a-zA-Z]{2}")]
    $LangCountry = $Host.CurrentUICulture.Name,

    #Specify where to download the installable MSI and EXE modules to 
    [Parameter(ParameterSetName="Install",Mandatory=$false)]
    $Folder = $null, #'C:\Users\Jos\Downloads',

#    [Parameter(ParameterSetName="Install",Mandatory=$false)]
#    $InstallPreview = $true,


    # Save the account credentials for later use        
    [Parameter(ParameterSetName="Test",Mandatory=$false)]    
    [Parameter(ParameterSetName="Install",Mandatory=$false)]
    [switch]$Test = $false, 

    #Not in a specific parameterset

    #Force asking for, and optionally force the Perstistance of the credentials.
    [Parameter(ParameterSetName="Admin",Mandatory=$false)]
    [Parameter(ParameterSetName="Install",Mandatory=$false)]
    [switch]$Force = $false

)

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

# 
Write-Verbose -Message 'Connect-O365 Parameters :'
$PSBoundParameters.GetEnumerator() | ForEach-Object { Write-Verbose -Message "$($PSItem)" }

#Parameter logic for explicit ans implicit -All 

If ( $PsCmdlet.ParameterSetName -iin "Close","Admin" ) 
{
    if ( $all -eq $false -and  $exchange -eq $false -and $skype -eq $false -and $Compliance -eq $false  -and $SharePoint -eq $false -and $AADRM -eq $false) {
        Write-Verbose "Online Workload specified, assume all workloads"
        $all = $true
    }
    if ($all) {
        $AAD = $true
        $Exchange = $true
        $Skype= $true
        $Compliance = $true
        $SharePoint= $true
        $AADRM = $TRUE
    }
}


If ( $PsCmdlet.ParameterSetName -eq "Close") {
    write-verbose "Closing open session(s) for :"
    #Close Existing (remote Powershell Sessions) 
    if ($Exchange)   { 
        write-verbose "- Exchange Online"
        Get-PSSession -Name "Exchange Online" -ea SilentlyContinue | Remove-PSSession  } 
    if ($Compliance) { 
        write-verbose "- Compliance Center"
        Get-PSSession -Name "Compliance Center"  -ea SilentlyContinue | Remove-PSSession }
    if ($Skype)      { 
        write-verbose "- Skype Online"
        Get-PSSession -Name "Skype Online" -ea SilentlyContinue| Remove-PSSession }
    
    if ($SharePoint) { 
        write-verbose "- SharePoint Online"
        Try {Disconnect-SPOService -ErrorAction Ignore } catch{}
        write-verbose "- Disconnect PNP Powershell"
        #Also Disconnect PNPPowershell
        Try { Disconnect-SPOnline -ErrorAction Ignore } catch{}
    } 
    If($AADRM) { 
        write-verbose "- Azure RMS"
        Disconnect-AadrmService 
    } 
    return
}


If ( $PsCmdlet.ParameterSetName -eq "Admin") {

    $admincredentials = Get-myCreds $account -Persist:$Persist -Force:$Force
    if ($admincredentials -eq $null){ throw "A valid Tenant Admin Account is required." } 


    if ( $AAD) {
        write-verbose "Connecting to Azure AD"
        #Imports the installed Azure Active Directory module.
        Import-Module MSOnline -Verbose:$false 
        if (-not (Get-Module MSOnline ) ) { Throw "Module not installed"}
        #Establishes Online Services connection to Office 365 Management Layer.
        Connect-MsolService -Credential $admincredentials
    }

    if ($Skype ){
        write-verbose "Connecting to Skype Online"
        #Imports the installed Skype for Business Online services module.
        Import-Module SkypeOnlineConnector -Verbose:$false  -Force 

        #Remove prior  Session 
        Get-PSSession -Name "Skype Online" -ea SilentlyContinue| Remove-PSSession 

        #Create a Skype for Business Powershell session using defined credential.
        $SkypeSession = New-CsOnlineSession -Credential $admincredentials -Verbose:$false
        $SkypeSession.Name="Skype Online"

        #Imports Skype for Business session commands into your local Windows PowerShell session.
        Import-PSSession -Session  $SkypeSession -AllowClobber -Verbose:$false

    }


    If ($SharePoint) {
        write-verbose "Connecting to SharePoint Online"
        if (!$AAD) {
            Throw "AAD Connection required"
        } else {
            #get tenant name for AAD Connection
            $tname= (Get-MsolDomain | ?{ $_.IsInitial -eq $true}).Name.Split(".")[0]
        }

        #Imports SharePoint Online session commands into your local Windows PowerShell session.
        Import-Module Microsoft.Online.Sharepoint.PowerShell -DisableNameChecking -Verbose:$false
        #lookup the tenant name based on the intial domain for the tenant
        Connect-SPOService -url https://$tname-admin.sharepoint.com -Credential $admincredentials

        try { 
            write-verbose "Connecting to SharePoint Online PNP"
            import-Module OfficeDevPnP.PowerShell.V16.Commands -DisableNameChecking -Verbose:$false
            Connect-SPOnline -Credential $admincredentials -url "https://$tname.sharepoint.com"
        } catch {
            Write-Warning "Unable to connecto to SharePoint Online using the PNP PowerShell module"
        }
    }


    if ($Exchange ) {
        write-verbose "Connecting to Exchange Online"

        #Remove prior  Session 
        Get-PSSession -Name "Exchange Online" -ea SilentlyContinue| Remove-PSSession 

        #Creates an Exchange Online session using defined credential.
        $ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $admincredentials -Authentication "Basic" -AllowRedirection
        $ExchangeSession.Name = "Exchange Online"
        #This imports the Office 365 session into your active Shell.
        Import-PSSession $ExchangeSession -AllowClobber -Verbose:$false

    }

    if ($Compliance) {
        write-verbose "Connecting to the Unified Compliance Center"
        #Remove prior  Session 
        Get-PSSession -Name "Compliance Center" -ea SilentlyContinue| Remove-PSSession 

        $PSCompliance = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $AdminCredentials -Authentication Basic -AllowRedirection
        $PSCompliance.Name = "Compliance Center"
        Import-PSSession $PSCompliance -AllowClobber -Verbose:$false 

    }


    If ($AADRM) {
        write-verbose "Connecting to Azure Rights Management"    
        #Azure RMS 

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

    #load the required modules from a configuration file on GitHub
    $Components = Import-DataFile -url 'https://raw.githubusercontent.com/Josverl/Connect-O365/master/RequiredModuleInfo.psd1' 

#>
    # (Get-Module aadrm -ListAvailable).Version
    <# use the below in order to update relevant information 
    Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | sort -Property DisplayName | select PSChildName, DisplayName, Publisher, DisplayVersion
    #>
    foreach ($c  in $Components.AdminComponents) {
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
                    $Current  = Get-Module -Name $c.Module -ListAvailable
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
                            Write-verbose "$($c.Name) Version : $($Current.Version) is already installed"
                        }
                        $NOW  = Get-Module $c.Module -ListAvailable
                        Write-Host $c.Name "Version :" $NOW.Version" is now installed"
                                       
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
    #test if all Local modules are installed correctly 
    foreach ($module in @( "MSonline","SkypeOnlineConnector","Microsoft.Online.Sharepoint.PowerShell","OfficeDevPnP.PowerShell.V16.Commands","AADRM" ) ) {
        Write-Host "Validating Module : $Module" 
        $M = Get-Module -Name $module -ListAvailable
        if ($m -eq $null) {
            Write-warning "Module : $Module Could not be found."
        } else {
            Try {
                Import-Module -Name $module -Force -DisableNameChecking
                remove-module -Name $module -Force
                Write-Host "- OK" -ForegroundColor Green
            } catch   {
                Write-warning "Module : $Module Could not be Loaded."
            }
        }
    }
}


