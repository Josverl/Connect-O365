<#PSScriptInfo
.TITLE Connect-O365
.VERSION 1.7.5
.GUID a3515355-c4b6-4ab8-8fa4-2150bbb88c96
.AUTHOR Jos Verlinde [MSFT]
.COMPANYNAME Microsoft
.COPYRIGHT Jos Verlinde 2016
.TAGS  O365 RMS 'Exchange Online' 'SharePoint Online' 'Skype for Business' 'PnP-Powershell' 'Office 365'
.LICENSEURI https://github.com/Josverl/Connect-O365/raw/master/License
.PROJECTURI https://github.com/Josverl/Connect-O365
.ICONURI https://raw.githubusercontent.com/Josverl/Connect-O365/master/Connect-O365.png
.EXTERNALMODULEDEPENDENCIES MSOnline, Microsoft.Online.SharePoint.PowerShell, AADRM
.REQUIREDSCRIPTS 
.EXTERNALSCRIPTDEPENDENCIES 
.RELEASENOTES
v1.7.2  Update tests for changed external dependency name SharePointPnPPowerShellOnline
V1.7.1  Minor improvements in account lookup     
V1.6.9  Updated changed external dependency name SharePointPnPPowerShellOnline
V1.6.8  Add global variables $TenantName and $AdminName, consistent parameter names 
V1.6.7  Correct script for CredentialManager 2.0.0.0 parameter changes 
V1.6.5  Add autocompletion for saved accounts and credential manager, change default for -AAD, improve connection error checks
V1.6.3  Add progress bars
V1.6.2  Resolve multiple Aliases per parameter bug on some PS flavors, 
V1.6.1  Add test for Sign-in Assistant,Add pro-acive check for modules during administration.
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

#Requires -Module @{ModuleName="CredentialManager";ModuleVersion="2.0"}
#Requires -Module @{ModuleName='ConnectO365';ModuleVersion="0.3"}

<# #Requires -Module SharePointPNPPowershellOnline #>

<#
.Synopsis
   Connect to Office 365 and get ready to administer all services. 
   Includes installation of PowerShell modules 1/4/2016
.DESCRIPTION
   Connect to Office 365 and most related services and get ready to administer all services.
   The commandlet supports saving your administrative credentials in a safe manner so that it can be used in unattended files, and allows easy recall from the command line using autocompletion of saved credentials.
   Allows Powershell administration of : O365, Azure AD , Azure RMS, Exchange Online, SharePoint Online including PNP Powershell
      
.EXAMPLE
   Connect-O365 -Account 'admin@contoso.com' -SharePoint 
   #Note: intellisense 
   
.EXAMPLE
   Connect-O365 -Account 'admin@contoso.com' -SPO -EXO -Skype -Compliance -AADRM

.EXAMPLE
   #close any previously opened PS remote sessions (Exchange , Skype , Compliance Center) 
   Connect-O365 -close

.EXAMPLE
   #Connect to MSOnline, and store securly store the credentials 
   connect-O365 -Account 'admin@contoso.com' -Persist
.EXAMPLE
   #Update the password for the admin account 
   connect-O365 -Account 'admin@contoso.com' -Persist

.EXAMPLE
   connect-O365 -Account 'admin@contoso.com' -Persist   

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
    [Alias("UCC")] 
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

    [Parameter(ParameterSetName="Admin",Mandatory=$false)]
    [Parameter(ParameterSetName="Test",Mandatory=$false)]
    [Parameter(ParameterSetName="Close",Mandatory=$false)]
    [switch]$PassThrough = $false,

    # Save the account credentials for later use        
    [Parameter(ParameterSetName="Test",Mandatory=$false)]    
    [Parameter(ParameterSetName="Install",Mandatory=$false)]
    [switch]$Test = $false, 

    #Force asking for, and optionally force the Perstistance of the credentials.
#   [Parameter(ParameterSetName="Admin",Mandatory=$false)]
    [Parameter(ParameterSetName="Install",Mandatory=$false)]
    [switch]$Force = $false
)

DynamicParam {
    if ($PSVersionTable.PSVersion.Major -ge 5) {
        Register-ArgumentCompleter  -ParameterName "account" -ScriptBlock { 
            # Generate and set the CompletionSet
            #Types : DynamicKeyword / Property story / ProviderItem

            #Find the credentials stored in the userprofile creds folder 
            $arrSet = Get-ChildItem -Path "$env:USERPROFILE\Creds" 
            $arrSet | ForEach-Object {
                # Completion text , ListItem text, Resulttype, Tooltip
                New-Object System.Management.Automation.CompletionResult $_.BaseName, $_.BaseName, 'ProviderItem', $_.FullName
            }
            #check if the credentialmanager module is installed 
            $CM = get-module credentialmanager -ListAvailable | select -Last 1
            if ($cm -ne $null -and $CM.Version -eq "2.0") {
                #Find the credentials stored in the credential manager (version 2.0
                $storedcredentials = Get-StoredCredential -Type GENERIC -AsCredentialObject -WarningAction SilentlyContinue

                #Only the onese with a specified targetname different from the username
                $credentials = $storedcredentials | where { $_.UserName -like '?*@?*' -and $_.Type -eq 'GENERIC'} | select -Property UserName, @{Name="TargetName";Expression={$($_.Targetname).Replace("LegacyGeneric:target=","")}} , Type, TargetAlias, Comment 
                $credentials = $credentials | where { $_.targetname -ne $_.Username}
                #now create the list               
                $credentials| ForEach-Object {
                    # Completion text , ListItem text, Resulttype, Tooltip
                    New-Object System.Management.Automation.CompletionResult $_.TargetName, $_.TargetName, 'DynamicKeyword', $_.Username
                }

                #Now All 
                $credentials = $storedcredentials | where { $_.UserName -like '?*@?*' -and $_.Type -eq 'GENERIC'} | select -Property UserName, TargetName, Type, TargetAlias, Comment 
                #now create the list               
                $credentials| ForEach-Object {
                    # Completion text , ListItem text, Resulttype, Tooltip
                    if ($_.Comment -ne $null ) {$TTIP = $_.Comment } else { $TTIP = $($_.Targetname).Replace("LegacyGeneric:target=","") }
                    New-Object System.Management.Automation.CompletionResult $_.UserName, $_.Username, 'History', $TTIP
                }
            }
        }
    }
}
begin {
    # Verbose log of the input parameters
    Write-Verbose -Message 'Connect-O365 Parameters :'
    $PSBoundParameters.GetEnumerator() | ForEach-Object { Write-Verbose -Message "$($PSItem)" }

    #Parameter logic for explicit and implicit -All 
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
                if ( get-module SharepointPnPPowerShellOnline )
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
            if ($PassThrough) { # Only return value if requested
                return $True
            }
        } catch {
            if ($PassThrough) { # Only return value if requested 
                return $false
            }
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
    
        #retrieve admin credentials for filestore and secure store 
        $admincredentials = retrieve-credentials -account $account -Persist:$persist
        if ($admincredentials -eq $null){ 
            Write-Verbose "No stored credentials could be found"
            throw "A valid Tenant Admin Account is required." 
        } 
        ${Global:AdminName} = $admincredentials.UserName
        Write-Host -f Cyan "UserName ${Global:AdminName}"
        
        ${Global:TenantName}  = $null
        if ( $AAD) {
            $operation = "Azure AD"
            write-verbose $Operation
            Write-Progress "Connect-O365" -CurrentOperation $Operation -PercentComplete $script:Prog_pct ; 
            $script:Prog_pct += $prog_step        

            $mod = 'MSOnline'
            if ( (get-module -Name $mod -ListAvailable) -eq $null ) {
                Write-warning "Required module: $mod is not installed or cannot be located."
                Write-Host "Install the missing module using the -Install parameter." -ForegroundColor Yellow
                if ($PassThrough ) { #Only return if requested 
                    return $false
                }
            }
            #Imports the installed Azure Active Directory module.
            Import-Module MSOnline -Verbose:$false 
            #Establishes Online Services connection to Office 365 Management Layer.
            try { 
                Connect-MsolService -Credential $admincredentials -ErrorAction SilentlyContinue
                #get tenant name for future use
                $Initial = Get-MsolDomain -ErrorAction SilentlyContinue | ?{ $_.IsInitial -eq $true}
                if ($Initial) {
                    ${Global:TenantName} = $Initial.Name.Split(".")[0]
                    Write-Host -f Green $operation
                } else  {
                    Write-Warning "Unable to connect to Azure AD"
                }
            } catch {
                Write-Warning "Unable to connect to Azure AD"
            }
        } else {
            #
        }

        if ($Skype ){
            $operation = "Skype Online"
            write-verbose $Operation
            Write-Progress "Connect-O365" -CurrentOperation $Operation -PercentComplete $script:Prog_pct ; 
            $script:Prog_pct += $prog_step
                
            $mod = 'SkypeOnlineConnector'
            if ( (get-module -Name $mod -ListAvailable) -eq $null ) {
                Write-warning "Required module: $mod is not installed or cannot be located. "
                Write-Host "Install the missing module using the -Install parameter; or restart PowerShell" -ForegroundColor Yellow
                if ($PassThrough ) { #Only return if requested 
                    return $false
                }
            }
            #Imports the installed Skype for Business Online services module.
            Import-Module SkypeOnlineConnector -Verbose:$false  -Force 

            #Remove prior  Session 
            Get-PSSession -Name "Skype Online" -ea SilentlyContinue| Remove-PSSession 

            #Create a Skype for Business Powershell session using defined credential.
            Try { 
                $SkypeSession = New-CsOnlineSession -Credential $admincredentials -Verbose:$false -ErrorAction SilentlyContinue -ErrorVariable ConnectError
                if ($SkypeSession) {
                    $SkypeSession.Name="Skype Online"
                    #Imports Skype for Business session commands into your local Windows PowerShell session.
                    Import-PSSession -Session  $SkypeSession -AllowClobber -Verbose:$false | Out-Null
                    Write-Host -f Green $operation
                } else {
                    Write-Warning $ConnectError[0].Message
                }
            } catch {
                Write-Warning $ConnectError[0].Message
            }

        }


        If ($SharePoint) {
            $operation = "SharePoint Online"
            write-verbose $Operation
            Write-Progress "Connect-O365" -CurrentOperation $Operation -PercentComplete $script:Prog_pct ; 
            $script:Prog_pct += $prog_step        

            if (!$AAD) {
                Throw "AAD Connection required"
            }
            try {
                $mod = 'Microsoft.Online.Sharepoint.PowerShell'
                if ( (get-module -Name $mod -ListAvailable) -eq $null ) {
                    Write-warning "Required module: $mod is not installed or cannot be located. "
                    Write-Host "Install the missing module using the -Install parameter." -ForegroundColor Yellow
                    if ($PassThrough ) { #Only return if requested 
                        return $false
                    }
                }
                #Imports SharePoint Online session commands into your local Windows PowerShell session.
                Import-Module Microsoft.Online.Sharepoint.PowerShell -DisableNameChecking -Verbose:$false
                #lookup the tenant name based on the intial domain for the tenant
                Connect-SPOService -url "https://${Global:TenantName}-admin.sharepoint.com" -Credential $admincredentials
                Write-Host -f Green $operation                         
            }
            catch {
                Write-Warning "Unable to connect to SharePoint Online."
            }
        }

        If ($SharePointPNP) {

            $operation = "SharePoint Online - PnP PowerShell"
            write-verbose $Operation
            Write-Progress "Connect-O365" -CurrentOperation $Operation -PercentComplete $script:Prog_pct ; 
            $script:Prog_pct += $prog_step        

            if (!$AAD) {
                Throw "AAD Connection required"
            } 

            try { 
                write-verbose $Operation
                $mod = 'SharepointPnPPowerShellOnline'
                if ( (get-module -Name $mod -ListAvailable) -eq $null ) {
                    Write-Warning "Required module: $mod is not installed or cannot be located. "
                    Write-Host "Install the missing module using the -Install parameter." -ForegroundColor Yellow
                    if ($PassThrough ) { #Only return if requested 
                        return $false
                    }
                }
                import-Module SharepointPnPPowerShellOnline -DisableNameChecking -Verbose:$false
                Connect-SPOnline -Credential $admincredentials -url "https://${Global:TenantName}.sharepoint.com"
                Write-Host -f Green $Operation
            } catch {
                Write-Warning "Unable to connect to SharePoint Online using the PnP PowerShell module"
            }
        }

        if ($Exchange ) {
            $operation = "Exchange Online"
            write-verbose $Operation
            Write-Progress "Connect-O365" -CurrentOperation $Operation -PercentComplete $script:Prog_pct ; 
            $script:Prog_pct += $prog_step        

            #Remove prior  Session 
            Get-PSSession -Name "Exchange Online" -ea SilentlyContinue| Remove-PSSession 

            #Creates an Exchange Online session using defined credential.
            $ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $admincredentials -Authentication "Basic" -WarningAction Ignore -ErrorAction SilentlyContinue -ErrorVariable ConnectError
            if ($ExchangeSession) {
                $ExchangeSession.Name = "Exchange Online"
                #This imports the Office 365 session into your active Shell.
                Import-PSSession $ExchangeSession -AllowClobber -Verbose:$false -DisableNameChecking | Out-Null
                Write-Host -f Green $Operation
            } else {
                Write-Warning $ConnectError[0].ErrorDetails
            }
        }

        if ($Compliance) {
            $operation = "Unified Compliance Center"
            write-verbose $Operation
            Write-Progress "Connect-O365" -CurrentOperation $Operation -PercentComplete $script:Prog_pct ; 
            $script:Prog_pct += $prog_step        
            #Remove prior  Session 
            Get-PSSession -Name "Compliance Center" -ea SilentlyContinue| Remove-PSSession 
            
            $PSCompliance = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $AdminCredentials -Authentication Basic -AllowRedirection -WarningAction Ignore -ErrorAction SilentlyContinue -ErrorVariable ConnectError
            if ($PSCompliance) {
                $PSCompliance.Name = "Compliance Center"
                Import-PSSession $PSCompliance -AllowClobber -Verbose:$false -DisableNameChecking | Out-Null
                Write-Host -f Green $Operation
            } else {
                Write-Warning $ConnectError[0].ErrorDetails
            }
        }
        If ($AADRM) {
            $operation = "Azure Rights Management"
            write-verbose $Operation
            Write-Progress "Connect-O365" -CurrentOperation $Operation -PercentComplete $script:Prog_pct ; 
            $script:Prog_pct += $prog_step        
            #Azure RMS 
            $mod = 'AADRM'
            if ( (get-module -Name $mod -ListAvailable) -eq $null ) {
                Write-Warning "Required module: $mod is not installed or cannot be located. "
                Write-Host "Install the missing module using the -Install parameter." -ForegroundColor Yellow
                if ($PassThrough ) { #Only return if requested 
                    return $false
                }
            }
            import-module AADRM -Verbose:$false

            #There is no good way to capture errors thrown 
            #it is only possible to ignore errors , but that also hides the extual error report 
            Connect-AadrmService -Credential $admincredentials
            try{
                $x =  Get-Aadrm -ErrorAction SilentlyContinue
                Write-Host -f Green $Operation
            } catch { 
                #do nothing 
            }             
            
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
                                            Start-Process -FilePath "msiexec" -ArgumentList "/package $msi /passive /promptrestart" -Wait
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

        $AllOk = $true    #Let's start Positive 

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
            $AllOk = $false
        } else {
            if ($SignInAssistant.Status -ine "Running" ) {
                Write-Host 
                Write-Warning "Service '$ServiceName' is not running"
                Write-Host "Install the missing module using the -Install parameter." -ForegroundColor Yellow
                $AllOk = $false
            }
            else {
                Write-Host " - OK" -ForegroundColor Green
            }
        }

        #test if all Local modules are installed correctly 
        foreach ($module in @( "MSonline","SkypeOnlineConnector","Microsoft.Online.Sharepoint.PowerShell","SharepointPnPPowerShellOnline","AADRM" ) ) {
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
                $AllOk = $false

            } else {
                Try {
                    Import-Module -Name $module -Force -DisableNameChecking
                    remove-module -Name $module -Force
                    Write-Host " - OK" -ForegroundColor Green
                } catch   {
                    Write-Host 
                    Write-warning "Module '$Module' could not be Imported."
                    Write-Host "Install the missing module using the -Install parameter, or restart PowerShell." -ForegroundColor Yellow
                    $AllOk = $false
                }
            }
        }
        #All test complete, and -Passthough specified, return the test result 
        if ($PassThrough) {
            return $AllOk
        }
    }
}

# SIG # Begin signature block
# MIIgNAYJKoZIhvcNAQcCoIIgJTCCICECAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUqlvn0i941uznz2i5dYi8M6Ml
# huygghtjMIIDtzCCAp+gAwIBAgIQDOfg5RfYRv6P5WD8G/AwOTANBgkqhkiG9w0B
# AQUFADBlMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMSQwIgYDVQQDExtEaWdpQ2VydCBBc3N1cmVk
# IElEIFJvb3QgQ0EwHhcNMDYxMTEwMDAwMDAwWhcNMzExMTEwMDAwMDAwWjBlMQsw
# CQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cu
# ZGlnaWNlcnQuY29tMSQwIgYDVQQDExtEaWdpQ2VydCBBc3N1cmVkIElEIFJvb3Qg
# Q0EwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCtDhXO5EOAXLGH87dg
# +XESpa7cJpSIqvTO9SA5KFhgDPiA2qkVlTJhPLWxKISKityfCgyDF3qPkKyK53lT
# XDGEKvYPmDI2dsze3Tyoou9q+yHyUmHfnyDXH+Kx2f4YZNISW1/5WBg1vEfNoTb5
# a3/UsDg+wRvDjDPZ2C8Y/igPs6eD1sNuRMBhNZYW/lmci3Zt1/GiSw0r/wty2p5g
# 0I6QNcZ4VYcgoc/lbQrISXwxmDNsIumH0DJaoroTghHtORedmTpyoeb6pNnVFzF1
# roV9Iq4/AUaG9ih5yLHa5FcXxH4cDrC0kqZWs72yl+2qp/C3xag/lRbQ/6GW6whf
# GHdPAgMBAAGjYzBhMA4GA1UdDwEB/wQEAwIBhjAPBgNVHRMBAf8EBTADAQH/MB0G
# A1UdDgQWBBRF66Kv9JLLgjEtUYunpyGd823IDzAfBgNVHSMEGDAWgBRF66Kv9JLL
# gjEtUYunpyGd823IDzANBgkqhkiG9w0BAQUFAAOCAQEAog683+Lt8ONyc3pklL/3
# cmbYMuRCdWKuh+vy1dneVrOfzM4UKLkNl2BcEkxY5NM9g0lFWJc1aRqoR+pWxnmr
# EthngYTffwk8lOa4JiwgvT2zKIn3X/8i4peEH+ll74fg38FnSbNd67IJKusm7Xi+
# fT8r87cmNW1fiQG2SVufAQWbqz0lwcy2f8Lxb4bG+mRo64EtlOtCt/qMHt1i8b5Q
# Z7dsvfPxH2sMNgcWfzd8qVttevESRmCD1ycEvkvOl77DZypoEd+A5wwzZr8TDRRu
# 838fYxAe+o0bJW1sj6W3YQGx0qMmoRBxna3iw/nDmVG3KwcIzi7mULKn+gpFL6Lw
# 8jCCBTAwggQYoAMCAQICEAQJGBtf1btmdVNDtW+VUAgwDQYJKoZIhvcNAQELBQAw
# ZTELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQ
# d3d3LmRpZ2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBS
# b290IENBMB4XDTEzMTAyMjEyMDAwMFoXDTI4MTAyMjEyMDAwMFowcjELMAkGA1UE
# BhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2lj
# ZXJ0LmNvbTExMC8GA1UEAxMoRGlnaUNlcnQgU0hBMiBBc3N1cmVkIElEIENvZGUg
# U2lnbmluZyBDQTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAPjTsxx/
# DhGvZ3cH0wsxSRnP0PtFmbE620T1f+Wondsy13Hqdp0FLreP+pJDwKX5idQ3Gde2
# qvCchqXYJawOeSg6funRZ9PG+yknx9N7I5TkkSOWkHeC+aGEI2YSVDNQdLEoJrsk
# acLCUvIUZ4qJRdQtoaPpiCwgla4cSocI3wz14k1gGL6qxLKucDFmM3E+rHCiq85/
# 6XzLkqHlOzEcz+ryCuRXu0q16XTmK/5sy350OTYNkO/ktU6kqepqCquE86xnTrXE
# 94zRICUj6whkPlKWwfIPEvTFjg/BougsUfdzvL2FsWKDc0GCB+Q4i2pzINAPZHM8
# np+mM6n9Gd8lk9ECAwEAAaOCAc0wggHJMBIGA1UdEwEB/wQIMAYBAf8CAQAwDgYD
# VR0PAQH/BAQDAgGGMBMGA1UdJQQMMAoGCCsGAQUFBwMDMHkGCCsGAQUFBwEBBG0w
# azAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMEMGCCsGAQUF
# BzAChjdodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVk
# SURSb290Q0EuY3J0MIGBBgNVHR8EejB4MDqgOKA2hjRodHRwOi8vY3JsNC5kaWdp
# Y2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMDqgOKA2hjRodHRw
# Oi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3Js
# ME8GA1UdIARIMEYwOAYKYIZIAYb9bAACBDAqMCgGCCsGAQUFBwIBFhxodHRwczov
# L3d3dy5kaWdpY2VydC5jb20vQ1BTMAoGCGCGSAGG/WwDMB0GA1UdDgQWBBRaxLl7
# KgqjpepxA8Bg+S32ZXUOWDAfBgNVHSMEGDAWgBRF66Kv9JLLgjEtUYunpyGd823I
# DzANBgkqhkiG9w0BAQsFAAOCAQEAPuwNWiSz8yLRFcgsfCUpdqgdXRwtOhrE7zBh
# 134LYP3DPQ/Er4v97yrfIFU3sOH20ZJ1D1G0bqWOWuJeJIFOEKTuP3GOYw4TS63X
# X0R58zYUBor3nEZOXP+QsRsHDpEV+7qvtVHCjSSuJMbHJyqhKSgaOnEoAjwukaPA
# JRHinBRHoXpoaK+bp1wgXNlxsQyPu6j4xRJon89Ay0BEpRPw5mQMJQhCMrI2iiQC
# /i9yfhzXSUWW6Fkd6fp0ZGuy62ZD2rOwjNXpDd32ASDOmTFjPQgaGLOBm0/GkxAG
# /AeB+ova+YJJ92JuoVP6EpQYhS6SkepobEQysmah5xikmmRR7zCCBTEwggQZoAMC
# AQICEAn8oNuSsuSr/+Oj5/k7HugwDQYJKoZIhvcNAQELBQAwcjELMAkGA1UEBhMC
# VVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0
# LmNvbTExMC8GA1UEAxMoRGlnaUNlcnQgU0hBMiBBc3N1cmVkIElEIENvZGUgU2ln
# bmluZyBDQTAeFw0xNjA5MjQwMDAwMDBaFw0xNzA4MDExMjAwMDBaMG4xCzAJBgNV
# BAYTAk5MMRUwEwYDVQQIEwxadWlkLUhvbGxhbmQxEjAQBgNVBAcTCVBpam5hY2tl
# cjEZMBcGA1UEChMQQWRyaWFhbiBWZXJsaW5kZTEZMBcGA1UEAxMQQWRyaWFhbiBW
# ZXJsaW5kZTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBANaBtMwH6hc7
# qnfSqNmlEUnx+WuUlc/s9iGc15cGoXX3C1iis/Eib+kBFrxiAmO5fZR1rfhHQqPu
# MWr5vLRzfSkPzU21SjeH1MWCV68UvxTEqG/GqMHZ0KfHiaG7BbL8+/j9a2PV8cBf
# JXyTh7J6xSdeK7jDvkNOPI/VbWGT4tsZ8LghzeKIv7po9jSdf2fVrkSxc1Nrhzd7
# JxdPmVxnzlFXg6V/d/i2+MqBBLsgOsnYcPpujq+KFL0iWrnGBs6YcAmbWgCuzo/1
# CBA8AJARVroaKjY2ublzPiO5vnvvjalJ6WZbo9Oy7TGXmOpnajk4HsgErTrPnyEj
# uqxRTLSZn0sCAwEAAaOCAcUwggHBMB8GA1UdIwQYMBaAFFrEuXsqCqOl6nEDwGD5
# LfZldQ5YMB0GA1UdDgQWBBS4iKChjKj7M8uDygqqUkPqQWbEajAOBgNVHQ8BAf8E
# BAMCB4AwEwYDVR0lBAwwCgYIKwYBBQUHAwMwdwYDVR0fBHAwbjA1oDOgMYYvaHR0
# cDovL2NybDMuZGlnaWNlcnQuY29tL3NoYTItYXNzdXJlZC1jcy1nMS5jcmwwNaAz
# oDGGL2h0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9zaGEyLWFzc3VyZWQtY3MtZzEu
# Y3JsMEwGA1UdIARFMEMwNwYJYIZIAYb9bAMBMCowKAYIKwYBBQUHAgEWHGh0dHBz
# Oi8vd3d3LmRpZ2ljZXJ0LmNvbS9DUFMwCAYGZ4EMAQQBMIGEBggrBgEFBQcBAQR4
# MHYwJAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBOBggrBgEF
# BQcwAoZCaHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0U0hBMkFz
# c3VyZWRJRENvZGVTaWduaW5nQ0EuY3J0MAwGA1UdEwEB/wQCMAAwDQYJKoZIhvcN
# AQELBQADggEBAB0RlGTVco8OVFOTE1oybL5rjKQSljdOiRuJje8iK71RuTJ/7jsi
# I2m8uMyY/6kjN7Hou3Ao/H0L4JVOGrYKOnLfDP64LowW9Kbsn3k+m6QLePPBk98M
# v74fSzBdqdPJrCJkw5+nZFHmXZb9EhKyinEg2rua3XA0oUV51QxSJfrHHuA7CXUs
# 7vGRCzIPewM+OxYu2MRvI6sOmLXk1jNOJnnkvx5cVoKTk2gmDKiTLLxLZCjXt9bf
# H8AoGLhoyGV6NLZ/Yz/brGGJwNqzjuqOIDZjJ3ShxP8Tx9NjdssSKBMs8oG8pjY5
# VWqMyLKaK0r4JNQD3tZmCHUy+HfovC8UAD0wggZqMIIFUqADAgECAhADAZoCOv9Y
# sWvW1ermF/BmMA0GCSqGSIb3DQEBBQUAMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQK
# EwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNV
# BAMTGERpZ2lDZXJ0IEFzc3VyZWQgSUQgQ0EtMTAeFw0xNDEwMjIwMDAwMDBaFw0y
# NDEwMjIwMDAwMDBaMEcxCzAJBgNVBAYTAlVTMREwDwYDVQQKEwhEaWdpQ2VydDEl
# MCMGA1UEAxMcRGlnaUNlcnQgVGltZXN0YW1wIFJlc3BvbmRlcjCCASIwDQYJKoZI
# hvcNAQEBBQADggEPADCCAQoCggEBAKNkXfx8s+CCNeDg9sYq5kl1O8xu4FOpnx9k
# WeZ8a39rjJ1V+JLjntVaY1sCSVDZg85vZu7dy4XpX6X51Id0iEQ7Gcnl9ZGfxhQ5
# rCTqqEsskYnMXij0ZLZQt/USs3OWCmejvmGfrvP9Enh1DqZbFP1FI46GRFV9GIYF
# jFWHeUhG98oOjafeTl/iqLYtWQJhiGFyGGi5uHzu5uc0LzF3gTAfuzYBje8n4/ea
# 8EwxZI3j6/oZh6h+z+yMDDZbesF6uHjHyQYuRhDIjegEYNu8c3T6Ttj+qkDxss5w
# RoPp2kChWTrZFQlXmVYwk/PJYczQCMxr7GJCkawCwO+k8IkRj3cCAwEAAaOCAzUw
# ggMxMA4GA1UdDwEB/wQEAwIHgDAMBgNVHRMBAf8EAjAAMBYGA1UdJQEB/wQMMAoG
# CCsGAQUFBwMIMIIBvwYDVR0gBIIBtjCCAbIwggGhBglghkgBhv1sBwEwggGSMCgG
# CCsGAQUFBwIBFhxodHRwczovL3d3dy5kaWdpY2VydC5jb20vQ1BTMIIBZAYIKwYB
# BQUHAgIwggFWHoIBUgBBAG4AeQAgAHUAcwBlACAAbwBmACAAdABoAGkAcwAgAEMA
# ZQByAHQAaQBmAGkAYwBhAHQAZQAgAGMAbwBuAHMAdABpAHQAdQB0AGUAcwAgAGEA
# YwBjAGUAcAB0AGEAbgBjAGUAIABvAGYAIAB0AGgAZQAgAEQAaQBnAGkAQwBlAHIA
# dAAgAEMAUAAvAEMAUABTACAAYQBuAGQAIAB0AGgAZQAgAFIAZQBsAHkAaQBuAGcA
# IABQAGEAcgB0AHkAIABBAGcAcgBlAGUAbQBlAG4AdAAgAHcAaABpAGMAaAAgAGwA
# aQBtAGkAdAAgAGwAaQBhAGIAaQBsAGkAdAB5ACAAYQBuAGQAIABhAHIAZQAgAGkA
# bgBjAG8AcgBwAG8AcgBhAHQAZQBkACAAaABlAHIAZQBpAG4AIABiAHkAIAByAGUA
# ZgBlAHIAZQBuAGMAZQAuMAsGCWCGSAGG/WwDFTAfBgNVHSMEGDAWgBQVABIrE5iy
# mQftHt+ivlcNK2cCzTAdBgNVHQ4EFgQUYVpNJLZJMp1KKnkag0v0HonByn0wfQYD
# VR0fBHYwdDA4oDagNIYyaHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0
# QXNzdXJlZElEQ0EtMS5jcmwwOKA2oDSGMmh0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNv
# bS9EaWdpQ2VydEFzc3VyZWRJRENBLTEuY3JsMHcGCCsGAQUFBwEBBGswaTAkBggr
# BgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMEEGCCsGAQUFBzAChjVo
# dHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURDQS0x
# LmNydDANBgkqhkiG9w0BAQUFAAOCAQEAnSV+GzNNsiaBXJuGziMgD4CH5Yj//7HU
# aiwx7ToXGXEXzakbvFoWOQCd42yE5FpA+94GAYw3+puxnSR+/iCkV61bt5qwYCbq
# aVchXTQvH3Gwg5QZBWs1kBCge5fH9j/n4hFBpr1i2fAnPTgdKG86Ugnw7HBi02JL
# sOBzppLA044x2C/jbRcTBu7kA7YUq/OPQ6dxnSHdFMoVXZJB2vkPgdGZdA0mxA5/
# G7X1oPHGdwYoFenYk+VVFvC7Cqsc21xIJ2bIo4sKHOWV2q7ELlmgYd3a822iYemK
# C23sEhi991VUQAOSK2vCUcIKSK+w1G7g9BQKOhvjjz3Kr2qNe9zYRDCCBs0wggW1
# oAMCAQICEAb9+QOWA63qAArrPye7uhswDQYJKoZIhvcNAQEFBQAwZTELMAkGA1UE
# BhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2lj
# ZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290IENBMB4X
# DTA2MTExMDAwMDAwMFoXDTIxMTExMDAwMDAwMFowYjELMAkGA1UEBhMCVVMxFTAT
# BgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEh
# MB8GA1UEAxMYRGlnaUNlcnQgQXNzdXJlZCBJRCBDQS0xMIIBIjANBgkqhkiG9w0B
# AQEFAAOCAQ8AMIIBCgKCAQEA6IItmfnKwkKVpYBzQHDSnlZUXKnE0kEGj8kz/E1F
# kVyBn+0snPgWWd+etSQVwpi5tHdJ3InECtqvy15r7a2wcTHrzzpADEZNk+yLejYI
# A6sMNP4YSYL+x8cxSIB8HqIPkg5QycaH6zY/2DDD/6b3+6LNb3Mj/qxWBZDwMiEW
# icZwiPkFl32jx0PdAug7Pe2xQaPtP77blUjE7h6z8rwMK5nQxl0SQoHhg26Ccz8m
# SxSQrllmCsSNvtLOBq6thG9IhJtPQLnxTPKvmPv2zkBdXPao8S+v7Iki8msYZbHB
# c63X8djPHgp0XEK4aH631XcKJ1Z8D2KkPzIUYJX9BwSiCQIDAQABo4IDejCCA3Yw
# DgYDVR0PAQH/BAQDAgGGMDsGA1UdJQQ0MDIGCCsGAQUFBwMBBggrBgEFBQcDAgYI
# KwYBBQUHAwMGCCsGAQUFBwMEBggrBgEFBQcDCDCCAdIGA1UdIASCAckwggHFMIIB
# tAYKYIZIAYb9bAABBDCCAaQwOgYIKwYBBQUHAgEWLmh0dHA6Ly93d3cuZGlnaWNl
# cnQuY29tL3NzbC1jcHMtcmVwb3NpdG9yeS5odG0wggFkBggrBgEFBQcCAjCCAVYe
# ggFSAEEAbgB5ACAAdQBzAGUAIABvAGYAIAB0AGgAaQBzACAAQwBlAHIAdABpAGYA
# aQBjAGEAdABlACAAYwBvAG4AcwB0AGkAdAB1AHQAZQBzACAAYQBjAGMAZQBwAHQA
# YQBuAGMAZQAgAG8AZgAgAHQAaABlACAARABpAGcAaQBDAGUAcgB0ACAAQwBQAC8A
# QwBQAFMAIABhAG4AZAAgAHQAaABlACAAUgBlAGwAeQBpAG4AZwAgAFAAYQByAHQA
# eQAgAEEAZwByAGUAZQBtAGUAbgB0ACAAdwBoAGkAYwBoACAAbABpAG0AaQB0ACAA
# bABpAGEAYgBpAGwAaQB0AHkAIABhAG4AZAAgAGEAcgBlACAAaQBuAGMAbwByAHAA
# bwByAGEAdABlAGQAIABoAGUAcgBlAGkAbgAgAGIAeQAgAHIAZQBmAGUAcgBlAG4A
# YwBlAC4wCwYJYIZIAYb9bAMVMBIGA1UdEwEB/wQIMAYBAf8CAQAweQYIKwYBBQUH
# AQEEbTBrMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQwYI
# KwYBBQUHMAKGN2h0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFz
# c3VyZWRJRFJvb3RDQS5jcnQwgYEGA1UdHwR6MHgwOqA4oDaGNGh0dHA6Ly9jcmwz
# LmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwOqA4oDaG
# NGh0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RD
# QS5jcmwwHQYDVR0OBBYEFBUAEisTmLKZB+0e36K+Vw0rZwLNMB8GA1UdIwQYMBaA
# FEXroq/0ksuCMS1Ri6enIZ3zbcgPMA0GCSqGSIb3DQEBBQUAA4IBAQBGUD7Jtygk
# pzgdtlspr1LPUukxR6tWXHvVDQtBs+/sdR90OPKyXGGinJXDUOSCuSPRujqGcq04
# eKx1XRcXNHJHhZRW0eu7NoR3zCSl8wQZVann4+erYs37iy2QwsDStZS9Xk+xBdIO
# PRqpFFumhjFiqKgz5Js5p8T1zh14dpQlc+Qqq8+cdkvtX8JLFuRLcEwAiR78xXm8
# TBJX/l/hHrwCXaj++wc4Tw3GXZG5D2dFzdaD7eeSDY2xaYxP+1ngIw/Sqq4AfO6c
# Qg7PkdcntxbuD8O9fAqg7iwIVYUiuOsYGk38KiGtSTGDR5V3cdyxG0tLHBCcdxTB
# nU8vWpUIKRAmMYIEOzCCBDcCAQEwgYYwcjELMAkGA1UEBhMCVVMxFTATBgNVBAoT
# DERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTExMC8GA1UE
# AxMoRGlnaUNlcnQgU0hBMiBBc3N1cmVkIElEIENvZGUgU2lnbmluZyBDQQIQCfyg
# 25Ky5Kv/46Pn+Tse6DAJBgUrDgMCGgUAoHgwGAYKKwYBBAGCNwIBDDEKMAigAoAA
# oQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4w
# DAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUbMpaKkPITk9FFos9ammW+iPI
# GKMwDQYJKoZIhvcNAQEBBQAEggEAYQvsWcvNCWid4nG+agwHozilNTjZ252lplUr
# VzNZh5OFuFqChU5YfkgSPVN0uuMGKZL8Jf2+8eYEkW2sIVU03ujP+qea6DClDUQD
# lvQ7PVgwTelSfBfMfIVjSBRrSMbPIwfWUYQZ4FbcFYCdewW6Y0k27/lBzy/lI2BW
# DeFdb13dvMHTdXTUfVPUbXIGTw3d3Yf+n2HVYDwpImjjS8NCQBPw4smp8jti52Fn
# 6zgzQb+PDhgAew/mdijXfdTtxbBHEYy+tyr/uQ2q6fZtbV+XZmtwN94d0wdnlqIL
# 514qN+b6Ub+ZWrYf/+Q+cOsdXEc71+eBuoP5pqO2nqz+XDjKnKGCAg8wggILBgkq
# hkiG9w0BCQYxggH8MIIB+AIBATB2MGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxE
# aWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMT
# GERpZ2lDZXJ0IEFzc3VyZWQgSUQgQ0EtMQIQAwGaAjr/WLFr1tXq5hfwZjAJBgUr
# DgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUx
# DxcNMTYwOTI5MjA1NjI2WjAjBgkqhkiG9w0BCQQxFgQUgpzflryXYvJf0ufD0GMP
# pNfyLa8wDQYJKoZIhvcNAQEBBQAEggEASAoxC0DtjoQC363dNrLdaaWfMSJ4FFW5
# vfIjrIS0DADbkSxnu17+uW0FOFbMN4Q7l5n2IxaG3GStVipBZFcWCD6yG67wr9qa
# DUqCjxvxrRbSmc0NS4rp384OTH1M1+yG7k4L5Om9ZK/E3bRrzIM6fUDDlfAVazkv
# HvtwG6/1yEnEI1p3JGqwJlw3B1nOpbwSrY4ov0p9SuFGFkd8BpDxOWHlwAAvr6vz
# 1GxqKwSLsORqtW5tBJxZIzCprcnYM0zXUGk354LIlZOcV2PnnGcJ9f7CHEc5axof
# Vft5eBIG1lYCrmYZJ8EN9SyEFwCswlr6v8b/kRh0X9Z+DudgZOcPEw==
# SIG # End signature block
