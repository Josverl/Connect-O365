



    <#
    .Synopsis
       Retrieve credentials using the UI and store these in a file in the userprofile\creds folder
       the credentials are also returned.
    .EXAMPLE
       Set-myCreds -UserName Admin@contoso.com
    #>
    $script:MSG_Cred = 'Please enter the Tenant Admin or Service Admin password'
    $script:MSG_CredCancel = 'No password entered or user canceled'
    function Set-myCreds {
    Param (
            # The Account or Username 
            [Parameter()]
            [ValidateNotNullOrEmpty()]
            [Alias("Username")]  # Backward compat with v1.6.7 and older 
            [string]$Account
    )
        $Credential = Get-Credential -Account $Account -Message $Script:MSG_Cred
        if ($Credential) { 
            $Store = "$env:USERPROFILE\creds\$Account.txt"
            MkDir "$env:USERPROFILE\Creds" -ea 0 | Out-Null
            $Credential.Password | ConvertFrom-SecureString | Set-Content $store
            Write-Verbose "Saved credentials to $store"
        } else {
            write-warning $script:MSG_CredCancel
        }    
        return $Credential 
     }
    
    <#
    .Synopsis
       Test if credentials for a specific username are stored in the \creds folder
    .EXAMPLE
       if ( Test-MyCreds -UserName Admin@contoso.com ) { "credentials found" }
    #>
    function Test-myCreds {
    param( 
            # The Account or Username 
            [Parameter()]
            [ValidateNotNullOrEmpty()]
            [Alias("Username")]  # Backward compat with v1.6.7 and older 
            [string]$Account
        )
        $Store = "$env:USERPROFILE\creds\$Account.txt"
        return (Test-Path $store)
    }
    <#
    .Synopsis
       retrieve credentials 
       -Persist indicates that the credentials should be saved 
       -Force   indicates that the password should be re-entered by the user 
    .EXAMPLE
       # retrieve the stored credentials, if not present just prompt for the password 
       Get-MyCreds -UserName Admin@contoso.com
   
    .EXAMPLE
       # store the credentials for future re-use, overwrites any existing credentials
       Get-MyCreds -UserName Admin@contoso.com -persist

    #>

    <#
    .Synopsis
       Short description
    .DESCRIPTION
       Long description
    .EXAMPLE
       Example of how to use this cmdlet
    .EXAMPLE
       Another example of how to use this cmdlet
    #>
    function Get-myCreds {
        Param
        (
            # The Account or Username 
            [Parameter()]
            [ValidateNotNullOrEmpty()]
            [Alias("Username")]  # Backward compat with v1.6.7 and older 
            [string]$Account,
            # Persist username and password 
            [switch] $Persist
        )
        $Store = "$env:USERPROFILE\creds\$Account.txt"
        if ( (Test-Path $store) -AND  $Persist -eq $false  ) {
            #use a stored password if found , unless -persist/-force is used to ask for and store a new password
            Write-Verbose "Retrieved credentials from $store"
            $Password = Get-Content $store | ConvertTo-SecureString
            $Credential = New-Object System.Management.Automation.PsCredential($Account,$Password)
            return $Credential
        } else {
            if ($persist -and -not [string]::IsNullOrEmpty($Account)) {
                WRITE-VERBOSE 'Ask and store new credentials'
                $admincredentials  = Set-myCreds $Account
                return $admincredentials
            } else {
                WRITE-VERBOSE 'Ask for credentials'
                return Get-Credential -Credential $Account 
            }
        }
     }

    <#
    .Synopsis
       Retrieves credentials that are stored either in the \creds folder, or in the windows storedcredentials 
       Windows stored credentials depend on an external module to be in installed (CredentialManager) 
    .EXAMPLE
        RetrieveCredentials -account admin@contso.com 
    .EXAMPLE
        RetrieveCredentials -account admin@contso.com -persist
    .EXAMPLE
        #retrieve a credentian using a alias from the credential manager
        RetrieveCredentials -account Production

    #>
    function RetrieveCredentials {
        Param
        (
            # The Account or Username 
            [Parameter()]
            [ValidateNotNullOrEmpty()]
            [string]$Account,
            [switch]$Persist
        )
        $admincredentials = $null
        #if credentials are stored in the filestore 
        if (test-myCreds  $account) {
            write-verbose 'Find credentials from credential folder'
            $admincredentials = Get-myCreds $account -Persist:$Persist 
        } else { 

            #check if the credentialmanager module is installed 
            $CM = get-module credentialmanager -ListAvailable | select -Last 1
            if ($cm -ne $null -and $CM.Version -eq "2.0") {
                write-verbose 'Find credentials stored in the credential manager'
                #Find the credentials stored in the credential manager
                #check match on target name
                $stored = Get-StoredCredential -Type GENERIC -Target  $account -AsCredentialObject| select -First 1
                #otherwise check on username
                if ($stored -eq $null) {
                    write-verbose 'Find credentials based on user name'
                    $credentials = Get-StoredCredential -Type GENERIC -AsCredentialObject
                    #work around pipeline constraints in get-stored 
                    $credentials = $credentials | where { $_.UserName -like '?*@?*' -and $_.Type -eq 'GENERIC'} | select -Property UserName, TargetName, Type, TargetAlias, Comment
                    $stored = $credentials | where {$_.UserName-ieq $account} | select -First 1
                }
                
                if ($persist) {
                    write-verbose 'Asking for a new password'
                    #if -Persist is specified we need to ask for a new password and update the stored password
                    if ($stored) { $name= $stored.Username } else { $name=$account}
                    $newCred = Get-Credential -UserName $name -Message $Script:MSG_Cred
                    if ($newCred -eq $null) {
                        write-warning $script:MSG_CredCancel
                    } else {
                        if ($stored) {
                            write-verbose 'Update existing Stored Credential'
                            $stored = New-StoredCredential -Comment "Connect-O365" -Password $newCred.GetNetworkCredential().Password -Persist ENTERPRISE -Target $stored.TargetName -Type GENERIC -UserName $newcred.UserName 
                        } else {
                            write-verbose 'Create New Stored Credential'
                            $stored = New-StoredCredential -Comment "Connect-O365" -Password $newCred.GetNetworkCredential().Password -Persist ENTERPRISE -Target $newcred.UserName -Type GENERIC -UserName $newcred.UserName 
                        }
                    }
                }
                #If a stored cred was found
                if ($stored -ne $null) {
                    write-verbose "Retrieving Target : $($stored.Targetname)"
                    $admincredentials = Get-StoredCredential -Target $stored.Targetname -Type 'GENERIC'
                } else {
                    #If not found, and if no -Persist then old fashioned 
                    write-verbose "Ask for credential"
                    $admincredentials = Get-Credential -UserName $Account -Message $Script:MSG_Cred 
                }
            }
        }
        #write-verbose "Cred : $($admincredentials.UserName)" -Verbose
        return $admincredentials
    }

