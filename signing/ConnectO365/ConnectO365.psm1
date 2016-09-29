



    <#
    .Synopsis
       Retrieve credentials using the UI and store these in a file in the userprofile\creds folder
       the credentials are also returned.
    .EXAMPLE
       Store-MyCreds -UserName Admin@contoso.com
    #>
    $script:MSG_Cred = 'Please enter the Tenant Admin or Service Admin password'
    $script:MSG_CredCancel = 'No password entered or user canceled'
    function global:Store-myCreds {
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
    function script:Test-myCreds {
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
    function global:Get-myCreds {
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
                $admincredentials  = Store-myCreds $Account
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
        retrieve-credentials -account admin@contso.com 
    .EXAMPLE
        retrieve-credentials -account admin@contso.com -persist
    .EXAMPLE
        #retrieve a credentian using a alias from the credential manager
        retrieve-credentials -account Production

    #>
    function global:retrieve-credentials {
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
                    $admincredentials= Get-Credential -UserName $name -Message $Script:MSG_Cred
                }


                

            }
        }
        return $admincredentials
    }



# SIG # Begin signature block
# MIIgNAYJKoZIhvcNAQcCoIIgJTCCICECAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUVZg30epkOVDN6HczsvVN7b0T
# aH6gghtjMIIDtzCCAp+gAwIBAgIQDOfg5RfYRv6P5WD8G/AwOTANBgkqhkiG9w0B
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
# DAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUSIBsq9f4rEFJBWSJSZgAmXpA
# 8pswDQYJKoZIhvcNAQEBBQAEggEAw0q8c7W+vd0/E2o832lLWWWKO9Yr2diVNFA1
# C1oMO1uT0xcpgn3ZG6ZHQc6GwEmgGumXn+Fqyj4fnP1H7wC08XnX07WhCr/hV2ui
# hmva7pYNviXi9Sqr9E+yFzUisna3XVUa+wkH1QzrghhhcIBC9H135390NrKHwns4
# caqeow8Mvt6fgLdtWVrlR1UWH5MK6VtrN3/m0TXwgiXXOMlRe3rXonj7mtFOKhcd
# DXXVBvrg45FXdfxTHEcwtK2LET8BzB9wA4wySoD4RDRBEXEDvGDo2rzAoK7x4774
# OTRoDnY9+Gbrr9Y6kLaYNxT5ELa8LgzGT4lbxyI4n0tnr2aHl6GCAg8wggILBgkq
# hkiG9w0BCQYxggH8MIIB+AIBATB2MGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxE
# aWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMT
# GERpZ2lDZXJ0IEFzc3VyZWQgSUQgQ0EtMQIQAwGaAjr/WLFr1tXq5hfwZjAJBgUr
# DgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUx
# DxcNMTYwOTI5MjA1NTAxWjAjBgkqhkiG9w0BCQQxFgQUwlaTmSVQDhDZo20jayB1
# Q/8JjMUwDQYJKoZIhvcNAQEBBQAEggEAcKrRa+KMrnIcxuHJIkPWLdAQf77KUQjb
# woRFJfjaaEHMD+BKzrjxCTtoQA35cP1cE0CbhpIf6t8aMMb5dUNkTGLC/P9emlgk
# GIKpR6XyD3Mqp74IwsNKbwDi0dwG72H8myPaFK7ZsqPJfP82RklauWJY2Vf0wIT1
# 3edfptPf2+hW76q9LxskGPmr4D/sCVqfNtKE4a1PATTOcUZDJrmxCl1G4sEWFJBn
# cUrr6ByHcGxkGGkvcxAAMHFbabo2cHUaF1nF2cgUon2WhTO9RXSZTzRKKsBpJ6ML
# U5f5dgPIvVM1bJx80aRDjLHadQ6EJg+I6PIjJ7CHxtOHxSBo+0QhmA==
# SIG # End signature block
