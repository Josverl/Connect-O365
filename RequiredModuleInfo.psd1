<#
 # Configuration file to define the admin modules used by the Connect-O365 module and script

    Variables that can be used (need to be defined prior to calling import-datafile ): 
    - $language         EN      NL 
    - $LangCountry      EN-US   NL-NL 
    - $cpu              AMD64   x86
    - $xcpu             x64   x86
    - $bitness          64      32
#>
write-host -f green "Required Module Info for Connect-O365 script. Configuration Version 1.8"
write-host -f green "Platform : $xcpu, Language : $language ; $LangCountry"
@{
    Version = "1.7";
    AdminComponents = @( 
        @{
            Tag=  "SIA"
            Name= "Microsoft Online Services Sign-In Assistant for IT Professionals"
            Source= "http://download.microsoft.com/download/5/0/1/5017D39B-8E29-48C8-91A8-8D0E4968E6D4/$Language/msoidcli_$bitness.msi"
            Type= "MSI"
            ID= "{D8AB93B0-6FBF-44A0-971F-C0669B5AE6DD}"

            Repository = "Downloads"
            #To find downloads
            Web= "https://www.microsoft.com/$LangCountry/download/details.aspx?id="
            DownloadID = '41950'
            Filter = "*_$bitness*"
        } , 

        @{ 
            Tag=   "AADv2"
            Name= "Azure Active Directory V2 PowerShell Module"
            Type= "Module"
            Repository= "PSGallery"
            Module=  "AzureAD"
            Web = "https://www.powershellgallery.com/packages/AzureAD/"
        },

        @{
            Module= "SkypeOnlineConnector"
            Tag=  "SKYPE"
            Name= "Skype for Business Online, Windows PowerShell Module"
            Type= "EXE"
            Source= "https://download.microsoft.com/download/2/0/5/2050B39B-4DA5-48E0-B768-583533B42C3B/SkypeOnlinePowershell.exe"
            SetupOptions=  "/Install /Passive"
            ID= "{D7334D5D-0FA2-4DA9-8D8A-883F8C0BD41B}"
            
            Repository = "Downloads"
            #To find downloads
            Web= "https://www.microsoft.com/$LangCountry/download/details.aspx?id=39366"
            DownloadID = '39366'
            Filter = '*'
        },
        @{
            Module=  "Microsoft.Online.SharePoint.PowerShell"
            Tag=  "SPO"
            Name= "SharePoint Online Management Shell"
            Type= "MSI"
            #need additional escaping for path expansion
            Source =  "https://download.microsoft.com/download/0/2/E/02E7E5BA-2190-44A8-B407-BC73CA0D6B87/sharepointonlinemanagementshell_6112-1200_$($XCPU)_$($LangCountry).msi"
            Version=  "16.0.5625.1200"
            ID= "{95160000-115B-0409-1000-0000000FF1CE}"

            Repository = "Downloads"
            #To find downloads
            Web= "https://www.microsoft.com/$LangCountry/download/details.aspx?id=35588"
            DownloadID = '35588'
            Filter = "*$xcpu*"

        },
        @{
            Tag=  "RMS"
            Name= "Azure Rights Management Administration Tool"
            Type= "EXE"
            Module=  "aadrm"
            Source= "https://download.microsoft.com/download/1/6/6/166A2668-2FA6-4C8C-BBC5-93409D47B339/WindowsAzureADRightsManagementAdministration_$xcpu.exe"  
            Version=  "2.4.0.0"
            
            Documentation = "https://msdn.microsoft.com/$LangCountry/library/azure/dn629398.aspx"
            ID= "{6EACEC8B-7174-4180-B8D6-528D7B2C09F0}"

            Repository = "Downloads"
            #To find downloads
            Web= "https://www.microsoft.com/$LangCountry/download/confirmation.aspx?id=30339"
            DownloadID = '30339'
            Filter = "*$xcpu*"
        },
        @{
            Tag=  "PNPPS"
            Name= "OfficeDevPnP.PowerShell"
            Type= "Module"
            Repository=  "PSGallery"
            Module= "SharepointPnPPowerShellOnline"
            Web= "https://github.com/OfficeDev/PnP-PowerShell"
        }  
    )
    Exclude=  @(
         @{ 
            Tag=   "AADv1"
            Name= "Windows Azure Active Directory Module for Windows PowerShell"
            Module=  "MSOnline V1"
            Type= "MSI"
            Source = "http://download.connect.microsoft.com/pr/AdministrationConfig_3.msi?t=2b52abb2-efe7-4ce1-92a9-a18aef766642&e=1488055536&h=38d0d68eb56e50766627cf52d24761a2"
            #Source= "https://bposast.vo.msecnd.net/MSOPMW/Current/$cpu/AdministrationConfig-$Language.msi"
            ID= "{43CC9C53-A217-4850-B5B2-8C347920E500}"
            Web = "https://msdn.microsoft.com/$LangCountry/library/jj151815.aspx?f=255&MSPPError=-2147217396"
            #Web= "https://www.microsoft.com/$LangCountry/download/details.aspx?id="
            DownloadID = ''
            Filter = ''
        },
        @{
            Tag=  "WMF5-Preview"
            Preview=  "Yes"
            Module=  "WMF5"
            Type= "MSI"
            Name= "Windows Management Framework 5.0 Production Preview"
            Source=  "https://download.microsoft.com/download/3/F/D/3FD04B49-26F9-4D9A-8C34-4533B9D5B020/Win8.1AndW2K12R2-KB3066437-x64.msu"
            SetupOptions=  "/quiet"
            OS= "6,8"
            XVersion=  ""
            Web= "https://www.microsoft.com/$LangCountry/download/confirmation.aspx?id=30339"
            ID= "{BE4B4004-DE97-4185-A2B4-C147DAC9AD2C}"
        },
        @{
            Tag=  "SPO-4915"
            Module=  "Microsoft.Online.SharePoint.PowerShell"
            Name= "SharePoint Online Management Shell"
            Source=  "https://download.microsoft.com/download/0/2/E/02E7E5BA-2190-44A8-B407-BC73CA0D6B87/sharepointonlinemanagementshell_4915-1200_x64_$LangCountry.msi"
            Type= "MSI"
            Version=  "16.0.4915.1200"
            Web= "https://www.microsoft.com/$LangCountry/download/confirmation.aspx?id=35588&6B49FDFB-8E5B-4B07-BC31-15695C5A2143=1"
            ID= "{95160000-115B-0409-1000-0000000FF1CE}"
        }
    ) 
}       

