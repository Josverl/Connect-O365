<#
Initial version , Assuming x64 
#>
@{
    AdminComponents = @( 
        @{
            Tag=  "SIA"
            Name= "Microsoft Online Services Sign-In Assistant for IT Professionals"
            Source= "http://download.microsoft.com/download/5/0/1/5017D39B-8E29-48C8-91A8-8D0E4968E6D4/$Language/msoidcli_64.msi"
            Type= "MSI"
            ID= "{D8AB93B0-6FBF-44A0-971F-C0669B5AE6DD}"
        } , 
        @{ 
            Module=  "MSOnline"
            Tag=   "AAD"
            Name= "Windows Azure Active Directory Module for Windows PowerShell"
            Type= "MSI"
            Source= "https://bposast.vo.msecnd.net/MSOPMW/Current/amd64/AdministrationConfig-$Language.msi"
            ID= "{43CC9C53-A217-4850-B5B2-8C347920E500}"
        },
        @{
            Module= "SkypeOnlineConnector"
            Tag=  "SKYPE"
            Name= "Skype for Business Online, Windows PowerShell Module"
            Type= "EXE"
            Source= "https://download.microsoft.com/download/2/0/5/2050B39B-4DA5-48E0-B768-583533B42C3B/SkypeOnlinePowershell.exe"
            SetupOptions=  "/Install /Passive"
            ID= "{D7334D5D-0FA2-4DA9-8D8A-883F8C0BD41B}"
        },
        @{
            Module=  "Microsoft.Online.SharePoint.PowerShell"
            Tag=  "SPO"
            Name= "SharePoint Online Management Shell"
            Type= "MSI"
            Source=  "https://download.microsoft.com/download/0/2/E/02E7E5BA-2190-44A8-B407-BC73CA0D6B87/sharepointonlinemanagementshell_5111-1200_x64_$LangCountry.msi"
            Version=  "16.0.5111.1200"
            Web= "https://www.microsoft.com/en-us/download/confirmation.aspx?id=35588&6B49FDFB-8E5B-4B07-BC31-15695C5A2143=1"
            ID= "{95160000-115B-0409-1000-0000000FF1CE}"
        },
        @{
            Tag=  "RMS"
            Name= "Windows Azure AD Rights Management Administration"
            Type= "EXE"
            Module=  "aadrm"
            Source=  "https://download.microsoft.com/download/1/6/6/166A2668-2FA6-4C8C-BBC5-93409D47B339/WindowsAzureADRightsManagementAdministration_x64.exe"
            Version=  " 1.0.1443.901"
            Web= "https://www.microsoft.com/en-us/download/confirmation.aspx?id=30339"
            ID= "{6EACEC8B-7174-4180-B8D6-528D7B2C09F0}"
        },
        @{
            Tag=  "PNPPS"
            Name= "OfficeDevPnP.PowerShell"
            Type= "Module"
            Preview=  "Yes"
            Source=  "PSGallery"
            Web= "https://github.com/OfficeDev/PnP-PowerShell"
        }  
    )
    Exclude=  @(
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
            Web= "https://www.microsoft.com/en-us/download/confirmation.aspx?id=30339"
            ID= "{BE4B4004-DE97-4185-A2B4-C147DAC9AD2C}"
            Source2=  "https://download.microsoft.com/download/3/F/D/3FD04B49-26F9-4D9A-8C34-4533B9D5B020/W2K12-KB3066438-x64.msu"
        },
        @{
            Tag=  "SPO-4915"
            Module=  "Microsoft.Online.SharePoint.PowerShell"
            Name= "SharePoint Online Management Shell"
            Source=  "https://download.microsoft.com/download/0/2/E/02E7E5BA-2190-44A8-B407-BC73CA0D6B87/sharepointonlinemanagementshell_4915-1200_x64_$LangCountry.msi"
            Type= "MSI"
            Version=  "16.0.4915.1200"
            Web= "https://www.microsoft.com/en-us/download/confirmation.aspx?id=35588&6B49FDFB-8E5B-4B07-BC31-15695C5A2143=1"
            ID= "{95160000-115B-0409-1000-0000000FF1CE}"
        }

      ) 
}       

