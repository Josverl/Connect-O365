<#
 #  use Invoke-Psake to start the build
 #  v0.7 more robust test-install
 #  v0.6 improve logic for Signing, 
        - avoid getting stuck on malformed *.psd1 ?
        - Display type and Module 

 #  v0.5 Add logic to use current folder if nothing else specified
 #  v0.4 Add logic to deal with Scripts and Modules
     - Support Module
     - Script

#>


Task default -Depends TestInstall

Properties {
    # The name of your module should match the basename of the PSD1 file.

    # Module 
    # ------
    # BasePath        c:\dev\foo
    # Modulename      foo
    # PublishRootDir  c:\dev\foo\Release
    # ReleaseDir      c:\dev\foo\Release\foo

    # Script 
    # ------
    # BasePath        c:\dev\foo
    # Modulename      Null
    # PublishRootDir  c:\dev\foo\Release
    # ReleaseDir      c:\dev\foo\Release

    
    if ($PSScriptRoot ) { 
        $BasePath = $PSScriptRoot 
    } else {
        #Handle run in ISE with open file
        $BasePath = split-path -parent $psISE.CurrentFile.Fullpath
    }
    if ([string]::IsNullOrEmpty($BasePath)) {
        Write-Verbose "Using the Working Directory as Base" -Verbose
        $BasePath =$pwd
    }
    #$Modules = @(Get-Item $BasePath\*.psd1 | Foreach-Object {$null = Test-ModuleManifest -Path $_ -ErrorAction SilentlyContinue; if ($?) {$_} else {Write-Warning $Error[0]}})
    $Modules = @(Get-Item $BasePath\*.psd1 | Foreach-Object {$null = Test-ModuleManifest -Path $_ ; if ($?) {$_} else {Write-Warning $Error[0]}})

    $Target=@{Type = "";Name ="" }

    if ($Modules.Count -gt 0) {
        $Target.Type = "Module"
        $Target.Name = $Modules[0].BaseName
        $Target.BaseName = $Modules[0].BaseName
    } else {
        Write-verbose "No modules found, looking for a script"
        #work around strange behaviour test-scriptFileInfo

        $scripts = @(Get-Item $BasePath\*.ps1|Foreach {
            Try { 
                $null =Test-ScriptFileInfo -Path $_ -ErrorAction SilentlyContinue; 
            } catch {}
            if ($?) {$_}})


        $Target.Type = "Script"
        $Target.Name = $Scripts[0].Name
        $Target.BaseName = $Scripts[0].BaseName
        Write-Verbose $Target.Name  -Verbose
    }

    If ($Target.Type -ieq "Module") {
        $ModuleName = $Target.Name
    } else {
        $ModuleName = $Null
    }
    Write-Host -ForegroundColor Green ("Building {0} : {1}" -f $Target.Type, $Target.Name )
    # The directory used to publish the module from.  If you are using Git, the
    # $PublishRootDir should be ignored if it is under the workspace directory.
    $PublishRootDir = join-path $BasePath 'Release'
    $ReleaseDir     =  join-path $PublishRootDir $ModuleName

    # The following items will not be copied to the $ReleaseDir.
    # Add items that should not be published with the module.
    $Exclude = @(
        'Release','Tests','.git*','.vscode','launch.json','*.tests.ps1',
        #donot copy nuget ballast
        'net40','tools','*.nupkg','Microsoft.ApplicationInsights.xml',
        #donot copy dev and build artefacts
        'scratch','build.ps1','default.ps1'
    )

    $TestRepository = "Dev" #$null
    # Name of the repository you wish to publish to. Default repo is the PSGallery.
    $PublishRepository = "PSGallery" #$null

    # Your NuGet API key for the PSGallery.  
    $NuGetApiKey = $null
    if ($NuGetApiKey -eq $Null) {
        $Creds = Get-StoredCredential  -Target 'PSGallery:NuGetApiKey'
        if ($Creds) {
            $NuGetApiKey = $creds.GetNetworkCredential().Password
        }
    }
}


FormatTaskName "|>-------- {0} --------<|"

Task Test  {
 
    Import-Module Pester
    $Results = Invoke-Pester -PassThru

    if  ($Results.FailedCount -gt 0) {
 #       Throw "Testing Failed"
    }
}

Task Clean  -requiredVariables PublishRootDir `
            -description "Clean the Release Folder" `
            -Depends Test {
    # Sanity check the dir we are about to "clean".  If $PublishRootDir were to
    # inadvertently get set to $null, the Remove-Item commmand removes the
    # contents of \*.  
    Write-verbose "Clean : $PublishRootDir" -Verbose
    if ((Test-Path $PublishRootDir) -and $PublishRootDir.Contains($BasePath)) {
        Remove-Item $PublishRootDir\* -Recurse -Force
    }
}

Task Copy   -description "Copy items to the release folder" `
            -Depends Clean `
            -requiredVariables BasePath, ReleaseDir, Exclude, Target {
    if ($target.Type -ieq "Module" ){

        Write-verbose "Copy Module: $BasePath --> $ReleaseDir" -Verbose              
        MD $ReleaseDir -ErrorAction SilentlyContinue | Out-Null
        Copy-Item -Path $BasePath\*.* -Destination $ReleaseDir -Recurse -Exclude $Exclude 

     } else {
        Write-verbose "Copy Script: $BasePath --> $ReleaseDir" -Verbose              
        MD $ReleaseDir -ErrorAction SilentlyContinue | Out-Null
        Copy-Item -Path (Join-path $BasePath $target.Name ) -Destination $ReleaseDir -Exclude $Exclude 
    }

}

Task Sign   -Depends Copy `
            -RequiredVariables ReleaseDir, Target {
   "Sign"
    
    #Just get the first codesigning cert 
    $CodeSigningCerts = @(gci cert:\currentuser\my -codesigning)

     
    if ($CodeSigningCerts.Count -ge 1)  {
        $SigningCert = $CodeSigningCerts[0]
        if ($target.Type -ieq "Module" ){
            #ToDo SIgn Module

        } else {
        
            $ScriptName = Join-Path -Path $ReleaseDir -ChildPath  $target.Name
            Write-Host -f Green "Signing " $target.Name
            #Sign the script 
            $sig = Set-AuthenticodeSignature -FilePath $ScriptName -Certificate $cert -IncludeChain "All" -TimeStampServer "http://timestamp.digicert.com/scripts/timstamper.dll" # #"http://timestamp.digicert.com/scripts/timstamp.dll"
        
            # TODO TimeStampCertificate does not get added
        }
    } else {
        Write-Warning "No Signing certificate; Signing is skipped " 
    }
}

Task TestPublish -Depends Sign `
                 -RequiredVariables ReleaseDir, Target, TestRepository {

    $publishParams = @{} 
    $publishParams['Repository'] = $TestRepository

    if ($target.Type -ieq "Module" ){
        #remove the same version form the test repo, if it already exists
        $MFT = Test-ModuleManifest -Path (Join-Path $ReleaseDir -ChildPath "$moduleName.psd1") 
        $filter = "{0}.{1}.nupkg" -f $target.Name , $MFT.Version.ToString()
        Get-ChildItem -path ( (Get-PSRepository -Name "$TestRepository").SourceLocation) -Filter $filter| remove-item

        # Consider not using -ReleaseNotes parameter when Update-ModuleManifest has been fixed.
        if ($ReleaseNotesPath) {
            $publishParams['ReleaseNotes'] = @(Get-Content $ReleaseNotesPath)
        }
        $publishParams['Path']= $ReleaseDir
        "Calling Publish-Module..."
        Publish-Module @publishParams 
    } else{
        #remove the same version form the test repo, if it already exists
        $MFT = Test-ScriptFileInfo -Path (Join-Path $ReleaseDir -ChildPath $target.Name ) 
        $filter = "{0}.{1}.nupkg" -f $target.BaseName , $MFT.Version.ToString()
        Get-ChildItem -path ( (Get-PSRepository -Name "$TestRepository").SourceLocation) -Filter $filter | remove-item

        "Calling Publish-Script..."
        $publishParams['Path']= Join-Path $ReleaseDir $target.Name

        Publish-Script @publishParams    

    }

}


Task TestInstall -Depends TestPublish{
   "Test Install"
    if ($target.Type -ieq "Module" ){
        $MFT = Test-ModuleManifest -Path (Join-Path $ReleaseDir -ChildPath "$moduleName.psd1") 

        #Remove the same version from the local machine if allready installed
        $Installed = @(get-InstalledModule -Name $mft.Name -RequiredVersion $mft.version -ErrorAction SilentlyContinue)
        if ($Installed.Count -Gt 0) {
            Uninstall-Module -Name $mft.Name -RequiredVersion $mft.version 
        }
        

        #find it in the repo 
        find-Module -Name $mft.Name -RequiredVersion $mft.version -Repository $TestRepository

        #install it form the repo 
        install-Module -Name $mft.Name -RequiredVersion $mft.version -Repository $TestRepository -Force 
        Get-InstalledModule -Name $mft.Name

    } else {
        $MFT = Test-ScriptFileInfo -Path (Join-Path $ReleaseDir -ChildPath $target.Name ) 

        find-script -Name $mft.Name -RequiredVersion $mft.version -Repository $TestRepository
        install-script -Name $mft.Name -RequiredVersion $mft.version -Repository $TestRepository -Force 
        Get-InstalledScript -Name $mft.Name 

    }

}

Task Publish -Depends TestInstall {


    $publishParams = @{} 

    $publishParams['NuGetApiKey'] = $NuGetApiKey
    $publishParams['Repository'] = $PublishRepository

    if ($target.Type -ieq "Module" ){
        #remove the same version form the test repo, if it already exists
        $MFT = Test-ModuleManifest -Path (Join-Path $ReleaseDir -ChildPath "$moduleName.psd1") 
        $filter = "{0}.{1}.nupkg" -f $target.Name , $MFT.Version.ToString()
        Get-ChildItem -path ( (Get-PSRepository -Name "$TestRepository").SourceLocation) -Filter $filter| remove-item

        # Consider not using -ReleaseNotes parameter when Update-ModuleManifest has been fixed.
        if ($ReleaseNotesPath) {
            $publishParams['ReleaseNotes'] = @(Get-Content $ReleaseNotesPath)
        }
        $publishParams['Path']= $ReleaseDir
        "Calling Publish-Module..."
        Publish-Module @publishParams 
    } else{
        #remove the same version form the test repo, if it already exists
        $MFT = Test-ScriptFileInfo -Path (Join-Path $ReleaseDir -ChildPath $target.Name ) 
        $filter = "{0}.{1}.nupkg" -f $target.BaseName , $MFT.Version.ToString()
        Get-ChildItem -path ( (Get-PSRepository -Name "$TestRepository").SourceLocation) -Filter $filter | remove-item

        "Calling Publish-Script..."
        $publishParams['Path']= Join-Path $ReleaseDir $target.Name

        Publish-Script @publishParams    
    }
}
