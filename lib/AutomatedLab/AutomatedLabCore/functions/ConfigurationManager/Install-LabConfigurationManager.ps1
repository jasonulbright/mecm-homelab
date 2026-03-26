function Install-LabConfigurationManager
{
    [CmdletBinding()]
    param ()

    $vms = Get-LabVm -Role ConfigurationManager
    Start-LabVm -Role ConfigurationManager -Wait

    #region Prereq: ADK, CM binaries, stuff
    Write-ScreenInfo -Message "Installing Prerequisites on $($vms.Count) machines"

    # ADK setup: use existing offline layouts if available, otherwise download and create them
    $adkOfflinePath = Join-Path $labSources 'SoftwarePackages/ADKoffline'
    $adkPeOfflinePath = Join-Path $labSources 'SoftwarePackages/ADKPEoffline'
    $adkBootstrapper = Join-Path $labSources 'SoftwarePackages/ADK/adksetup.exe'
    $adkPeBootstrapper = Join-Path $labSources 'SoftwarePackages/ADKPE/adkwinpesetup.exe'

    # Create offline layouts if they don't exist
    if (-not (Test-Path (Join-Path $adkOfflinePath 'Installers'))) {
        Write-ScreenInfo -Message 'Creating ADK offline layout (this may take a few minutes)...'
        if (-not (Test-Path $adkBootstrapper)) {
            $adkUrl = Get-LabConfigurationItem -Name WindowsAdk
            $adkFile = Get-LabInternetFile -Uri $adkUrl -Path "$labSources/SoftwarePackages/ADK" -FileName 'adksetup.exe' -PassThru -NoDisplay
            $adkBootstrapper = $adkFile.FullName
        }
        Start-Process -FilePath $adkBootstrapper -ArgumentList "/quiet /layout `"$adkOfflinePath`"" -Wait -NoNewWindow
    }

    if (-not (Test-Path (Join-Path $adkPeOfflinePath 'Installers'))) {
        Write-ScreenInfo -Message 'Creating ADK PE offline layout...'
        if (-not (Test-Path $adkPeBootstrapper)) {
            $adkPeUrl = Get-LabConfigurationItem -Name WindowsAdkPe
            $adkPeFile = Get-LabInternetFile -Uri $adkPeUrl -Path "$labSources/SoftwarePackages/ADKPE" -FileName 'adkwinpesetup.exe' -PassThru -NoDisplay
            $adkPeBootstrapper = $adkPeFile.FullName
        }
        Start-Process -FilePath $adkPeBootstrapper -ArgumentList "/quiet /layout `"$adkPeOfflinePath`"" -Wait -NoNewWindow
    }

    # Ensure bootstrapper is inside the offline layout folder (needed for VM-side install)
    if ((Test-Path $adkBootstrapper) -and -not (Test-Path (Join-Path $adkOfflinePath 'adksetup.exe'))) {
        Copy-Item $adkBootstrapper $adkOfflinePath -Force
    }
    if ((Test-Path $adkPeBootstrapper) -and -not (Test-Path (Join-Path $adkPeOfflinePath 'adkwinpesetup.exe'))) {
        Copy-Item $adkPeBootstrapper $adkPeOfflinePath -Force
    }

    # Set up VM install directory
    $deployDebugPath = Invoke-LabCommand -ComputerName $vms -ScriptBlock {
        (New-Item -ItemType Directory -Path $ExecutionContext.InvokeCommand.ExpandString($AL_DeployDebugFolder) -ErrorAction SilentlyContinue -Force).FullName
    } -PassThru -Variable (Get-Variable -Name AL_DeployDebugFolder -Scope Global) | Select-Object -First 1

    # Copy layouts to VM
    if ($(Get-Lab).DefaultVirtualizationEngine -eq 'Azure')
    {
        Install-LabSoftwarePackage -Path (Join-Path $adkOfflinePath 'adksetup.exe') -ComputerName $vms -CommandLine "/quiet /layout `"$deployDebugPath\ADKoffline`"" -NoDisplay
        Install-LabSoftwarePackage -Path (Join-Path $adkPeOfflinePath 'adkwinpesetup.exe') -ComputerName $vms -CommandLine "/quiet /layout `"$deployDebugPath\ADKPEoffline`"" -NoDisplay
    }
    else
    {
        Copy-LabFileItem -Path $adkOfflinePath -ComputerName $vms -DestinationFolderPath $deployDebugPath -Recurse
        Copy-LabFileItem -Path $adkPeOfflinePath -ComputerName $vms -DestinationFolderPath $deployDebugPath -Recurse

        # Flatten nested folders from Copy-LabFileItem
        Invoke-LabCommand -ComputerName $vms -ActivityName 'Flatten ADK folders' -ScriptBlock {
            param($basePath)
            foreach ($name in @('ADKoffline', 'ADKPEoffline')) {
                $target = Join-Path $basePath $name
                $nested = Join-Path $target $name
                if (Test-Path $nested) {
                    Get-ChildItem $nested | Move-Item -Destination $target -Force
                    Remove-Item $nested -Recurse -Force -ErrorAction SilentlyContinue
                }
            }
        } -ArgumentList $deployDebugPath
    }

    Install-LabSoftwarePackage -LocalPath "$deployDebugPath\ADKoffline\adksetup.exe" -ComputerName $vms -CommandLine '/norestart /q /ceip off /features OptionId.DeploymentTools OptionId.UserStateMigrationTool OptionId.ImagingAndConfigurationDesigner' -NoDisplay
    Install-LabSoftwarePackage -LocalPath "$deployDebugPath\ADKPEoffline\adkwinpesetup.exe" -ComputerName $vms -CommandLine '/norestart /q /ceip off /features OptionId.WindowsPreinstallationEnvironment' -NoDisplay

    $ncliUrl = Get-LabConfigurationItem -Name SqlServerNativeClient2012
    try
    {
        $ncli = Get-LabInternetFile -Uri $ncliUrl -Path "$labSources/SoftwarePackages" -FileName sqlncli.msi -ErrorAction "Stop" -ErrorVariable "GetLabInternetFileErr" -PassThru
    }
    catch
    {
        $Message = "Failed to download SQL Native Client from '{0}' ({1})" -f $ncliUrl, $GetLabInternetFileErr.ErrorRecord.Exception.Message
        Write-LogFunctionExitWithError -Message $Message
    }

    Install-LabSoftwarePackage -Path $ncli.FullName -ComputerName $vms -CommandLine "/qn /norestart IAcceptSqlncliLicenseTerms=Yes" -ExpectedReturnCodes 0

    #region VC++ 14.50 runtimes (required by MSOLEDB 19)
    Write-ScreenInfo -Message 'Installing VC++ runtimes (14.50+)'
    $vcx64 = Join-Path $labSources 'SoftwarePackages/VCRedist/vc_redist.x64.exe'
    $vcx86 = Join-Path $labSources 'SoftwarePackages/VCRedist/vc_redist.x86.exe'
    if (-not (Test-Path $vcx64)) {
        Write-ScreenInfo -Message 'Downloading VC++ x64 runtime...'
        $vcx64 = (Get-LabInternetFile -Uri 'https://aka.ms/vs/18/release/vc_redist.x64.exe' -Path "$labSources/SoftwarePackages/VCRedist" -FileName 'vc_redist.x64.exe' -PassThru -ErrorAction Stop).FullName
    }
    if (-not (Test-Path $vcx86)) {
        Write-ScreenInfo -Message 'Downloading VC++ x86 runtime...'
        $vcx86 = (Get-LabInternetFile -Uri 'https://aka.ms/vs/18/release/vc_redist.x86.exe' -Path "$labSources/SoftwarePackages/VCRedist" -FileName 'vc_redist.x86.exe' -PassThru -ErrorAction Stop).FullName
    }
    Install-LabSoftwarePackage -Path $vcx64 -ComputerName $vms -CommandLine '/quiet /norestart' -ExpectedReturnCodes 0, 3010
    Install-LabSoftwarePackage -Path $vcx86 -ComputerName $vms -CommandLine '/quiet /norestart' -ExpectedReturnCodes 0, 3010
    #endregion

    #region ODBC Driver 18.5.x (required by CM 2509+, NOT 18.6.x which has NULL regression)
    Write-ScreenInfo -Message 'Installing ODBC Driver 18.5.x'
    $odbcMsi = Join-Path $labSources 'SoftwarePackages/ODBC/msodbcsql.msi'
    if (-not (Test-Path $odbcMsi)) {
        Write-ScreenInfo -Message 'Downloading ODBC Driver 18.5.2.1...'
        $odbcMsi = (Get-LabInternetFile -Uri 'https://go.microsoft.com/fwlink/?linkid=2335671' -Path "$labSources/SoftwarePackages/ODBC" -FileName 'msodbcsql.msi' -PassThru -ErrorAction Stop).FullName
    }
    Install-LabSoftwarePackage -Path $odbcMsi -ComputerName $vms -CommandLine '/qn /norestart IACCEPTMSODBCSQLLICENSETERMS=YES' -ExpectedReturnCodes 0, 3010
    #endregion

    $WMIv2Zip = "{0}\WmiExplorer.zip" -f (Get-LabSourcesLocation -Local)
    $WMIv2Exe = "{0}\WmiExplorer.exe" -f (Get-LabSourcesLocation -Local)
    $wmiExpUrl = Get-LabConfigurationItem -Name ConfigurationManagerWmiExplorer

    try
    {
        Get-LabInternetFile -Uri $wmiExpUrl -Path (Split-Path -Path $WMIv2Zip -Parent) -FileName (Split-Path -Path $WMIv2Zip -Leaf) -ErrorAction "Stop" -ErrorVariable "GetLabInternetFileErr"
    }
    catch
    {
        Write-ScreenInfo -Message ("Could not download from '{0}' ({1})" -f $wmiExpUrl, $GetLabInternetFileErr.ErrorRecord.Exception.Message) -Type "Warning"
    }

    Expand-Archive -Path $WMIv2Zip -DestinationPath "$(Get-LabSourcesLocation -Local)/Tools" -ErrorAction "Stop" -Force
    try
    {
        Remove-Item -Path $WMIv2Zip -Force -ErrorAction "Stop" -ErrorVariable "RemoveItemErr"
    }
    catch
    {
        Write-ScreenInfo -Message ("Failed to delete '{0}' ({1})" -f $WMIZip, $RemoveItemErr.ErrorRecord.Exception.Message) -Type "Warning"
    }

    if ((Get-Lab).DefaultVirtualizationEngine -eq 'Azure') { Sync-LabAzureLabSources -Filter WmiExplorer.exe }

    # ConfigurationManager
    foreach ($vm in $vms)
    {
        $role = $vm.Roles.Where( { $_.Name -eq 'ConfigurationManager' })
        $cmVersion = if ($role.Properties.ContainsKey('Version')) { $role.Properties.Version } else { '2103' }
        $cmBranch = if ($role.Properties.ContainsKey('Branch')) { $role.Properties.Branch } else { 'CB' }

        $VMInstallDirectory = "$deployDebugPath\Install"
        $CMBinariesDirectory = "$labSources\SoftwarePackages\CM-$($cmVersion)-$cmBranch"
        $CMPreReqsDirectory = "$labSources\SoftwarePackages\CM-Prereqs-$($cmVersion)-$cmBranch"
        $VMCMBinariesDirectory = "{0}\CM" -f $VMInstallDirectory
        $VMCMPreReqsDirectory = "{0}\CM-PreReqs" -f $VMInstallDirectory

        # Check for local CM source first (avoids hardcoded version URLs)
        $cmLocalSource = Get-LabConfigurationItem -Name 'ConfigurationManagerLocalSource'
        $cmDownloadUrl = Get-LabConfigurationItem -Name "ConfigurationManagerUrl$($cmVersion)$($cmBranch)"

        if ($cmLocalSource -and (Test-Path $cmLocalSource))
        {
            Write-ScreenInfo -Message "Using local CM source: $cmLocalSource"
            $CMBinariesDirectory = $cmLocalSource
        }
        elseif ($cmDownloadUrl)
        {
            #region CM binaries download
            $CMZipPath = "{0}\SoftwarePackages\{1}" -f $labsources, ((Split-Path $CMDownloadURL -Leaf) -replace "\.exe$", ".zip")

            try
            {
                $CMZipObj = Get-LabInternetFile -Uri $CMDownloadURL -Path (Split-Path -Path $CMZipPath -Parent) -FileName (Split-Path -Path $CMZipPath -Leaf) -PassThru -ErrorAction "Stop" -ErrorVariable "GetLabInternetFileErr"
            }
            catch
            {
                $Message = "Failed to download from '{0}' ({1})" -f $CMDownloadURL, $GetLabInternetFileErr.ErrorRecord.Exception.Message
                Write-LogFunctionExitWithError -Message $Message
            }
            #endregion
        }
        else
        {
            # Check if CM source exists in standard SoftwarePackages location
            $autoDetect = Get-ChildItem "$labSources\SoftwarePackages\CM" -Directory -ErrorAction SilentlyContinue |
                Where-Object { Test-Path (Join-Path $_.FullName 'SMSSETUP\BIN\X64\setup.exe') } |
                Select-Object -First 1
            if ($autoDetect)
            {
                Write-ScreenInfo -Message "Auto-detected CM source: $($autoDetect.FullName)"
                $CMBinariesDirectory = $autoDetect.FullName
            }
            else
            {
                Write-LogFunctionExitWithError -Message "No URI configuration for CM version $cmVersion, branch $cmBranch, and no local source found in $labSources\SoftwarePackages\CM\"
            }
        }

        #region Extract CM binaries (only if downloaded, not local)
        if ($CMZipObj)
        {
        try
        {
            if ((Get-Lab).DefaultVirtualizationEngine -eq 'Azure')
            {
                Invoke-LabCommand -Computer $vm -ScriptBlock {
                    $null = mkdir -Force $VMCMBinariesDirectory
                    Expand-Archive -Path $CMZipObj.FullName -DestinationPath $VMCMBinariesDirectory -Force
                } -Variable (Get-Variable VMCMBinariesDirectory, CMZipObj)
            }
            else
            {
                Expand-Archive -Path $CMZipObj.FullName -DestinationPath $CMBinariesDirectory -Force -ErrorAction "Stop" -ErrorVariable "ExpandArchiveErr"
                Copy-LabFileItem -Path $CMBinariesDirectory/* -Destination $VMCMBinariesDirectory -ComputerName $vm -Recurse
            }
        
        }
        catch
        {
            $Message = "Failed to initiate extraction to '{0}' ({1})" -f $CMBinariesDirectory, $ExpandArchiveErr.ErrorRecord.Exception.Message
            Write-LogFunctionExitWithError -Message $Message
        }
        }
        else
        {
            # Local source — copy directly to VM
            Copy-LabFileItem -Path "$CMBinariesDirectory\*" -Destination $VMCMBinariesDirectory -ComputerName $vm -Recurse
        }
        #endregion

        #region Download CM prerequisites
        switch ($cmBranch)
        {
            "CB"
            {
                if ((Get-Lab).DefaultVirtualizationEngine -eq 'Azure')
                {
                    Install-LabSoftwarePackage -ComputerName $vm -LocalPath $VMCMBinariesDirectory\SMSSETUP\BIN\X64\setupdl.exe -CommandLine "/NOUI $VMCMPreReqsDirectory" -UseShellExecute -AsScheduledJob
                    break       
                }
                
                try
                {
                    $p = Start-Process -FilePath $CMBinariesDirectory\SMSSETUP\BIN\X64\setupdl.exe -ArgumentList "/NOUI", $CMPreReqsDirectory -PassThru -ErrorAction "Stop" -ErrorVariable "StartProcessErr" -Wait
                    Copy-LabFileItem -Path $CMPreReqsDirectory/* -Destination $VMCMPreReqsDirectory -Recurse -ComputerName $vm
                }
                catch
                {
                    $Message = "Failed to initiate download of CM pre-req files to '{0}' ({1})" -f $CMPreReqsDirectory, $StartProcessErr.ErrorRecord.Exception.Message
                    Write-LogFunctionExitWithError -Message $Message
                }
            }
            "TP"
            {
                $Messages = @(
                    "Directory '{0}' is intentionally empty." -f $CMPreReqsDirectory
                    "The prerequisites will be downloaded by the installer within the VM."
                    "This is a workaround due to a known issue with TP 2002 baseline: https://twitter.com/codaamok/status/1268588138437509120"
                )

                try
                {
                    $CMPreReqsDirectory = "$(Get-LabSourcesLocation -Local)\SoftwarePackages\CM-Prereqs-$($cmVersion)-$cmBranch"
                    $PreReqDirObj = New-Item -Path $CMPreReqsDirectory -ItemType "Directory" -Force -ErrorAction "Stop" -ErrorVariable "CreateCMPreReqDir"
                    Set-Content -Path ("{0}\readme.txt" -f $PreReqDirObj.FullName) -Value $Messages -ErrorAction "SilentlyContinue"
                }
                catch
                {
                    $Message = "Failed to create CM prerequisite directory '{0}' ({1})" -f $CMPreReqsDirectory, $CreateCMPreReqDir.ErrorRecord.Exception.Message
                    Write-LogFunctionExitWithError -Message $Message
                }
            }
        }

        $siteParameter = @{
            CMServerName        = $vm
            CMBinariesDirectory = $CMBinariesDirectory
            Branch              = $cmBranch
            CMPreReqsDirectory  = $CMPreReqsDirectory
            CMSiteCode          = 'AL1'
            CMSiteName          = 'AutomatedLab-01'
            CMRoles             = 'Management Point', 'Distribution Point'
            DatabaseName        = 'ALCMDB'
        }

        if ($role.Properties.ContainsKey('SiteCode'))
        {
            $siteParameter.CMSiteCode = $role.Properties.SiteCode
        }

        if ($role.Properties.ContainsKey('SiteName'))
        {
            $siteParameter.CMSiteName = $role.Properties.SiteName
        }

        if ($role.Properties.ContainsKey('ProductId'))
        {
            $siteParameter.CMProductId = $role.Properties.ProductId
        }

        $validRoles = @(
            "None",
            "Management Point", 
            "Distribution Point", 
            "Software Update Point", 
            "Reporting Services Point", 
            "Endpoint Protection Point"
        )
        if ($role.Properties.ContainsKey('Roles'))
        {
            $siteParameter.CMRoles = if ($role.Properties.Roles.Split(',') -contains 'None')
            {
                'None'
            }
            else
            {
                $role.Properties.Roles.Split(',') | Where-Object { $_ -in $validRoles } | Sort-Object -Unique
            }
        }

        if ($role.Properties.ContainsKey('SqlServerName'))
        {
            $sql = $role.Properties.SqlServerName

            if (-not (Get-LabVm -ComputerName $sql.Split('.')[0]))
            {
                Write-ScreenInfo -Type Warning -Message "No SQL server called $sql found in lab. If you wanted to use an existing instance, don't forget to add it with the -SkipDeployment parameter"
            }

            $siteParameter.SqlServerName = $sql
        }
        else
        {
            $sql = (Get-LabVM -Role SQLServer2014, SQLServer2016, SQLServer2017, SQLServer2019 | Select-Object -First 1).Fqdn

            if (-not $sql)
            {
                Write-LogFunctionExitWithError -Message "No SQL server found in lab. Cannot install SCCM"
            }

            $siteParameter.SqlServerName = $sql
        }

        Invoke-LabCommand -ComputerName $sql.Split('.')[0] -ActivityName 'Add computer account as local admin (why...)' -ScriptBlock {
            Add-LocalGroupMember -Group Administrators -Member "$($vm.DomainName)\$($vm.Name)`$"
        } -Variable (Get-Variable vm)

        if ($role.Properties.ContainsKey('DatabaseName'))
        {
            $siteParameter.DatabaseName = $role.Properties.DatabaseName
        }

        if ($role.Properties.ContainsKey('AdminUser'))
        {
            $siteParameter.AdminUser = $role.Properties.AdminUser
        }

        if ($role.Properties.ContainsKey('WsusContentPath'))
        {
            $siteParameter.WsusContentPath = $role.Properties.WsusContentPath
        }
        Install-CMSite @siteParameter

        Restart-LabVM -ComputerName $vm

        if (Test-LabMachineInternetConnectivity -ComputerName $vm)
        {
            Write-ScreenInfo -Type Verbose -Message "$vm is connected, beginning update process"
            $updateParameter = Sync-Parameter -Command (Get-Command Update-CMSite) -Parameters $siteParameter
            Update-CMSite @updateParameter
        }
    }
    #endregion
}
