function Checkpoint-LWAzureVM
{
    [Cmdletbinding()]
    Param
    (
        [Parameter(Mandatory)]
        [string[]]$ComputerName,

        [Parameter(Mandatory)]
        [string]$SnapshotName
    )

    Test-LabHostConnected -Throw -Quiet

    Write-LogFunctionEntry

    $lab = Get-Lab
    $resourceGroupName = $lab.AzureSettings.DefaultResourceGroup.ResourceGroupName
    $runningMachines = Get-LabVM -IsRunning -ComputerName $ComputerName -IncludeLinux
    if ($runningMachines)
    {
        Stop-LWAzureVM -ComputerName $runningMachines -StayProvisioned $true
        Wait-LabVMShutdown -ComputerName $runningMachines
    }

    $jobs = foreach ($machine in $ComputerName)
    {
        $vm = Get-AzVM -ResourceGroupName $resourceGroupName -Name $machine -ErrorAction SilentlyContinue
        if (-not $vm)
        {
            Write-ScreenInfo -Message "$machine could not be found in $($resourceGroupName). Skipping snapshot." -type Warning
            continue
        }

        $vmSnapshotName = '{0}_{1}' -f $machine, $SnapshotName
        $existingSnapshot = Get-AzSnapshot -ResourceGroupName $resourceGroupName -SnapshotName $vmSnapshotName -ErrorAction SilentlyContinue
        if ($existingSnapshot)
        {
            Write-ScreenInfo -Message "Snapshot $SnapshotName for $machine already exists as $($existingSnapshot.Name). Not creating it again." -Type Warning
            continue
        }

        $osSourceDisk = Get-AzDisk -ResourceGroupName $resourceGroupName -DiskName $vm.StorageProfile.OsDisk.Name
        $snapshotConfig = New-AzSnapshotConfig -SourceUri $osSourceDisk.Id -CreateOption Copy -Location $vm.Location
        New-AzSnapshot -Snapshot $snapshotConfig -SnapshotName $vmSnapshotName -ResourceGroupName $resourceGroupName -AsJob
    }

    if ($jobs.State -contains 'Failed')
    {
        Write-ScreenInfo -Type Error -Message "At least one snapshot creation failed: $($jobs.Name -join ',')."
        $skipRemove = $true
    }

    if ($jobs)
    {
        $null = $jobs | Wait-Job
        $jobs | Remove-Job
    }

    if ($runningMachines)
    {
        Start-LWAzureVM -ComputerName $runningMachines
        Wait-LabVM -ComputerName $runningMachines
    }

    Write-LogFunctionExit
}


function Connect-LWAzureLabSourcesDrive
{
    param(
        [Parameter(Mandatory, Position = 0)]
        [System.Management.Automation.Runspaces.PSSession]$Session,

        [switch]$SuppressErrors
    )

    Test-LabHostConnected -Throw -Quiet

    Write-LogFunctionEntry

    $azureRetryCount = Get-LabConfigurationItem -Name AzureRetryCount
    $labSourcesStorageAccount = Get-LabAzureLabSourcesStorage -ErrorAction SilentlyContinue

    if (Get-LabConfigurationItem -Name AzureDisableLabSourcesStorage) {
        Write-ScreenInfo -Type Verbose -Message "User opted out of storage account creation."
        return
    }

    if ($Session.Runspace.ConnectionInfo.AuthenticationMechanism -notin 'CredSsp', 'Negotiate' -or -not $labSourcesStorageAccount)
    {
        return
    }

    $result = Invoke-Command -Session $Session -ScriptBlock {
        #Add *.windows.net to Local Intranet Zone
        $path = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\windows.net'
        if (-not (Test-Path -Path $path)) {
            New-Item -Path $path -Force

            New-ItemProperty $path -Name http -Value 1 -Type DWORD
            New-ItemProperty $path -Name file -Value 1 -Type DWORD
        }

        $hostName = ([uri]$args[0]).Host
	    $dnsRecord = Resolve-DnsName -Name $hostname | Where-Object { $_ -is [Microsoft.DnsClient.Commands.DnsRecord_A] }
        $ipAddress = $dnsRecord.IPAddress
        $rangeName = $ipAddress.Replace('.', '')

        $path = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Ranges\$rangeName"
        if (-not (Test-Path -Path $path)) {
            New-Item -Path $path -Force

            New-ItemProperty $path -Name :Range -Value $ipAddress -Type String
            New-ItemProperty $path -Name http -Value 1 -Type DWORD
            New-ItemProperty $path -Name file -Value 1 -Type DWORD
        }

        $pattern = '^(OK|Unavailable) +(?<DriveLetter>\w): +\\\\automatedlab'

        #remove all drive connected to an Azure LabSources share that are no longer available
        $drives = net.exe use
        $netRemoveResult = @()
        foreach ($line in $drives)
        {
            if ($line -match $pattern)
            {
                $netRemoveResult += net.exe use "$($Matches.DriveLetter):" /d
            }
        }

        $cmd = 'net.exe use * {0} /u:{1} {2}' -f $args[0], $args[1], $args[2]
        $cmd = [scriptblock]::Create($cmd)
        $netConnectResult = &$cmd 2>&1

        if (-not $LASTEXITCODE)
        {
            $ALLabSourcesMapped = $true
            $alDriveLetter = (Get-PSDrive | Where-Object DisplayRoot -like \\automatedlabsources*).Name
            Get-ChildItem -Path "$($alDriveLetter):" | Out-Null #required, otherwise sometimes accessing the UNC path did not work
        }

        New-Object PSObject -Property @{
            ReturnCode         = $LASTEXITCODE
            ALLabSourcesMapped = [bool](-not $LASTEXITCODE)
            NetConnectResult   = $netConnectResult
            NetRemoveResult    = $netRemoveResult
        }

    } -ArgumentList $labSourcesStorageAccount.Path, $labSourcesStorageAccount.StorageAccountName, $labSourcesStorageAccount.StorageAccountKey

    $Session | Add-Member -Name ALLabSourcesMappingResult -Value $result -MemberType NoteProperty -Force
    $Session | Add-Member -Name ALLabSourcesMapped -Value $result.ALLabSourcesMapped -MemberType NoteProperty -Force

    if ($result.ReturnCode -ne 0 -and -not $SuppressErrors)
    {
        $netResult = $result | Where-Object { $_.ReturnCode -gt 0 }
        Write-LogFunctionExitWithError -Message "Connecting session '$($s.Name)' to LabSources folder failed" -Details $netResult.NetConnectResult
    }

    Write-LogFunctionExit
}


function Disable-LWAzureAutoShutdown
{
    param
    (
        [string[]]
        $ComputerName,

        [switch]
        $Wait
    )

    $lab = Get-Lab -ErrorAction Stop
    $labVms = Get-AzVm -ResourceGroupName $lab.AzureSettings.DefaultResourceGroup.ResourceGroupName
    if ($ComputerName)
    {
        $labVms = $labVms | Where-Object Name -in $ComputerName
    }
    $resourceIdString = '{0}/providers/microsoft.devtestlab/schedules/shutdown-computevm-' -f $lab.AzureSettings.DefaultResourceGroup.ResourceId

    $jobs = foreach ($vm in $labVms)
    {
        Remove-AzResource -ResourceId ("$($resourceIdString)$($vm.Name)") -Force -ErrorAction SilentlyContinue -AsJob
    }

    if ($jobs -and $Wait.IsPresent)
    {
        $null = $jobs | Wait-Job
    }
}


function Dismount-LWAzureIsoImage
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseCompatibleCmdlets", "", Justification = "Not relevant, used in Invoke-LabCommand")]
    param
    (
        [Parameter(Mandatory, Position = 0)]
        [string[]]
        $ComputerName
    )

    Test-LabHostConnected -Throw -Quiet

    $azureRetryCount = Get-LabConfigurationItem -Name AzureRetryCount

    Invoke-LabCommand -ComputerName $ComputerName -ActivityName "Dismounting ISO Images on Azure machines $($ComputerName -join ',')" -ScriptBlock {

        Get-Volume | 
        Where-Object DriveType -eq CD-ROM |
        ForEach-Object {
            Get-DiskImage -DevicePath $_.Path.TrimEnd('\') -ErrorAction SilentlyContinue
        } |
        ForEach-Object {
            Write-Verbose -Message "Dismounting '$($_.ImagePath)'"
            $_ | Dismount-DiskImage
        }

        Get-ChildItem -Path C:\ALMounts\*.iso -ErrorAction SilentlyContinue | Remove-Item
    } -NoDisplay
}


function Enable-LWAzureAutoShutdown
{
    param
    (
        [string[]]
        $ComputerName,

        [timespan]
        $Time,

        [string]
        $TimeZone = (Get-TimeZone).Id,

        [switch]
        $Wait
    )

    $lab = Get-Lab -ErrorAction Stop
    $labVms = Get-AzVm -ResourceGroupName $lab.AzureSettings.DefaultResourceGroup.ResourceGroupName
    if ($ComputerName)
    {
        $labVms = $labVms | Where-Object Name -in $ComputerName
    }
    $resourceIdString = '{0}/providers/microsoft.devtestlab/schedules/shutdown-computevm-' -f $lab.AzureSettings.DefaultResourceGroup.ResourceId

    $jobs = foreach ($vm in $labVms)
    {
        $properties = @{
            status           = 'Enabled'
            taskType         = 'ComputeVmShutdownTask'
            dailyRecurrence  = @{time = $Time.ToString('hhmm') }
            timeZoneId       = $TimeZone
            targetResourceId = $vm.Id
        }

        New-AzResource -ResourceId ("$($resourceIdString)$($vm.Name)") -Location $vm.Location -Properties $properties -Force -ErrorAction SilentlyContinue -AsJob
    }

    if ($jobs -and $Wait.IsPresent)
    {
        $null = $jobs | Wait-Job
    }
}


function Enable-LWAzureVMRemoting
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseCompatibleCmdlets", "", Justification = "Not enabling CredSSP a third time on Linux")]
    param
    (
        [Parameter(Mandatory, Position = 0)]
        [string[]]$ComputerName,

        [switch]$UseSSL
    )

    Test-LabHostConnected -Throw -Quiet

    $azureRetryCount = Get-LabConfigurationItem -Name AzureRetryCount

    if ($ComputerName)
    {
        $machines = Get-LabVM -All -IncludeLinux | Where-Object Name -in $ComputerName
    }
    else
    {
        $machines = Get-LabVM -All -IncludeLinux
    }

    $script = {
        param ($DomainName, $UserName, $Password)

        $RegPath = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon'

        Set-ItemProperty -Path $RegPath -Name AutoAdminLogon -Value 1 -ErrorAction SilentlyContinue
        Set-ItemProperty -Path $RegPath -Name DefaultUserName -Value $UserName -ErrorAction SilentlyContinue
        Set-ItemProperty -Path $RegPath -Name DefaultPassword -Value $Password -ErrorAction SilentlyContinue
        Set-ItemProperty -Path $RegPath -Name DefaultDomainName -Value $DomainName -ErrorAction SilentlyContinue

        #Enable-WSManCredSSP works fine when called remotely on 2012 servers but not on 2008 (Access Denied). In case Enable-WSManCredSSP fails
        #the settings are done in the registry directly
        try
        {
            Enable-WSManCredSSP -Role Server -Force | Out-Null
        }
        catch
        {
            New-ItemProperty -Path HKLM:\software\Microsoft\Windows\CurrentVersion\WSMAN\Service -Name auth_credssp -Value 1 -PropertyType DWORD -Force
            New-ItemProperty -Path HKLM:\software\Microsoft\Windows\CurrentVersion\WSMAN\Service -Name allow_remote_requests -Value 1 -PropertyType DWORD -Force
        }
    }

    foreach ($machine in $machines)
    {
        $cred = $machine.GetCredential((Get-Lab))
        try
        {
            Invoke-LabCommand -ComputerName $machine -ActivityName SetLabVMRemoting -ScriptBlock $script -DoNotUseCredSsp -NoDisplay `
                -ArgumentList $machine.DomainName, $cred.UserName, $cred.GetNetworkCredential().Password -ErrorAction Stop -UseLocalCredential
        }
        catch
        {
            if ($IsLinux)
            {
                return
            }

            if ($UseSSL)
            {
                Connect-WSMan -ComputerName $machine.AzureConnectionInfo.DnsName -Credential $cred -Port $machine.AzureConnectionInfo.Port -UseSSL -SessionOption (New-WSManSessionOption -SkipCACheck -SkipCNCheck)
            }
            else
            {
                Connect-WSMan -ComputerName $machine.AzureConnectionInfo.DnsName -Credential $cred -Port $machine.AzureConnectionInfo.Port
            }

            Set-Item -Path "WSMan:\$($machine.AzureConnectionInfo.DnsName)\Service\Auth\CredSSP" -Value $true
            Disconnect-WSMan -ComputerName $machine.AzureConnectionInfo.DnsName
        }
    }
}


function Enable-LWAzureWinRm
{
    param
    (
        [Parameter(Mandatory)]
        [AutomatedLab.Machine[]]
        $Machine,

        [switch]
        $PassThru,

        [switch]
        $Wait
    )

    Test-LabHostConnected -Throw -Quiet

    Write-LogFunctionEntry

    $azureRetryCount = Get-LabConfigurationItem -Name AzureRetryCount

    $lab = Get-Lab
    $jobs = @()

    $tempFileName = Join-Path -Path ([IO.Path]::GetTempPath()) -ChildPath enableazurewinrm.labtempfile.ps1
    $customScriptContent = @"
`$deployDebug = (New-Item -ItemType Directory -Path `$ExecutionContext.InvokeCommand.ExpandString("$AL_DeployDebugFolder") -Force).FullName
`$null = New-Item -ItemType Directory -Path `$deployDebug\ALAzure -ErrorAction SilentlyContinue
'Trying to enable Remoting and CredSSP' | Out-File `$deployDebug\ALAzure\WinRmActivation.log -Append
try
{
Enable-PSRemoting -Force -ErrorAction Stop
"Successfully called Enable-PSRemoting" | Out-File `$deployDebug\ALAzure\WinRmActivation.log -Append
}
catch
{
"Error calling Enable-PSRemoting. `$(`$_.Exception.Message)" | Out-File `$deployDebug\ALAzure\WinRmActivation.log -Append
}
try
{
Enable-WSManCredSSP -Role Server -Force | Out-Null
"Successfully enabled CredSSP" | Out-File `$deployDebug\ALAzure\WinRmActivation.log -Append
}
catch
{
try
{
New-ItemProperty -Path HKLM:\software\Microsoft\Windows\CurrentVersion\WSMAN\Service -Name auth_credssp -Value 1 -PropertyType DWORD -Force -ErrorACtion Stop
New-ItemProperty -Path HKLM:\software\Microsoft\Windows\CurrentVersion\WSMAN\Service -Name allow_remote_requests -Value 1 -PropertyType DWORD -Force -ErrorAction Stop
"Enabled CredSSP via Registry" | Out-File `$deployDebug\ALAzure\WinRmActivation.log -Append
}
catch
{
"Could not enable CredSSP via cmdlet or registry!" | Out-File `$deployDebug\ALAzure\WinRmActivation.log -Append
}
}
"@
    $customScriptContent | Out-File $tempFileName -Force -Encoding utf8
    $rgName = Get-LabAzureDefaultResourceGroup

    $jobs = foreach ($m in $Machine)
    {
        if ($Lab.AzureSettings.IsAzureStack)
        {
            $sa = Get-AzStorageAccount -ResourceGroupName $lab.AzureSettings.DefaultResourceGroup.ResourceGroupName -ErrorAction SilentlyContinue
            if (-not $sa)
            {
                $sa = New-AzStorageAccount -Name "cse$(-join (1..10 | % {[char](Get-Random -Min 97 -Max 122)}))" -ResourceGroupName $lab.AzureSettings.DefaultResourceGroup.ResourceGroupName -SkuName Standard_LRS -Kind Storage -Location (Get-LabAzureDefaultLocation).Location
            }

            $co = $sa | Get-AzStorageContainer -Name customscriptextension -ErrorAction SilentlyContinue
            if (-not $co)
            {
                $co = $sa | New-AzStorageContainer -Name customscriptextension
            }

            $content = Set-AzStorageBlobContent -File $tempFileName -CloudBlobContainer $co.CloudBlobContainer -Blob $(Split-Path -Path $tempFileName -Leaf) -Context $sa.Context -Force -ErrorAction Stop
            $token = New-AzStorageBlobSASToken -CloudBlob $content.ICloudBlob -StartTime (Get-Date) -ExpiryTime $(Get-Date).AddHours(1) -Protocol HttpsOnly -Context $sa.Context -Permission r -ErrorAction Stop
            $uri = '{0}{1}/{2}{3}' -f $co.Context.BlobEndpoint, 'customscriptextension', $(Split-Path -Path $tempFileName -Leaf), $token
            [version] $typehandler = (Get-AzVMExtensionImage -PublisherName Microsoft.Compute -Type CustomScriptExtension -Location (Get-LabAzureDefaultLocation).Location | Sort-Object { [version]$_.Version } | Select-Object -Last 1).Version
            
            $extArg = @{
                ResourceGroupName  = $lab.AzureSettings.DefaultResourceGroup.ResourceGroupName
                VMName             = $m.ResourceName
                FileUri            = $uri
                TypeHandlerVersion = '{0}.{1}' -f $typehandler.Major, $typehandler.Minor
                Name               = 'initcustomizations'
                Location           = (Get-LabAzureDefaultLocation).Location
                Run                = Split-Path -Path $tempFileName -Leaf
                NoWait             = $true
            }
            $Null = Set-AzVMCustomScriptExtension @extArg
        }
        else
        {
            Invoke-AzVMRunCommand -ResourceGroupName $rgName -VMName $m.ResourceName -ScriptPath $tempFileName -CommandId 'RunPowerShellScript' -ErrorAction Stop -AsJob
        }
    }

    if ($Wait)
    {
        Wait-LWLabJob -Job $jobs

        $results = $jobs | Receive-Job -Keep -ErrorAction SilentlyContinue -ErrorVariable +AL_AzureWinrmActivationErrors
        $failedJobs = $jobs | Where-Object -Property Status -eq 'Failed'

        if ($failedJobs)
        {
            $machineNames = $($($failedJobs).Name -replace "'").ForEach( { $($_ -split '\s')[-1] })
            Write-ScreenInfo -Type Error -Message ('Enabling CredSSP on the following lab machines failed: {0}. Check the output of "Get-Job -Id {1} | Receive-Job -Keep" as well as the variable $AL_AzureWinrmActivationErrors' -f $($machineNames -join ','), $($failedJobs.Id -join ','))
        }
    }

    if ($PassThru)
    {
        $jobs
    }

    Remove-Item $tempFileName -Force -ErrorAction SilentlyContinue
    Write-LogFunctionExit
}


function Get-LWAzureAutoShutdown
{
    [CmdletBinding()]
    param ( )

    $lab = Get-Lab -ErrorAction Stop
    $resourceGroup = $lab.AzureSettings.DefaultResourceGroup.ResourceGroupName

    $schedules = (Get-AzResource -ResourceGroupName $resourceGroup -ResourceType Microsoft.DevTestLab/schedules -ExpandProperties -ErrorAction SilentlyContinue).Properties

    foreach ($schedule in $schedules)
    {
        $hour, $minute = Get-StringSection -SectionSize 2 -String $schedule.dailyRecurrence.time

        if ($schedule)
        {
            [PSCustomObject]@{
                ComputerName = ($schedule.targetResourceId -split '/')[-1]
                Time         = New-TimeSpan -Hours $hour -Minutes $minute
                TimeZone     = Get-TimeZone -Id $schedule.timeZoneId
            }
        }
    }
}


function Get-LWAzureSku
{
    [Cmdletbinding()]
    param
    (
        [Parameter(Mandatory)]
        [AutomatedLab.Machine]$Machine
    )

    $lab = Get-Lab

    #if this machine has a SQL Server role
    foreach ($role in $Machine.Roles)
    {
        if ($role.Name -match 'SQLServer(?<SqlVersion>\d{4})')
        {
            #get the SQL Server version defined in the role
            $sqlServerRoleName = $Matches[0]
            $sqlServerVersion = $Matches.SqlVersion

            if ($role.Properties.Keys | Where-Object { $_ -ne 'InstallSampleDatabase' })
            {
                $useStandardVm = $true
            }
        }

        if ($role.Name -match 'VisualStudio(?<Version>\d{4})')
        {
            $visualStudioRoleName = $Matches[0]
            $visualStudioVersion = $Matches.Version
        }
    }

    if ($sqlServerRoleName -and -not $useStandardVm)
    {
        Write-PSFMessage -Message 'This is going to be a SQL Server VM'
        $pattern = 'SQL(?<SqlVersion>\d{4})(?<SqlIsR2>R2)??(?<SqlServicePack>SP\d)?-(?<OS>WS\d{4}(R2)?)'

        #get all SQL images matching the RegEx pattern and then get only the latest one
        $sqlServerImages = $lab.AzureSettings.VmImages | Where-Object Offer -notlike "*BYOL*"

        if ([System.Convert]::ToBoolean($Machine.AzureProperties['UseByolImage']))
        {
            $sqlServerImages = $lab.AzureSettings.VmImages | Where-Object Offer -like '*-BYOL'
        }

        $sqlServerImages = $sqlServerImages |
        Where-Object Offer -Match $pattern |
        Group-Object -Property Sku, Offer |
        ForEach-Object {
            $_.Group | Sort-Object -Property PublishedDate -Descending | Select-Object -First 1
        }

        #add the version, SP Level and OS from the ImageFamily field to the image object
        foreach ($sqlServerImage in $sqlServerImages)
        {
            $sqlServerImage.Offer -match $pattern | Out-Null

            $sqlServerImage | Add-Member -Name SqlVersion -Value $Matches.SqlVersion -MemberType NoteProperty -Force
            $sqlServerImage | Add-Member -Name SqlIsR2 -Value $Matches.SqlIsR2 -MemberType NoteProperty -Force
            $sqlServerImage | Add-Member -Name SqlServicePack -Value $Matches.SqlServicePack -MemberType NoteProperty -Force

            $sqlServerImage | Add-Member -Name OS -Value (New-Object AutomatedLab.OperatingSystem($Matches.OS)) -MemberType NoteProperty -Force
        }

        #get the image that matches the OS and SQL server version
        $machineOs = New-Object AutomatedLab.OperatingSystem($machine.OperatingSystem)
        $vmImage = $sqlServerImages | Where-Object { $_.SqlVersion -eq $sqlServerVersion -and $_.OS.Version -eq $machineOs.Version } |
        Sort-Object -Property SqlServicePack -Descending | Select-Object -First 1
        $offerName = $vmImageName = $vmImage.Offer
        $publisherName = $vmImage.PublisherName
        $skusName = $vmImage.Skus

        if (-not $vmImageName)
        {
            Write-ScreenInfo 'SQL Server image could not be found. The following combinations are currently supported by Azure:' -Type Warning
            foreach ($sqlServerImage in $sqlServerImages)
            {
                Write-PSFMessage -Level Host $sqlServerImage.Offer
            }

            throw "There is no Azure VM image for '$sqlServerRoleName' on operating system '$($machine.OperatingSystem)'. The machine cannot be created. Cancelling lab setup. Please find the available images above."
        }
    }
    elseif ($visualStudioRoleName)
    {
        Write-PSFMessage -Message 'This is going to be a Visual Studio VM'

        $pattern = 'VS-(?<Version>\d{4})-(?<Edition>\w+)-VSU(?<Update>\d)-AzureSDK-\d{2,3}-((?<OS>WIN\d{2})|(?<OS>WS\d{4,6}))'

        #get all SQL images machting the RegEx pattern and then get only the latest one
        $visualStudioImages = $lab.AzureSettings.VmImages |
        Where-Object Offer -EQ VisualStudio

        #add the version, SP Level and OS from the ImageFamily field to the image object
        foreach ($visualStudioImage in $visualStudioImages)
        {
            $visualStudioImage.Skus -match $pattern | Out-Null

            $visualStudioImage | Add-Member -Name Version -Value $Matches.Version -MemberType NoteProperty -Force
            $visualStudioImage | Add-Member -Name Update -Value $Matches.Update -MemberType NoteProperty -Force

            $visualStudioImage | Add-Member -Name OS -Value (New-Object AutomatedLab.OperatingSystem($Matches.OS)) -MemberType NoteProperty -Force
        }

        #get the image that matches the OS and SQL server version
        $machineOs = New-Object AutomatedLab.OperatingSystem($machine.OperatingSystem)
        $vmImage = $visualStudioImages | Where-Object { $_.Version -eq $visualStudioVersion -and $_.OS.Version.Major -eq $machineOs.Version.Major } |
        Sort-Object -Property Update -Descending | Select-Object -First 1
        $offerName = $vmImageName = ($vmImage).Offer
        $publisherName = ($vmImage).PublisherName
        $skusName = ($vmImage).Skus

        if (-not $vmImageName)
        {
            Write-ScreenInfo 'Visual Studio image could not be found. The following combinations are currently supported by Azure:' -Type Warning
            foreach ($visualStudioImage in $visualStudioImages)
            {
                Write-ScreenInfo ('{0} - {1} - {2}' -f $visualStudioImage.Offer, $visualStudioImage.Skus, $visualStudioImage.Id)
            }

            throw "There is no Azure VM image for '$visualStudioRoleName' on operating system '$($machine.OperatingSystem)'. The machine cannot be created. Cancelling lab setup. Please find the available images above."
        }
    }
    else
    {
        $vmImage = $lab.AzureSettings.VmImages |
        Where-Object { $_.AutomatedLabOperatingSystemName -eq $Machine.OperatingSystem.OperatingSystemName -and $_.HyperVGeneration -eq "V$($Machine.VmGeneration)" } |
        Select-Object -First 1

        if (-not $vmImage)
        {
            throw "There is no Azure VM image for the operating system '$($Machine.OperatingSystem)'. The machine cannot be created. Cancelling lab setup."
        }

        $offerName = ($vmImage).Offer
        $publisherName = ($vmImage).PublisherName
        $skusName = ($vmImage).Skus
        $version = $vmImage.Version
    }

    Write-PSFMessage -Message "We selected the SKUs $skusName from offer $offerName by publisher $publisherName"
    @{
        offer     = $offerName
        publisher = $publisherName
        sku       = $skusName
        version   = if ($version) { $version } else { 'latest' }
    }
}


function Get-LWAzureVm
{
    [CmdletBinding()]
    param (
        [Parameter()]
        [string[]]$ComputerName
    )

    Test-LabHostConnected -Throw -Quiet

    #required to suporess verbose messages, warnings and errors
    Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState

    Write-LogFunctionEntry

    $azureRetryCount = Get-LabConfigurationItem -Name AzureRetryCount

    $azureVms = Get-AzVM -Status -ResourceGroupName (Get-LabAzureDefaultResourceGroup).ResourceGroupName -ErrorAction SilentlyContinue -ErrorVariable getazvmerror
    $count = 1
    while (-not $azureVms -and $count -le $azureRetryCount)
    {
        Write-ScreenInfo -Type Verbose -Message "Get-AzVM did not return anything, attempt $count of $($azureRetryCount) attempts. Azure presented us with the error: $($getazvmerror.Exception.Message)"
        Start-Sleep -Seconds 2
        $azureVms = Get-AzVM -Status -ResourceGroupName (Get-LabAzureDefaultResourceGroup).ResourceGroupName -ErrorAction SilentlyContinue -ErrorVariable getazvmerror
        $count++
    }

    if (-not $azureVms)
    {
        Write-ScreenInfo -Message "Get-AzVM did not return anything in $($azureRetryCount) attempts, stopping lab deployment. Azure presented us with the error: $($getazvmerror.Exception.Message)"
        throw "Get-AzVM did not return anything in $($azureRetryCount) attempts, stopping lab deployment. Azure presented us with the error: $($getazvmerror.Exception.Message)"
    }

    if ($ComputerName.Count -eq 0) { return $azureVms }
    $azureVms | Where-Object Name -in $ComputerName
}


function Get-LWAzureVMConnectionInfo
{
    param (
        [Parameter(Mandatory)]
        [AutomatedLab.Machine[]]$ComputerName
    )

    Test-LabHostConnected -Throw -Quiet

    Write-LogFunctionEntry

    $azureRetryCount = Get-LabConfigurationItem -Name AzureRetryCount

    $lab = Get-Lab -ErrorAction SilentlyContinue
    $retryCount = 5

    if (-not $lab)
    {
        Write-PSFMessage "Could not retrieve machine info for '$($ComputerName.Name -join ',')'. No lab was imported."
    }

    if (-not ((Get-AzContext).Subscription.Name -eq $lab.AzureSettings.DefaultSubscription))
    {
        Set-AzContext -Subscription $lab.AzureSettings.DefaultSubscription
    }

    $resourceGroupName = (Get-LabAzureDefaultResourceGroup).ResourceGroupName
    $azureVMs = Get-AzVM -ResourceGroupName $resourceGroupName | Where-Object Name -in $ComputerName.ResourceName
    $ips = Get-AzPublicIpAddress -ResourceGroupName $resourceGroupName -ErrorAction SilentlyContinue

    foreach ($name in $ComputerName)
    {
        $azureVM = $azureVMs | Where-Object Name -eq $name.ResourceName

        if (-not $azureVM)
        { continue }

        $net = $lab.VirtualNetworks.Where({ $_.Name -eq $name.Network[0] })
        $ip = $ips | Where-Object { $_.Tag['Vnet'] -eq $net.ResourceName }

        if (-not $ip)
        {
            $ip = $ips | Where-Object Name -eq "$($resourceGroupName)$($net.ResourceName)lbfrontendip"
        }

        if (-not $ip)
        {
            Write-ScreenInfo -Type Error -Message "No public IP address found for VM $($name.ResourceName) with tag $($net.ResourceName) or name $($resourceGroupName)$($net.ResourceName)lbfrontendip"
            continue
        }

        $result = [AutomatedLab.Azure.AzureConnectionInfo] @{
            ComputerName      = $name.Name
            DnsName           = $ip.DnsSettings.Fqdn
            HttpsName         = $ip.DnsSettings.Fqdn
            VIP               = $ip.IpAddress
            Port              = $name.LoadBalancerWinrmHttpPort
            HttpsPort         = $name.LoadBalancerWinrmHttpsPort
            RdpPort           = $name.LoadBalancerRdpPort
            SshPort           = $name.LoadBalancerSshPort
            ResourceGroupName = $azureVM.ResourceGroupName
        }

        Write-PSFMessage "Get-LWAzureVMConnectionInfo created connection info for VM '$name'"
        Write-PSFMessage "ComputerName      = $($name.Name)"
        Write-PSFMessage "DnsName           = $($ip.DnsSettings.Fqdn)"
        Write-PSFMessage "HttpsName         = $($ip.DnsSettings.Fqdn)"
        Write-PSFMessage "VIP               = $($ip.IpAddress)"
        Write-PSFMessage "Port              = $($name.LoadBalancerWinrmHttpPort)"
        Write-PSFMessage "HttpsPort         = $($name.LoadBalancerWinrmHttpsPort)"
        Write-PSFMessage "RdpPort           = $($name.LoadBalancerRdpPort)"
        Write-PSFMessage "SshPort           = $($name.LoadBalancerSshPort)"
        Write-PSFMessage "ResourceGroupName = $($azureVM.ResourceGroupName)"

        $result
    }

    Write-LogFunctionExit -ReturnValue $result
}


function Get-LWAzureVmSize
{
    [Cmdletbinding()]
    param
    (
        [Parameter(Mandatory)]
        [AutomatedLab.Machine]$Machine
    )

    $lab = Get-Lab

    if ($machine.AzureRoleSize)
    {
        $roleSize = $lab.AzureSettings.RoleSizes |
        Where-Object { $_.Name -eq $machine.AzureRoleSize }
        Write-PSFMessage -Message "Using specified role size of '$($roleSize.Name)'"
    }
    elseif ($machine.AzureProperties.RoleSize)
    {
        $roleSize = $lab.AzureSettings.RoleSizes |
        Where-Object { $_.Name -eq $machine.AzureProperties.RoleSize }
        Write-PSFMessage -Message "Using specified role size of '$($roleSize.Name)'"
    }
    elseif ($machine.AzureProperties.UseAllRoleSizes)
    {
        $DefaultAzureRoleSize = Get-LabConfigurationItem -Name DefaultAzureRoleSize
        $roleSize = $lab.AzureSettings.RoleSizes |
        Where-Object { $_.MemoryInMB -ge ($machine.Memory / 1MB) -and $_.NumberOfCores -ge $machine.Processors -and $machine.Disks.Count -le $_.MaxDataDiskCount } |
        Sort-Object -Property MemoryInMB, NumberOfCores |
        Select-Object -First 1

        Write-PSFMessage -Message "Using specified role size of '$($roleSize.InstanceSize)'. VM was configured to all role sizes but constrained to role size '$DefaultAzureRoleSize' by psd1 file"
    }
    else
    {
        $pattern = switch ($lab.AzureSettings.DefaultRoleSize)
        {
            'A' { '^Standard_A\d{1,2}(_v\d{1,3})|Basic_A\d{1,2}' }
            'AS' { '^Standard_AS\d{1,2}(_v\d{1,3})' }
            'AC' { '^Standard_AC\d{1,2}(_v\d{1,3})' }
            'D' { '^Standard_D\d{1,2}s(_v\d{1,3})' }
            'DS' { '^Standard_DS\d{1,2}(-\d\d?)?(_v\d{1,3})' }
            'DC' { '^Standard_DC\d{1,2}s(_v\d{1,3})' }
            "E" { '^Standard_E\d{1,2}s(_v\d{1,3})' }
            "EC" { '^Standard_EC\d{1,2}([a-z]+)(_cc)?(_v\d{1,3})' }
            'F' { '^Standard_F\d{1,2}s(_v\d{1,3})' }
            'G' { '^Standard_G\d{1,2}(-\d{1,2})?' }
            'GS' { '^Standard_GS\d{1,2}(-\d{1,2})?' }
            'HB' { '^Standard_HB(\d{1,3})(-\d{1,3})?rs(_v\d{1,3})' }
            'L' { '^Standard_L\d{1,2}s(_v\d{1,3})' }
            'NV' { '^Standard_NV(\d{1,3})adm?s_(V\d{3})(_v\d{1,3})' }
            'NC' { '^Standard_NC(\d{1,3})ad?s_([A|T]\d{1,3})(_v\d{1,3})'}
            default { '^Standard_DS\d{1,2}(-\d\d?)?(_v\d{1,3})' }
        }

        $roleSize = $lab.AzureSettings.RoleSizes |
            Where-Object { $_.Name -Match $pattern -and $_.Name -notlike '*promo*' } |
            Where-Object { $_.MemoryInMB -ge ($machine.Memory / 1MB) -and $_.NumberOfCores -ge $machine.Processors } |
            Where-Object { 
                if ($Machine.VmGeneration -eq 2) {
                    $_.Gen2Supported -eq $true
                }
                elseif ($Machine.VmGeneration -eq 1) {
                    $_.Gen1Supported -eq $true
                }
            } |
            Sort-Object -Property MemoryInMB, NumberOfCores, @{ Expression = { if ($_.Name -match '.+_v(?<Version>\d{1,2})') { $Matches.Version } }; Ascending = $false } |
            Select-Object -First 1

        Write-PSFMessage -Message "Using specified role size of '$($roleSize.Name)' out of role sizes '$pattern'"
    }

    $roleSize
}


function Get-LWAzureVmSnapshot
{
    param
    (
        [Parameter()]
        [Alias('VMName')]
        [string[]]
        $ComputerName,

        [Parameter()]
        [Alias('Name')]
        [string]
        $SnapshotName
    )

    Test-LabHostConnected -Throw -Quiet

    $snapshots = Get-AzSnapshot -ResourceGroupName (Get-LabAzureDefaultResourceGroup).Name -ErrorAction SilentlyContinue

    if ($SnapshotName)
    {
        $snapshots = $snapshots | Where-Object { ($_.Name -split '_')[1] -eq $SnapshotName }
    }

    if ($ComputerName)
    {
        $snapshots = $snapshots | Where-Object { ($_.Name -split '_')[0] -in $ComputerName }
    }

    $snapshots.ForEach({
            [AutomatedLab.Snapshot]::new(($_.Name -split '_')[1], ($_.Name -split '_')[0], $_.TimeCreated)
        })
}


function Get-LWAzureVMStatus
{
    param (
        [Parameter(Mandatory)]
        [string[]]$ComputerName
    )

    Test-LabHostConnected -Throw -Quiet

    #required to suporess verbose messages, warnings and errors
    Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState

    Write-LogFunctionEntry

    $azureRetryCount = Get-LabConfigurationItem -Name AzureRetryCount

    $result = @{ }
    $azureVms = Get-LWAzureVm @PSBoundParameters

    $resourceGroups = (Get-LabVM -IncludeLinux).AzureConnectionInfo.ResourceGroupName | Select-Object -Unique
    $azureVms = $azureVms | Where-Object { $_.Name -in $ComputerName -and $_.ResourceGroupName -in $resourceGroups }

    $vmTable = @{ }
    Get-LabVm -IncludeLinux | Where-Object FriendlyName -in $ComputerName | ForEach-Object { $vmTable[$_.FriendlyName] = $_.Name }

    foreach ($azureVm in $azureVms)
    {
        $vmName = if ($vmTable[$azureVm.Name]) { $vmTable[$azureVm.Name] } else { $azureVm.Name }
        if ($azureVm.PowerState -eq 'VM running')
        {
            $result.Add($vmName, 'Started')
        }
        elseif ($azureVm.PowerState -eq 'VM stopped' -or $azureVm.PowerState -eq 'VM deallocated')
        {
            $result.Add($vmName, 'Stopped')
        }
        else
        {
            $result.Add($vmName, 'Unknown')
        }
    }

    $result

    Write-LogFunctionExit
}


function Initialize-LWAzureVM
{
    [Cmdletbinding()]
    Param (
        [Parameter(Mandatory)]
        [AutomatedLab.Machine[]]$Machine
    )

    Test-LabHostConnected -Throw -Quiet
    Write-LogFunctionEntry

    $azureRetryCount = Get-LabConfigurationItem -Name AzureRetryCount
    $lab = Get-Lab
    $azvms = Get-LWAzureVm -ComputerName $Machine | Where-Object {-not $_.Tags['InitDone']}

    if ($azvms.Count -eq 0) {
        Write-ScreenInfo -Message "Azure VMs already initialized" -Type Verbose
        Write-LogFunctionExit
        return
    }

    $initScript = {
        param(
            [string]
            $UserLocale,

            [string]
            $TimeZoneId,

            [string]
            $Disks,

            [string]
            $LabSourcesPath,

            [string]
            $StorageAccountName,

            [string]
            $StorageAccountKey,

            [string[]]
            $DnsServers,

            [int]
            $WinRmMaxEnvelopeSizeKb,

            [int]
            $WinRmMaxConcurrentOperationsPerUser,

            [int]
            $WinRmMaxConnections,

            [string]
            $PublicKey,

            [string]
            $DeployDebugPath
        )

        $defaultSettings = @{
            WinRmMaxEnvelopeSizeKb              = 500
            WinRmMaxConcurrentOperationsPerUser = 1500
            WinRmMaxConnections                 = 300
        }

        $deployDebug = (New-Item -ItemType Directory -Path $ExecutionContext.InvokeCommand.ExpandString($DeployDebugPath) -Force).FullName
        $null = Start-Transcript -OutputDirectory $deployDebug
    
        Start-Service WinRm
        foreach ($setting in $defaultSettings.GetEnumerator())
        {
            if ($PSBoundParameters[$setting.Key].Value -ne $setting.Value)
            {
                $subdir = if ($setting.Key -match 'MaxEnvelope') { $null } else { 'Service\' }
                Set-Item "WSMAN:\localhost\$subdir$($setting.Key.Replace('WinRm',''))" $($PSBoundParameters[$setting.Key]) -Force
            }
        }

        Enable-PSRemoting -Force -SkipNetworkProfileCheck
        Enable-WSManCredSSP -Role Server -Force

        #region Region Settings Xml
        $regionSettings = @'
<gs:GlobalizationServices xmlns:gs="urn:longhornGlobalizationUnattend">

 <!-- user list -->
 <gs:UserList>
    <gs:User UserID="Current" CopySettingsToDefaultUserAcct="true" CopySettingsToSystemAcct="true"/>
 </gs:UserList>

 <!-- GeoID -->
 <gs:LocationPreferences>
    <gs:GeoID Value="{1}"/>
 </gs:LocationPreferences>

 <!-- system locale -->
 <gs:SystemLocale Name="{0}"/>

<!-- user locale -->
 <gs:UserLocale>
    <gs:Locale Name="{0}" SetAsCurrent="true" ResetAllSettings="true"/>
 </gs:UserLocale>

</gs:GlobalizationServices>
'@
        #endregion

        try
        {
            $geoId = [System.Globalization.RegionInfo]::new($UserLocale).GeoId
        }
        catch
        {
            $geoId = 244 #default is US
        }

        if (-not (Test-Path (Join-Path $deployDebug AL)))
        {
            $alDir = New-Item -ItemType Directory -Path (Join-Path $deployDebug AL) -Force
        }

        $alDir = Join-Path $deployDebug AL

        $tempFile = Join-Path -Path $alDir -ChildPath RegionalSettings
        $regionSettings -f $UserLocale, $geoId | Out-File -FilePath $tempFile
        $argument = 'intl.cpl,,/f:"{0}"' -f $tempFile
        control.exe $argument
        Start-Sleep -Seconds 1

        Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope LocalMachine -Force

        $idx = (Get-NetIPInterface | Where-object { $_.AddressFamily -eq "IPv4" -and $_.InterfaceAlias -like "*Ethernet*" }).ifIndex
        $dnsServer = Get-DnsClientServerAddress -InterfaceIndex $idx -AddressFamily IPv4
        Set-DnsClientServerAddress -InterfaceIndex $idx -ServerAddresses 168.63.129.16
        $release = Invoke-RestMethod -Uri 'https://api.github.com/repos/powershell/powershell/releases/latest' -UseBasicParsing -ErrorAction SilentlyContinue
        $uri = ($release.assets | Where-Object name -like '*-win-x64.msi').browser_download_url
        if (-not $uri)
        {
            $uri = 'https://github.com/PowerShell/PowerShell/releases/download/v7.2.5/PowerShell-7.2.5-win-x64.msi'
        }
    
        Invoke-WebRequest -Uri $uri -UseBasicParsing -OutFile C:\PS7.msi -ErrorAction SilentlyContinue    
        Start-Process -Wait -FilePath msiexec '/package C:\PS7.msi /quiet ADD_EXPLORER_CONTEXT_MENU_OPENPOWERSHELL=0 ENABLE_PSREMOTING=0 REGISTER_MANIFEST=0 USE_MU=0 ENABLE_MU=0' -NoNewWindow -PassThru -ErrorAction SilentlyContinue
        Remove-Item -Path C:\PS7.msi -ErrorAction SilentlyContinue

        # Configure SSHD for PowerShell Remoting alternative that also works on Linux
        if (Get-WindowsCapability -Online | Where-Object Name -like 'OpenSSH*')
        {
            Add-WindowsCapability -Online -Name OpenSSH.Server~~~~0.0.1.0 -ErrorAction SilentlyContinue
            Start-Service sshd -ErrorAction SilentlyContinue
            Set-Service -Name sshd -StartupType 'Automatic' -ErrorAction SilentlyContinue

            if (-not (Get-NetFirewallRule -Name "OpenSSH-Server-In-TCP" -ErrorAction SilentlyContinue)) 
            {
                New-NetFirewallRule -Name 'OpenSSH-Server-In-TCP' -DisplayName 'OpenSSH Server (sshd)' -Enabled True -Direction Inbound -Protocol TCP -Action Allow -LocalPort 22 -Profile Any
            }

            New-ItemProperty -Path "HKLM:\SOFTWARE\OpenSSH" -Name DefaultShell -Value "C:\Program Files\powershell\7\pwsh.exe" -PropertyType String -Force -ErrorAction SilentlyContinue
            $null = New-Item -Force -Path $alDir\SSH -ItemType Directory
            if ($PublicKey) { $PublicKey | Set-Content -Path (Join-Path -Path $alDir\SSH -ChildPath 'keys') }
            Start-Process -Wait -FilePath icacls.exe -ArgumentList "$(Join-Path -Path $alDir\SSH -ChildPath 'keys') /inheritance:r /grant ""Administrators:F"" /grant ""SYSTEM:F""" -ErrorAction SilentlyContinue
            $sshdConfig = @"
Port 22
PasswordAuthentication no
PubkeyAuthentication yes
GSSAPIAuthentication yes
AllowGroups Users Administrators
AuthorizedKeysFile c:/al/ssh/keys
Subsystem powershell c:/progra~1/powershell/7/pwsh.exe -sshs -NoLogo
"@
            $sshdConfig | Set-Content -Path (Join-Path -Path $env:ProgramData -ChildPath 'ssh/sshd_config') -ErrorAction SilentlyContinue    
            Restart-Service -Name sshd -ErrorAction SilentlyContinue    
        }

        Set-DnsClientServerAddress -InterfaceIndex $idx -ServerAddresses $dnsServer.ServerAddresses

        #Set Power Scheme to High Performance
        powercfg.exe -setactive 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c

        #Create a scheduled tasks that maps the Azure lab sources drive during each logon
        if (-not [string]::IsNullOrWhiteSpace($LabSourcesPath))
        {
            $script = @'
$output = ''
$labSourcesPath = '{0}'

$pattern = '^(OK|Unavailable) +(?<DriveLetter>\w): +\\\\automatedlab'

#remove all drive connected to an Azure LabSources share that are no longer available
$drives = net.exe use
foreach ($line in $drives)
{{
    if ($line -match $pattern)
    {{
        $output += net.exe use "$($Matches.DriveLetter):" /d
    }}
}}

$output += cmdkey.exe /add:{1} /user:{2} /pass:{3}

Start-Sleep -Seconds 1

net.exe use * {0} /u:{2} {3}

$initialErrorCode = $LASTEXITCODE
    
if ($LASTEXITCODE -eq 2) {{
    $hostName = ([uri]$labSourcesPath).Host
	$dnsRecord = Resolve-DnsName -Name $hostname | Where-Object {{ $_ -is [Microsoft.DnsClient.Commands.DnsRecord_A] }}
    $ipAddress = $dnsRecord.IPAddress
    $alternativeLabSourcesPath = $labSourcesPath.Replace($hostName, $ipAddress)
    $output += net.exe use * $alternativeLabSourcesPath /u:{2} {3}
}}

$finalErrorCode = $LASTEXITCODE

[pscustomobject]@{{
    Output = $output
    InitialErrorCode = $initialErrorCode
    FinalErrorCode = $finalErrorCode
    LabSourcesPath = $labSourcesPath
    AlternativeLabSourcesPath  = $alternativeLabSourcesPath 
}}
'@

            $cmdkeyTarget = ($LabSourcesPath -split '\\')[2]
            $script = $script -f $LabSourcesPath, $cmdkeyTarget, $StorageAccountName, $StorageAccountKey

            [pscustomobject]@{
                Path               = $LabSourcesPath
                StorageAccountName = $StorageAccountName
                StorageAccountKey  = $StorageAccountKey
            } | Export-Clixml -Path $alDir\LabSourcesStorageAccount.xml
            $script | Out-File $alDir\AzureLabSources.ps1 -Force
        }

        #set the time zone
        Set-TimeZone -Name $TimeZoneId

        reg.exe add 'HKLM\SOFTWARE\Microsoft\ServerManager\oobe' /v DoNotOpenInitialConfigurationTasksAtLogon /d 1 /t REG_DWORD /f
        reg.exe add 'HKLM\SOFTWARE\Microsoft\ServerManager' /v DoNotOpenServerManagerAtLogon /d 1 /t REG_DWORD /f
        reg.exe add 'HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' /v EnableFirstLogonAnimation /d 0 /t REG_DWORD /f
        reg.exe add 'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' /v FilterAdministratorToken /t REG_DWORD /d 0 /f
        reg.exe add 'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' /v EnableLUA /t REG_DWORD /d 0 /f
        reg.exe add 'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' /v ConsentPromptBehaviorAdmin /t REG_DWORD /d 0 /f
        reg.exe add 'HKLM\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A7-37EF-4b3f-8CFC-4F3A74704073}' /v IsInstalled /t REG_DWORD /d 0 /f #disable admin IE Enhanced Security Configuration
        reg.exe add 'HKLM\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A8-37EF-4b3f-8CFC-4F3A74704073}' /v IsInstalled /t REG_DWORD /d 0 /f #disable user IE Enhanced Security Configuration
        reg.exe add 'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run' /v BgInfo /t REG_SZ /d "$alDir\BgInfo.exe $alDir\BgInfo.bgi /Timer:0 /nolicprompt" /f

        #turn off the Windows firewall
        Set-NetFirewallProfile -All -Enabled False -PolicyStore PersistentStore

        if ($DnsServers.Count -gt 0)
        {
            Write-Verbose "Configuring $($DnsServers.Count) DNS Servers"
            $idx = (Get-NetIPInterface | Where-object { $_.AddressFamily -eq "IPv4" -and $_.InterfaceAlias -like "*Ethernet*" }).ifIndex
            Set-DnsClientServerAddress -InterfaceIndex $idx -ServerAddresses $DnsServers
        }

        #Add *.windows.net to Local Intranet Zone
        $path = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\windows.net'
        New-Item -Path $path -Force
        New-ItemProperty $path -Name http -Value 1 -Type DWORD
        New-ItemProperty $path -Name file -Value 1 -Type DWORD

        #Add *.azure.com to Local Intranet Zone
        $path = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\azure.com'
        New-Item -Path $path -Force
        New-ItemProperty $path -Name http -Value 1 -Type DWORD
        New-ItemProperty $path -Name https -Value 1 -Type DWORD
        New-ItemProperty $path -Name file -Value 1 -Type DWORD

        if (-not $Disks) { $null = try { Stop-Transcript -ErrorAction Stop } catch { }; return }
        
        # Azure InvokeRunAsCommand is not very clever, so we sent the stuff as JSON
        $Disks | Set-Content -Path $alDir\disks.json
        [object[]] $diskObjects = $Disks | ConvertFrom-Json
        Write-Verbose -Message "Disk count for $env:COMPUTERNAME`: $($diskObjects.Count)"
        foreach ($diskObject in $diskObjects.Where({ -not $_.SkipInitialization }))
        {
            $disk = Get-Disk | Where-Object Location -like "*LUN $($diskObject.LUN)"
            $disk | Set-Disk -IsReadOnly $false
            $disk | Set-Disk -IsOffline $false
            $disk | Initialize-Disk -PartitionStyle GPT
            $party = if ($diskObject.DriveLetter)
            {
                $disk | New-Partition -UseMaximumSize -DriveLetter $diskObject.DriveLetter
            }
            else
            {
                $disk | New-Partition -UseMaximumSize -AssignDriveLetter
            }
            $party | Format-Volume -Force -UseLargeFRS:$diskObject.UseLargeFRS -AllocationUnitSize $diskObject.AllocationUnitSize -NewFileSystemLabel $diskObject.Label
        }

        $null = try { Stop-Transcript -ErrorAction Stop } catch { }
    }

    $initScriptFile = New-Item -ItemType File -Path (Join-Path -Path ([IO.Path]::GetTempPath()) -ChildPath "$($Lab.Name)vminit.ps1") -Force
    $initScript.ToString() | Set-Content -Path $initScriptFile -Force

    # Configure AutoShutdown
    if ($lab.AzureSettings.AutoShutdownTime)
    {
        $time = $lab.AzureSettings.AutoShutdownTime
        $tz = if (-not $lab.AzureSettings.AutoShutdownTimeZone) { Get-TimeZone } else { Get-TimeZone -Id $lab.AzureSettings.AutoShutdownTimeZone }
        Write-ScreenInfo -Message "Configuring auto-shutdown of all VMs daily at $($time) in timezone $($tz.Id)"
        Enable-LWAzureAutoShutdown -ComputerName (Get-LabVm -IncludeLinux | Where-Object Name -notin $machineSpecific.Name) -Time $time -TimeZone $tz.Id -Wait
    }

    $machineSpecific = Get-LabVm -SkipConnectionInfo -IncludeLinux | Where-Object {
        $_.AzureProperties.ContainsKey('AutoShutdownTime')
    }

    foreach ($machine in $machineSpecific)
    {
        $time = $machine.AzureProperties.AutoShutdownTime
        $tz = if (-not $machine.AzureProperties.AutoShutdownTimezoneId) { Get-TimeZone } else { Get-TimeZone -Id $machine.AzureProperties.AutoShutdownTimezoneId }
        Write-ScreenInfo -Message "Configure shutdown of $machine daily at $($time) in timezone $($tz.Id)"
        Enable-LWAzureAutoShutdown -ComputerName $machine -Time $time -TimeZone $tz.Id -Wait
    }

    Write-ScreenInfo -Message 'Configuring localization and additional disks' -TaskStart -NoNewLine
    if (-not $lab.AzureSettings.IsAzureStack) { $labsourcesStorage = Get-LabAzureLabSourcesStorage }
    $jobs = [System.Collections.ArrayList]::new()
    foreach ($m in ($Machine | Where-Object OperatingSystemType -eq 'Windows'))
    {
        [string[]]$DnsServers = ($m.NetworkAdapters | Where-Object { $_.VirtualSwitch.Name -eq $Lab.Name }).Ipv4DnsServers.AddressAsString
        $azVmDisks = (Get-AzVm -Name $m.ResourceName -ResourceGroupName $lab.AzureSettings.DefaultResourceGroup.ResourceGroupName).StorageProfile.DataDisks
        foreach ($machDisk in $m.Disks)
        {
            $machDisk.Lun = $azVmDisks.Where({ $_.Name -eq $machDisk.Name }).Lun
        }
        
        $diskJson = $m.disks | ConvertTo-Json -Compress

        $scriptParam = @{
            UserLocale                          = $m.UserLocale
            TimeZoneId                          = $m.TimeZone
            WinRmMaxEnvelopeSizeKb              = Get-LabConfigurationItem -Name WinRmMaxEnvelopeSizeKb
            WinRmMaxConcurrentOperationsPerUser = Get-LabConfigurationItem -Name WinRmMaxConcurrentOperationsPerUser
            WinRmMaxConnections                 = Get-LabConfigurationItem -Name WinRmMaxConnections
            DeployDebugPath                     = $AL_DeployDebugFolder
        }
        $azsArgumentLine = '-UserLocale "{0}" -TimeZoneId "{1}" -WinRmMaxEnvelopeSizeKb {2} -WinRmMaxConcurrentOperationsPerUser {3} -WinRmMaxConnections {4}' -f $m.UserLocale, $m.TimeZone, (Get-LabConfigurationItem -Name WinRmMaxEnvelopeSizeKb), (Get-LabConfigurationItem -Name WinRmMaxConcurrentOperationsPerUser), (Get-LabConfigurationItem -Name WinRmMaxConnections)

        if ($DnsServers.Count -gt 0)
        {
            $scriptParam.DnsServers = $DnsServers
            $azsArgumentLine += ' -DnsServers "{0}"' -f ($DnsServers -join '","')
        }

        if ($m.SshPublicKey)
        {
            $scriptParam.PublicKey = $m.SshPublicKey
            $azsArgumentLine += ' -PublicKey "{0}"' -f $m.SshPublicKey
        }

        if ($diskJson)
        {
            $scriptParam.Disks = $diskJson
            $azsArgumentLine += " -Disks '{0}'" -f $diskJson
        }

        if ($labsourcesStorage)
        {            
            $scriptParam.LabSourcesPath = $labsourcesStorage.Path
            $scriptParam.StorageAccountName = $labsourcesStorage.StorageAccountName
            $scriptParam.StorageAccountKey = $labsourcesStorage.StorageAccountKey
            $azsArgumentLine += '-LabSourcesPath {0} -StorageAccountName {1} -StorageAccountKey {2}' -f $labsourcesStorage.Path, $labsourcesStorage.StorageAccountName, $labsourcesStorage.StorageAccountKey
        }

        if ($m.IsDomainJoined)
        {
            $domain = $lab.Domains | Where-Object Name -eq $m.DomainName
        }

        # Azure Stack - Create temporary storage account to upload script and use extension - sad, but true.
        if ($Lab.AzureSettings.IsAzureStack)
        {
            $sa = Get-AzStorageAccount -ResourceGroupName $lab.AzureSettings.DefaultResourceGroup.ResourceGroupName -ErrorAction SilentlyContinue
            if (-not $sa)
            {
                $sa = New-AzStorageAccount -Name "cse$(-join (1..10 | % {[char](Get-Random -Min 97 -Max 122)}))" -ResourceGroupName $lab.AzureSettings.DefaultResourceGroup.ResourceGroupName -SkuName Standard_LRS -Kind Storage -Location (Get-LabAzureDefaultLocation).Location
            }

            $co = $sa | Get-AzStorageContainer -Name customscriptextension -ErrorAction SilentlyContinue
            if (-not $co)
            {
                $co = $sa | New-AzStorageContainer -Name customscriptextension
            }

            $content = Set-AzStorageBlobContent -File $initScriptFile -CloudBlobContainer $co.CloudBlobContainer -Blob $(Split-Path -Path $initScriptFile -Leaf) -Context $sa.Context -Force -ErrorAction Stop
            $token = New-AzStorageBlobSASToken -CloudBlob $content.ICloudBlob -StartTime (Get-Date) -ExpiryTime $(Get-Date).AddHours(1) -Protocol HttpsOnly -Context $sa.Context -Permission r -ErrorAction Stop
            $uri = '{0}{1}/{2}{3}' -f $co.Context.BlobEndpoint, 'customscriptextension', $(Split-Path -Path $initScriptFile -Leaf), $token
            [version] $typehandler = (Get-AzVMExtensionImage -PublisherName Microsoft.Compute -Type CustomScriptExtension -Location (Get-LabAzureDefaultLocation).Location | Sort-Object { [version]$_.Version } | Select-Object -Last 1).Version
            
            $extArg = @{
                ResourceGroupName  = $lab.AzureSettings.DefaultResourceGroup.ResourceGroupName
                VMName             = $m.ResourceName
                FileUri            = $uri
                TypeHandlerVersion = '{0}.{1}' -f $typehandler.Major, $typehandler.Minor
                Name               = 'initcustomizations'
                Location           = (Get-LabAzureDefaultLocation).Location
                Run                = Split-Path -Path $initScriptFile -Leaf
                Argument           = $azsArgumentLine
                NoWait             = $true
            }
            $Null = Set-AzVMCustomScriptExtension @extArg
        }
        else
        {
            $null = $jobs.Add((Invoke-AzVMRunCommand -ResourceGroupName $lab.AzureSettings.DefaultResourceGroup.ResourceGroupName -VMName $m.ResourceName -ScriptPath $initScriptFile -Parameter $scriptParam -CommandId 'RunPowerShellScript' -ErrorAction Stop -AsJob))
        }
    }


    $initScriptLinux = @'
sudo sed -i 's|[#]*GSSAPIAuthentication yes|GSSAPIAuthentication yes|g' /etc/ssh/sshd_config
sudo sed -i 's|[#]*PasswordAuthentication yes|PasswordAuthentication no|g' /etc/ssh/sshd_config
sudo sed -i 's|[#]*PubkeyAuthentication yes|PubkeyAuthentication yes|g' /etc/ssh/sshd_config
if [ -n "$(sudo cat /etc/ssh/sshd_config | grep 'Subsystem powershell')" ]; then
    echo "PowerShell subsystem configured"
else
    echo "Subsystem powershell /usr/bin/pwsh -sshs -NoLogo -NoProfile" | sudo tee --append /etc/ssh/sshd_config
fi
sudo mkdir -p /usr/local/share/powershell 2>/dev/null
sudo chmod 777 -R /usr/local/share/powershell

if [ -n "$(which apt 2>/dev/null)" ]; then
    curl -sSL https://packages.microsoft.com/keys/microsoft.asc | sudo apt-key add -
    curl -sSL https://packages.microsoft.com/keys/microsoft.asc | sudo tee /etc/apt/trusted.gpg.d/microsoft.asc
    sudo apt-get update
    sudo apt-get install -y wget apt-transport-https software-properties-common
    wget -q "https://packages.microsoft.com/config/ubuntu/$(lsb_release -rs)/packages-microsoft-prod.deb"
    sudo dpkg -i packages-microsoft-prod.deb
    sudo apt-get update
    sudo apt-get install -y powershell
    sudo apt-get install -y openssl omi omi-psrp-server
    sudo apt-get install -y oddjob oddjob-mkhomedir sssd adcli krb5-workstation realmd samba-common samba-common-tools authselect-compat openssh-server
elif [ -n "$(which yum 2>/dev/null)" ]; then
    sudo rpm -Uvh "https://packages.microsoft.com/config/rhel/$(sudo cat /etc/redhat-release | grep -oP "(\d)" | head -1)/packages-microsoft-prod.rpm"
    sudo yum install -y powershell
    sudo yum install -y openssl omi omi-psrp-server
    sudo yum install -y oddjob oddjob-mkhomedir sssd adcli krb5-workstation realmd samba-common samba-common-tools authselect-compat openssh-server
elif [ -n "$(which dnf 2>/dev/null)" ]; then
    sudo rpm -Uvh https://packages.microsoft.com/config/rhel/$(sudo cat /etc/redhat-release | grep -oP "(\d)" | head -1)/packages-microsoft-prod.rpm
    sudo dnf install -y powershell
    sudo dnf install -y openssl omi omi-psrp-server
    sudo dnf install -y oddjob oddjob-mkhomedir sssd adcli krb5-workstation realmd samba-common samba-common-tools authselect-compat openssh-server
fi
sudo systemctl restart sshd
'@
    $linuxInitFiles = foreach ($m in ($Machine | Where-Object OperatingSystemType -eq 'Linux'))
    {
        if ($Lab.AzureSettings.IsAzureStack)
        {
            Write-ScreenInfo -Type Warning -Message 'Linux VMs not yet implemented on Azure Stack, sorry.'
            continue
        }

        $initScriptFileLinux = New-Item -ItemType File -Path (Join-Path -Path ([IO.Path]::GetTempPath()) -ChildPath "$($Lab.Name)$($m.Name)vminitlinux.bash") -Force
        $initScriptLinux | Set-Content -Path $initScriptFileLinux -Force
        $initScriptFileLinux

        $null = $jobs.Add((Invoke-AzVMRunCommand -ResourceGroupName $lab.AzureSettings.DefaultResourceGroup.ResourceGroupName -VMName $m.ResourceName -ScriptPath $initScriptFileLinux.FullName -CommandId 'RunShellScript' -ErrorAction Stop -AsJob))
    }

    if ($jobs)
    {
        Wait-LWLabJob -Job $jobs -ProgressIndicator 5 -Timeout 30 -NoDisplay
    }

    $initScriptFile | Remove-Item -ErrorAction SilentlyContinue
    $linuxInitFiles | Copy-Item -Destination $Lab.LabPath
    $linuxInitFiles | Remove-Item -ErrorAction SilentlyContinue

    # And once again for all the VMs that for some unknown reason did not *really* execute the RunCommand
    if (Get-Command ssh -ErrorAction SilentlyContinue)
    {
        Install-LabSshKnownHost
        foreach ($m in ($Machine | Where-Object {$_.OperatingSystemType -eq 'Linux' -and $_.SshPrivateKeyPath}))
        {
            $ci = $m.AzureConnectionInfo
            $null = ssh -p $ci.SshPort "automatedlab@$($ci.DnsName)" -i $m.SshPrivateKeyPath $initScriptLinux 2>$null
        }
    }

    # Wait for VM extensions to be "done"
    if ($lab.AzureSettings.IsAzureStack)
    {
        $extensionStatuus = Get-LabVm -IncludeLinux | Foreach-Object { Get-AzVMCustomScriptExtension -ResourceGroupName $lab.AzureSettings.DefaultResourceGroup.ResourceGroupName -VMName $_.ResourceName -Name initcustomizations -ErrorAction SilentlyContinue }
        $start = Get-Date
        $timeout = New-TimeSpan -Minutes 5
        while (($extensionStatuus.ProvisioningState -contains 'Updating' -or $extensionStatuus.ProvisioningState -contains 'Creating') -and ((Get-Date) - $start) -lt $timeout)
        {
            Start-Sleep -Seconds 5
            $extensionStatuus = Get-LabVm -IncludeLinux | Foreach-Object { Get-AzVMCustomScriptExtension -ResourceGroupName $lab.AzureSettings.DefaultResourceGroup.ResourceGroupName -VMName $_.ResourceName -Name initcustomizations -ErrorAction SilentlyContinue }
        }

        foreach ($network in $Lab.VirtualNetworks)
        {
            if ($network.DnsServers.Count -eq 0) { continue }
            $vnet = Get-AzVirtualNetwork -Name $network.ResourceName -ResourceGroupName $lab.AzureSettings.DefaultResourceGroup.ResourceGroupName
            $vnet.dhcpOptions.dnsServers = [string[]]($network.DnsServers.AddressAsString)
            $null = $vnet | Set-AzVirtualNetwork
        }
    }

    $deployDebug = Invoke-LabCommand -ComputerName ($Machine | Where OperatingSystemType -eq 'Windows') -Variable (Get-Variable -Name AL_DeployDebugFolder) -PassThru -ScriptBlock {
        (Get-Item -Path "$($ExecutionContext.InvokeCommand.ExpandString($DeployDebugPath))/AL").FullName
    } | Select-Object -First 1
    Copy-LabFileItem -Path (Get-ChildItem -Path "$((Get-Module -Name AutomatedLabCore)[0].ModuleBase)\Tools\HyperV\*") -DestinationFolderPath $deployDebug -ComputerName ($Machine | Where OperatingSystemType -eq 'Windows') -UseAzureLabSourcesOnAzureVm $false
    $sessions = if ($PSVersionTable.PSVersion -ge [System.Version]'7.0')
    {
        New-LabPSSession $Machine
    }
    else
    {
        Write-ScreenInfo -Type Warning -Message "Skipping copy of AutomatedLab.Common to Linux VMs as Windows PowerShell is used on the host and not PowerShell 7+."
        New-LabPSSession ($Machine | Where-Object OperatingSystemType -eq 'Windows')
    }

    Send-ModuleToPSSession -Module (Get-Module -ListAvailable -Name AutomatedLab.Common | Select-Object -First 1) -Session $sessions -IncludeDependencies -Force

    $null = $azvms | ForEach-Object {
        $_.Tags['InitDone'] = Get-Date -Format u
        $_ | Update-AzVM
    }

    Write-ScreenInfo -Message 'Finished' -TaskEnd

    Write-ScreenInfo -Message 'Stopping all new machines except domain controllers'
    $machinesToStop = $Machine | Where-Object { $_.Roles.Name -notcontains 'RootDC' -and $_.Roles.Name -notcontains 'FirstChildDC' -and $_.Roles.Name -notcontains 'DC' -and $_.IsDomainJoined }
    if ($machinesToStop)
    {
        Stop-LWAzureVM -ComputerName $machinesToStop -StayProvisioned $true
        Wait-LabVMShutdown -ComputerName $machinesToStop
    }

    if ($machinesToStop)
    {
        Write-ScreenInfo -Message "$($Machine.Count) new Azure machines were configured. Some machines were stopped as they are not to be domain controllers '$($machinesToStop -join ', ')'"
    }
    else
    {
        Write-ScreenInfo -Message "($($Machine.Count)) new Azure machines were configured"
    }

    Write-PSFMessage "Removing all sessions after VmInit"
    Remove-LabPSSession

    Write-LogFunctionExit
}


function Mount-LWAzureIsoImage
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseCompatibleCmdlets", "", Justification = "Not relevant, used in Invoke-LabCommand")]
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory, Position = 0)]
        [string[]]
        $ComputerName,

        [Parameter(Mandatory, Position = 1)]
        [string]
        $IsoPath,

        [switch]$PassThru
    )

    Test-LabHostConnected -Throw -Quiet

    $azureRetryCount = Get-LabConfigurationItem -Name AzureRetryCount
    $azureIsoPath = $IsoPath -replace '/', '\' -replace 'https:'
    # ISO file should already exist on Azure storage share, as it was initially retrieved from there as well.

    # Path is local (usually Azure Stack which has no storage file shares)
    if (-not (Test-LabPathIsOnLabAzureLabSourcesStorage -Path $azureIsoPath))
    {
        Write-ScreenInfo -type Info -Message "Copying $azureIsoPath to $($ComputerName -join ',')"
        Copy-LabFileItem -Path $azureIsoPath -ComputerName $ComputerName -DestinationFolderPath C:\ALMounts
        $result = Invoke-LabCommand -ActivityName "Mounting $(Split-Path $azureIsoPath -Leaf) on $($ComputerName -join ',')" -ComputerName $ComputerName -ScriptBlock {
            $drive = Mount-DiskImage -ImagePath C:\ALMounts\$(Split-Path -Leaf -Path $azureIsoPath) -StorageType ISO -PassThru | Get-Volume
            $drive | Add-Member -MemberType NoteProperty -Name DriveLetter -Value ($drive.CimInstanceProperties.Item('DriveLetter').Value + ":") -Force
            $drive | Add-Member -MemberType NoteProperty -Name InternalComputerName -Value $env:COMPUTERNAME -Force
            $drive | Select-Object -Property *
        } -Variable (Get-Variable azureIsoPath) -PassThru:$PassThru.IsPresent

        if ($PassThru.IsPresent) { return $result } else { return }
    }

    Invoke-LabCommand -ActivityName "Mounting $(Split-Path $azureIsoPath -Leaf) on $($ComputerName -join ',')" -ComputerName $ComputerName -ScriptBlock {

        if (-not (Test-Path -Path $azureIsoPath))
        {
            throw "'$azureIsoPath' is not accessible."
        }

        $drive = Mount-DiskImage -ImagePath $azureIsoPath -StorageType ISO -PassThru | Get-Volume
        $drive | Add-Member -MemberType NoteProperty -Name DriveLetter -Value ($drive.CimInstanceProperties.Item('DriveLetter').Value + ":") -Force
        $drive | Add-Member -MemberType NoteProperty -Name InternalComputerName -Value $env:COMPUTERNAME -Force
        $drive | Select-Object -Property *

    } -ArgumentList $azureIsoPath -Variable (Get-Variable -Name azureIsoPath) -PassThru:$PassThru
}


function New-LabAzureResourceGroupDeployment
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory)]
        [AutomatedLab.Lab]
        $Lab,

        [Parameter()]
        [switch]
        $PassThru,

        [Parameter()]
        [switch]
        $Wait
    )

    Write-LogFunctionEntry

    $template = @{
        '$schema'      = "https://schema.management.azure.com/schemas/2019-04-01/deploymentTemplate.json#"
        contentVersion = '1.0.0.0'  
        parameters     = @{ }
        resources      = @()
    }

    # The handy providers() function was deprecated and the latest provider APIs started getting error-prone and unpredictable
    # The following list was generated on Jul 12 2022
    $apiVersions = if (Get-LabConfigurationItem -Name UseLatestAzureProviderApi)
    {
        $providers = Get-AzResourceProvider -Location $lab.AzureSettings.DefaultLocation.Location -ErrorAction SilentlyContinue | Where-Object RegistrationState -eq 'Registered'
        $provHash = @{
            NicApi            = (($providers | Where-Object ProviderNamespace -eq 'Microsoft.Network').ResourceTypes | Where-Object ResourceTypeName -eq 'networkInterfaces').ApiVersions[0] # 2022-01-01
            DiskApi           = (($providers | Where-Object ProviderNamespace -eq 'Microsoft.Compute').ResourceTypes | Where-Object ResourceTypeName -eq 'disks').ApiVersions[0] # 2022-01-01
            LoadBalancerApi   = (($providers | Where-Object ProviderNamespace -eq 'Microsoft.Network').ResourceTypes | Where-Object ResourceTypeName -eq 'loadBalancers').ApiVersions[0] # 2022-01-01
            PublicIpApi       = (($providers | Where-Object ProviderNamespace -eq 'Microsoft.Network').ResourceTypes | Where-Object ResourceTypeName -eq 'publicIpAddresses').ApiVersions[0] # 2022-01-01
            VirtualNetworkApi = (($providers | Where-Object ProviderNamespace -eq 'Microsoft.Network').ResourceTypes | Where-Object ResourceTypeName -eq 'virtualNetworks').ApiVersions[0] # 2022-01-01
            NsgApi            = (($providers | Where-Object ProviderNamespace -eq 'Microsoft.Network').ResourceTypes | Where-Object ResourceTypeName -eq 'networkSecurityGroups').ApiVersions[0] # 2022-01-01
            VmApi             = (($providers | Where-Object ProviderNamespace -eq 'Microsoft.Compute').ResourceTypes | Where-Object ResourceTypeName -eq 'virtualMachines').ApiVersions[1] # 2022-03-01
        }
        if (-not $lab.AzureSettings.IsAzureStack)
        {
            $provHash.BastionHostApi = (($providers | Where-Object ProviderNamespace -eq 'Microsoft.Network').ResourceTypes | Where-Object ResourceTypeName -eq 'bastionHosts').ApiVersions[0] # 2022-01-01
        }
        if ($lab.AzureSettings.IsAzureStack)
        {
            $provHash.VmApi = (($providers | Where-Object ProviderNamespace -eq 'Microsoft.Compute').ResourceTypes | Where-Object ResourceTypeName -eq 'virtualMachines').ApiVersions[0]
        }
        $provHash
    }
    elseif ($Lab.AzureSettings.IsAzureStack)
    {
        @{
            NicApi            = '2018-11-01'
            DiskApi           = '2018-11-01'
            LoadBalancerApi   = '2018-11-01'
            PublicIpApi       = '2018-11-01'
            VirtualNetworkApi = '2018-11-01'
            NsgApi            = '2018-11-01'
            VmApi             = '2020-06-01'
        }
    }
    else
    {
        @{
            NicApi            = '2022-01-01'
            DiskApi           = '2022-01-01'
            LoadBalancerApi   = '2022-01-01'
            PublicIpApi       = '2022-01-01'
            VirtualNetworkApi = '2022-01-01'
            BastionHostApi    = '2022-01-01'
            NsgApi            = '2022-01-01'
            VmApi             = '2022-03-01'
        }
    }
    
    #region Network Security Group
    Write-ScreenInfo -Type Verbose -Message 'Adding network security group to template, enabling traffic to ports 3389,5985,5986,22 for VMs behind load balancer'
    [string[]]$allowedIps = (Get-LabVm -IncludeLinux).AzureProperties["LoadBalancerAllowedIp"] | Foreach-Object { $_ -split '\s*[,;]\s*' } | Where-Object { -not [string]::IsNullOrWhitespace($_) }
    $nsg = @{
        type       = "Microsoft.Network/networkSecurityGroups"
        apiVersion = $apiVersions['NsgApi']
        name       = "nsg"
        location   = "[resourceGroup().location]"
        tags       = @{ 
            AutomatedLab = $Lab.Name
            CreationTime = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        }
        properties = @{
            securityRules = @(
                # Necessary mgmt ports for AutomatedLab
                @{
                    name       = "NecessaryPorts"
                    properties = @{
                        protocol                   = "TCP"
                        sourcePortRange            = "*"
                        sourceAddressPrefix        = if ($allowedIps) { $null } else { "*" }
                        destinationAddressPrefix   = "VirtualNetwork"
                        access                     = "Allow"
                        priority                   = 100
                        direction                  = "Inbound"
                        sourcePortRanges           = @()
                        destinationPortRanges      = @(
                            "22"
                            "3389"
                            "5985"
                            "5986"
                        )
                        sourceAddressPrefixes      = @()
                        destinationAddressPrefixes = @()
                    }
                }
                # Rules for bastion host deployment - always included to be able to deploy bastion at a later stage
                @{
                    name       = "BastionIn"
                    properties = @{
                        protocol                   = "TCP"
                        sourcePortRange            = "*"
                        sourceAddressPrefix        = if ($allowedIps) { $null } else { "*" }
                        destinationAddressPrefix   = "*"
                        access                     = "Allow"
                        priority                   = 101
                        direction                  = "Inbound"
                        sourcePortRanges           = @()
                        destinationPortRanges      = @(
                            "443"
                        )
                        sourceAddressPrefixes      = @()
                        destinationAddressPrefixes = @()
                    }
                }
                if (-not $Lab.AzureSettings.IsAzureStack)
                {
                    @{
                        name       = "BastionMgmtOut"
                        properties = @{
                            protocol                   = "TCP"
                            sourcePortRange            = "*"
                            sourceAddressPrefix        = "*"
                            destinationAddressPrefix   = "AzureCloud"
                            access                     = "Allow"
                            priority                   = 100
                            direction                  = "Outbound"
                            sourcePortRanges           = @()
                            destinationPortRanges      = @(
                                "443"
                            )
                            sourceAddressPrefixes      = @()
                            destinationAddressPrefixes = @()
                        }
                    }
                    @{
                        name       = "BastionRdsOut"
                        properties = @{
                            protocol                   = "TCP"
                            sourcePortRange            = "*"
                            sourceAddressPrefix        = "*"
                            destinationAddressPrefix   = "VirtualNetwork"
                            access                     = "Allow"
                            priority                   = 101
                            direction                  = "Outbound"
                            sourcePortRanges           = @()
                            destinationPortRanges      = @(
                                "3389"
                                "22"
                            )
                            sourceAddressPrefixes      = @()
                            destinationAddressPrefixes = @()
                        }
                    }
                }
            )
        }
    }

    if ($allowedIps)
    {
        $nsg.properties.securityrules | Where-Object { $_.properties.direction -eq 'Inbound' } | Foreach-object { $_.properties.sourceAddressPrefixes = $allowedIps }
    }
    $template.resources += $nsg
    #endregion

    #region Wait for availability of Bastion
    if ($Lab.AzureSettings.AllowBastionHost -and -not $lab.AzureSettings.IsAzureStack)
    {
        $bastionFeature = Get-AzProviderFeature -FeatureName AllowBastionHost -ProviderNamespace Microsoft.Network
        while (($bastionFeature).RegistrationState -ne 'Registered')
        {
            if ($bastionFeature.RegistrationState -eq 'NotRegistered')
            {
                $null = Register-AzProviderFeature -FeatureName AllowBastionHost -ProviderNamespace Microsoft.Network
                $null = Register-AzProviderFeature -FeatureName bastionShareableLink -ProviderNamespace Microsoft.Network
            }

            Start-Sleep -Seconds 5
            Write-ScreenInfo -Type Verbose -Message "Waiting for registration of bastion host feature. Current status: $(($bastionFeature).RegistrationState)"
            $bastionFeature = Get-AzProviderFeature -FeatureName AllowBastionHost -ProviderNamespace Microsoft.Network
        }
    }

    $vnetCount = 0
    $loadbalancers = @{}
    foreach ($network in $Lab.VirtualNetworks)
    {
        #region VNet
        Write-ScreenInfo -Type Verbose -Message ('Adding vnet {0} ({1}) to template' -f $network.ResourceName, $network.AddressSpace)
        $vNet = @{
            type       = "Microsoft.Network/virtualNetworks"
            apiVersion = $apiVersions['VirtualNetworkApi']
            tags       = @{ 
                AutomatedLab = $Lab.Name
                CreationTime = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            }
            dependsOn  = @(
                "[resourceId('Microsoft.Network/networkSecurityGroups', 'nsg')]"
            )
            name       = $network.ResourceName
            location   = "[resourceGroup().location]"
            properties = @{
                addressSpace = @{
                    addressPrefixes = @(
                        $network.AddressSpace.ToString()
                    )
                }
                subnets      = @()
                dhcpOptions  = @{
                    dnsServers = @()
                }
            }
        }

        if (-not $network.Subnets)
        {
            Write-ScreenInfo -Type Verbose -Message ('Adding default subnet ({0}) to VNet' -f $network.AddressSpace)
            $vnet.properties.subnets += @{
                name       = "default"
                properties = @{
                    addressPrefix        = $network.AddressSpace.ToString()
                    networkSecurityGroup = @{
                        id = "[resourceId('Microsoft.Network/networkSecurityGroups', 'nsg')]"
                    }
                }
            }
        }

        foreach ($subnet in $network.Subnets)
        {
            Write-ScreenInfo -Type Verbose -Message ('Adding subnet {0} ({1}) to VNet' -f $subnet.Name, $subnet.AddressSpace)
            $vnet.properties.subnets += @{
                name       = $subnet.Name
                properties = @{
                    addressPrefix        = $subnet.AddressSpace.ToString()
                    networkSecurityGroup = @{
                        id = "[resourceId('Microsoft.Network/networkSecurityGroups', 'nsg')]"
                    }
                }
            }
        }

        if ($Lab.AzureSettings.AllowBastionHost -and -not $lab.AzureSettings.IsAzureStack)
        {
            if ($network.Subnets.Name -notcontains 'AzureBastionSubnet')
            {
                $sourceMask = $network.AddressSpace.Cidr
                $sourceMaskIp = $network.AddressSpace.NetMask
                $sourceRange = Get-NetworkRange -IPAddress $network.AddressSpace.IpAddress.AddressAsString -SubnetMask $network.AddressSpace.NetMask
                $sourceInfo = Get-NetworkSummary -IPAddress $network.AddressSpace.IpAddress.AddressAsString -SubnetMask $network.AddressSpace.NetMask
                $superNetMask = $sourceMask - 1
                $superNetIp = $network.AddressSpace.IpAddress.AddressAsString
                $superNet = [AutomatedLab.VirtualNetwork]::new()
                $superNet.AddressSpace = '{0}/{1}' -f $superNetIp, $superNetMask
                $superNetInfo = Get-NetworkSummary -IPAddress $superNet.AddressSpace.IpAddress.AddressAsString -SubnetMask $superNet.AddressSpace.NetMask

                foreach ($address in (Get-NetworkRange -IPAddress $superNet.AddressSpace.IpAddress.AddressAsString -SubnetMask $superNet.AddressSpace.NetMask))
                {
                    if ($address -in @($sourceRange + $sourceInfo.Network + $sourceInfo.Broadcast))
                    {
                        continue
                    }

                    $bastionNet = [AutomatedLab.VirtualNetwork]::new()
                    $bastionNet.AddressSpace = '{0}/{1}' -f $address, $sourceMask
                    break
                }

                $vNet.properties.addressSpace.addressPrefixes = @(
                    $superNet.AddressSpace.ToString()
                )
                $vnet.properties.subnets += @{
                    name       = 'AzureBastionSubnet'
                    properties = @{
                        addressPrefix        = $bastionNet.AddressSpace.ToString()
                        networkSecurityGroup = @{
                            id = "[resourceId('Microsoft.Network/networkSecurityGroups', 'nsg')]"
                        }
                    }
                }
            }

            $dnsLabel = "[concat('azbastion', uniqueString(resourceGroup().id))]"
            Write-ScreenInfo -Type Verbose -Message ('Adding Azure bastion public static IP with DNS label {0} to template' -f $dnsLabel)
            $template.resources +=
            @{
                apiVersion = $apiVersions['PublicIpApi']
                tags       = @{ 
                    AutomatedLab = $Lab.Name
                    CreationTime = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
                }
                type       = "Microsoft.Network/publicIPAddresses"
                name       = "$($vnetCount)bip"
                location   = "[resourceGroup().location]"
                properties = @{
                    publicIPAllocationMethod = "static"
                    dnsSettings              = @{
                        domainNameLabel = $dnsLabel
                    }
                }
                sku        = @{
                    name = if ($Lab.AzureSettings.IsAzureStack) { 'Basic' } else { 'Standard' }
                }
            }

            $template.resources += @{
                apiVersion = $apiVersions['BastionHostApi']
                type       = "Microsoft.Network/bastionHosts"
                name       = "bastion$vnetCount"
                tags       = @{ 
                    AutomatedLab = $Lab.Name
                    CreationTime = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
                }
                location   = "[resourceGroup().location]"
                dependsOn  = @(
                    "[resourceId('Microsoft.Network/virtualNetworks', '$($network.ResourceName)')]"
                    "[resourceId('Microsoft.Network/publicIPAddresses', '$($vnetCount)bip')]"
                )
                properties = @{
                    ipConfigurations = @(
                        @{
                            name       = "IpConf"
                            properties = @{
                                subnet          = @{
                                    id = "[resourceId('Microsoft.Network/virtualNetworks/subnets', '$($network.ResourceName)','AzureBastionSubnet')]"
                                }
                                publicIPAddress = @{
                                    id = "[resourceId('Microsoft.Network/publicIPAddresses', '$($vnetCount)bip')]"
                                }
                            }
                        }
                    )
                }
            }
        }

        $template.resources += $vNet
        #endregion

        #region Peering
        foreach ($peer in $network.ConnectToVnets)
        {
            Write-ScreenInfo -Type Verbose -Message ('Adding peering from {0} to {1} to VNet template' -f $network.ResourceName, $peer)
            $template.Resources += @{
                apiVersion = $apiVersions['VirtualNetworkApi']
                dependsOn  = @(
                    "[resourceId('Microsoft.Network/virtualNetworks', '$($network.ResourceName)')]"
                    "[resourceId('Microsoft.Network/virtualNetworks', '$($peer)')]"
                )
                type       = "Microsoft.Network/virtualNetworks/virtualNetworkPeerings"
                name       = "$($network.ResourceName)/$($network.ResourceName)To$($peer)"
                properties = @{
                    allowVirtualNetworkAccess = $true
                    allowForwardedTraffic     = $false
                    allowGatewayTransit       = $false
                    useRemoteGateways         = $false
                    remoteVirtualNetwork      = @{
                        id = "[resourceId('Microsoft.Network/virtualNetworks', '$peer')]"
                    }
                }
            }
            $template.Resources += @{
                apiVersion = $apiVersions['VirtualNetworkApi']
                dependsOn  = @(
                    "[resourceId('Microsoft.Network/virtualNetworks', '$($network.ResourceName)')]"
                    "[resourceId('Microsoft.Network/virtualNetworks', '$($peer)')]"
                )
                type       = "Microsoft.Network/virtualNetworks/virtualNetworkPeerings"
                name       = "$($peer)/$($peer)To$($network.ResourceName)"
                properties = @{
                    allowVirtualNetworkAccess = $true
                    allowForwardedTraffic     = $false
                    allowGatewayTransit       = $false
                    useRemoteGateways         = $false
                    remoteVirtualNetwork      = @{
                        id = "[resourceId('Microsoft.Network/virtualNetworks', '$($network.ResourceName)')]"
                    }
                }
            }
        }

        foreach ($externalPeer in $network.PeeringVnetResourceIds) {
            $peerName = $externalPeer -split '/' | Select-Object -Last 1
            Write-ScreenInfo -Type Verbose -Message ('Adding peering from {0} to {1} to VNet template' -f $network.ResourceName, $peerName)
            $template.Resources += @{
                apiVersion = $apiVersions['VirtualNetworkApi']
                dependsOn  = @(
                    "[resourceId('Microsoft.Network/virtualNetworks', '$($network.ResourceName)')]"
                )
                type       = "Microsoft.Network/virtualNetworks/virtualNetworkPeerings"
                name       = "$($network.ResourceName)/$($network.ResourceName)To$($peerName)"
                properties = @{
                    allowVirtualNetworkAccess = $true
                    allowForwardedTraffic     = $false
                    allowGatewayTransit       = $false
                    useRemoteGateways         = $false
                    remoteVirtualNetwork      = @{
                        id = $externalPeer
                    }
                }
            }
        }
        #endregion

        #region Public Ip
        $dnsLabel = "[concat('al$vnetCount-', uniqueString(resourceGroup().id))]"

        if ($network.AzureDnsLabel)
        {
            $dnsLabel = $network.AzureDnsLabel
        }

        Write-ScreenInfo -Type Verbose -Message ('Adding public static IP with DNS label {0} to template' -f $dnsLabel)
        $template.resources +=
        @{
            apiVersion = $apiVersions['PublicIpApi']
            tags       = @{ 
                AutomatedLab = $Lab.Name
                CreationTime = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
                Vnet         = $network.ResourceName
            }
            type       = "Microsoft.Network/publicIPAddresses"
            name       = "lbip$vnetCount"
            location   = "[resourceGroup().location]"
            properties = @{
                publicIPAllocationMethod = "static"
                dnsSettings              = @{
                    domainNameLabel = $dnsLabel
                }
            }
            sku        = @{
                name = if ($Lab.AzureSettings.IsAzureStack) { 'Basic' } else { 'Standard' }
            }
        }
        #endregion

        #region Load balancer
        Write-ScreenInfo -Type Verbose -Message ('Adding load balancer to template')
        $loadbalancers[$network.ResourceName] = @{
            Name    = "lb$vnetCount"
            Backend = "$($vnetCount)lbbc"
        }
        $loadBalancer = @{
            type       = "Microsoft.Network/loadBalancers"
            tags       = @{ 
                AutomatedLab = $Lab.Name
                CreationTime = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
                Vnet         = $network.ResourceName
            }
            apiVersion = $apiVersions['LoadBalancerApi']
            name       = "lb$vnetCount"
            location   = "[resourceGroup().location]"
            sku        = @{
                name = if ($Lab.AzureSettings.IsAzureStack) { 'Basic' } else { 'Standard' }
            }
            dependsOn  = @(
                "[resourceId('Microsoft.Network/publicIPAddresses', 'lbip$vnetCount')]"
            )
            properties = @{
                frontendIPConfigurations = @(
                    @{
                        name       = "$($vnetCount)lbfc"
                        properties = @{
                            publicIPAddress = @{
                                id = "[resourceId('Microsoft.Network/publicIPAddresses', 'lbip$vnetCount')]"
                            }
                        }
                    }
                )
                backendAddressPools      = @(
                    @{
                        name = "$($vnetCount)lbbc"
                    }
                )
            }
        }

        if (-not $Lab.AzureSettings.IsAzureStack)
        {
            $loadbalancer.properties.outboundRules = @(
                @{
                    name       = "InternetAccess"
                    properties = @{
                        allocatedOutboundPorts   = 0 # In order to use automatic allocation
                        frontendIPConfigurations = @(
                            @{
                                id = "[resourceId('Microsoft.Network/loadBalancers/frontendIPConfigurations', 'lb$vnetCount', '$($vnetCount)lbfc')]"
                            }
                        )
                        backendAddressPool       = @{
                            id = "[concat(resourceId('Microsoft.Network/loadBalancers', 'lb$vnetCount'), '/backendAddressPools/$($vnetCount)lbbc')]"
                        }
                        protocol                 = "All"
                        enableTcpReset           = $true
                        idleTimeoutInMinutes     = 4
                    }
                }
            )
        }

        $rules = foreach ($machine in ($Lab.Machines | Where-Object -FilterScript { $_.Network -EQ $network.Name -and -not $_.SkipDeployment }))
        {
            Write-ScreenInfo -Type Verbose -Message ('Adding inbound NAT rules for {0}: {1}:3389, {2}:5985, {3}:5986, {4}:22' -f $machine, $machine.LoadBalancerRdpPort, $machine.LoadBalancerWinRmHttpPort, $machine.LoadBalancerWinrmHttpsPort, $machine.LoadBalancerSshPort)
            @{
                name       = "$($machine.ResourceName.ToLower())rdpin"
                properties = @{
                    frontendIPConfiguration = @{
                        id = "[resourceId('Microsoft.Network/loadBalancers/frontendIPConfigurations', 'lb$vnetCount', '$($vnetCount)lbfc')]"
                    }
                    frontendPort            = $machine.LoadBalancerRdpPort
                    backendPort             = 3389
                    enableFloatingIP        = $false
                    protocol                = "Tcp"
                }
            }
            @{
                name       = "$($machine.ResourceName.ToLower())winrmin"
                properties = @{
                    frontendIPConfiguration = @{
                        id = "[resourceId('Microsoft.Network/loadBalancers/frontendIPConfigurations', 'lb$vnetCount', '$($vnetCount)lbfc')]"
                    }
                    frontendPort            = $machine.LoadBalancerWinRmHttpPort
                    backendPort             = 5985
                    enableFloatingIP        = $false
                    protocol                = "Tcp"
                }
            }
            @{
                name       = "$($machine.ResourceName.ToLower())winrmhttpsin"
                properties = @{
                    frontendIPConfiguration = @{
                        id = "[resourceId('Microsoft.Network/loadBalancers/frontendIPConfigurations', 'lb$vnetCount', '$($vnetCount)lbfc')]"
                    }
                    frontendPort            = $machine.LoadBalancerWinrmHttpsPort
                    backendPort             = 5986
                    enableFloatingIP        = $false
                    protocol                = "Tcp"
                }
            }
            @{
                name       = "$($machine.ResourceName.ToLower())sshin"
                properties = @{
                    frontendIPConfiguration = @{
                        id = "[resourceId('Microsoft.Network/loadBalancers/frontendIPConfigurations', 'lb$vnetCount', '$($vnetCount)lbfc')]"
                    }
                    frontendPort            = $machine.LoadBalancerSshPort
                    backendPort             = 22
                    enableFloatingIP        = $false
                    protocol                = "Tcp"
                }
            }
        }

        $loadBalancer.properties.inboundNatRules = $rules
        $template.resources += $loadBalancer
        #endregion

        $vnetCount++
    }

    #region Disks
    foreach ($disk in $Lab.Disks)
    {
        if (-not $disk) { continue } # Due to an issue with the disk collection being enumerated even if it is empty
        Write-ScreenInfo -Type Verbose -Message ('Creating managed data disk {0} ({1} GB)' -f $disk.Name, $disk.DiskSize)
        $vm = $lab.Machines | Where-Object { $_.Disks.Name -contains $disk.Name }
        $template.resources += @{
            type       = "Microsoft.Compute/disks"
            tags       = @{ 
                AutomatedLab = $Lab.Name
                CreationTime = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            }
            apiVersion = $apiVersions['DiskApi']
            name       = $disk.Name
            location   = "[resourceGroup().location]"
            sku        = @{
                name = if ($vm.AzureProperties.StorageSku)
                {
                    $vm.AzureProperties['StorageSku']
                }
                else
                {
                    "Standard_LRS"
                }
            }
            properties = @{
                creationData = @{
                    createOption = "Empty"
                }
                diskSizeGB   = $disk.DiskSize
            }
        }
    }
    #endregion

    foreach ($machine in $Lab.Machines.Where({ -not $_.SkipDeployment }))
    {
        $niccount = 0
        foreach ($nic in $machine.NetworkAdapters)
        {
            Write-ScreenInfo -Type Verbose -Message ('Creating NIC {0}' -f $nic.InterfaceName)
            $subnetName = 'default'

            foreach ($subnetConfig in $nic.VirtualSwitch.Subnets)
            {
                if ($subnetConfig.Name -eq 'AzureBastionSubnet') { continue }

                $usable = Get-NetworkRange -IPAddress $subnetConfig.AddressSpace.IpAddress.AddressAsString -SubnetMask $subnetConfig.AddressSpace.Cidr
                if ($nic.Ipv4Address[0].IpAddress.AddressAsString -in $usable)
                {
                    $subnetName = $subnetConfig.Name
                }
            }

            $machineInboundRules = @(
                @{
                    id = "[concat(resourceId('Microsoft.Network/loadBalancers', '$($loadBalancers[$nic.VirtualSwitch.ResourceName].Name)'),'/inboundNatRules/$($machine.ResourceName.ToLower())rdpin')]"
                }
                @{
                    id = "[concat(resourceId('Microsoft.Network/loadBalancers', '$($loadBalancers[$nic.VirtualSwitch.ResourceName].Name)'),'/inboundNatRules/$($machine.ResourceName.ToLower())winrmin')]"
                }
                @{
                    id = "[concat(resourceId('Microsoft.Network/loadBalancers', '$($loadBalancers[$nic.VirtualSwitch.ResourceName].Name)'),'/inboundNatRules/$($machine.ResourceName.ToLower())winrmhttpsin')]"
                }
                @{
                    id = "[concat(resourceId('Microsoft.Network/loadBalancers', '$($loadBalancers[$nic.VirtualSwitch.ResourceName].Name)'),'/inboundNatRules/$($machine.ResourceName.ToLower())sshin')]"
                }
            )
             
            $nicTemplate = @{
                dependsOn  = @(
                    "[resourceId('Microsoft.Network/virtualNetworks', '$($nic.VirtualSwitch.ResourceName)')]"
                    "[resourceId('Microsoft.Network/loadBalancers', '$($loadBalancers[$nic.VirtualSwitch.ResourceName].Name)')]"
                )
                properties = @{
                    enableAcceleratedNetworking = $false
                    ipConfigurations            = @(
                        @{
                            properties = @{
                                subnet                    = @{
                                    id = "[resourceId('Microsoft.Network/virtualNetworks/subnets', '$($nic.VirtualSwitch.ResourceName)', '$subnetName')]"
                                }
                                primary                   = $true
                                privateIPAllocationMethod = "Static"
                                privateIPAddress          = $nic.Ipv4Address[0].IpAddress.AddressAsString
                                privateIPAddressVersion   = "IPv4"
                            }
                            name       = "ipconfig1"
                        }
                    )
                    enableIPForwarding          = $false
                }
                name       = "$($machine.ResourceName)nic$($niccount)"
                apiVersion = $apiVersions['NicApi']
                type       = "Microsoft.Network/networkInterfaces"
                location   = "[resourceGroup().location]"
                tags       = @{ 
                    AutomatedLab = $Lab.Name
                    CreationTime = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
                }
            }

            # Add NAT only to first nic
            if ($niccount -eq 0)
            {
                $nicTemplate.properties.ipConfigurations[0].properties.loadBalancerInboundNatRules = $machineInboundRules
                $nicTemplate.properties.ipConfigurations[0].properties.loadBalancerBackendAddressPools = @(
                    @{
                        id = "[concat(resourceId('Microsoft.Network/loadBalancers', '$($loadBalancers[$nic.VirtualSwitch.ResourceName].Name)'), '/backendAddressPools/$($loadBalancers[$nic.VirtualSwitch.ResourceName].Backend)')]"
                    }
                )
            }

            if (($Lab.VirtualNetworks | Where-Object ResourceName -eq $nic.VirtualSwitch).DnsServers)
            {
                $nicTemplate.properties.dnsSettings = @{
                    dnsServers = [string[]](($Lab.VirtualNetworks | Where-Object ResourceName -eq $nic.VirtualSwitch).DnsServers.AddressAsString)
                }
            }
            if ($nic.Ipv4DnsServers)
            {
                $nicTemplate.properties.dnsSettings = @{
                    dnsServers = [string[]]($nic.Ipv4DnsServers.AddressAsString)
                }
            }
            $template.resources += $nicTemplate
            $niccount++
        }

        Write-ScreenInfo -Type Verbose -Message ('Adding machine template')
        $vmSize = Get-LWAzureVmSize -Machine $Machine
        $imageRef = Get-LWAzureSku -Machine $machine

        if (-not $vmSize)
        {
            throw "No valid VM size found for '$Machine'. For a list of available role sizes, use the command 'Get-LabAzureAvailableRoleSize -LocationName $($lab.AzureSettings.DefaultLocation.Location)'"
        }

        Write-ScreenInfo -Type Verbose -Message "Adding $Machine with size $vmSize, publisher $($imageRef.publisher), offer $($imageRef.offer), sku $($imageRef.sku)!"

        $machTemplate = @{
            name       = $machine.ResourceName
            tags       = @{ 
                AutomatedLab = $Lab.Name
                CreationTime = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            }
            dependsOn  = @()
            properties = @{
                storageProfile  = @{
                    osDisk         = @{
                        createOption = "FromImage"
                        osType       = $Machine.OperatingSystemType.ToString()
                        caching      = "ReadWrite"
                        managedDisk  = @{
                            storageAccountType = if ($Machine.AzureProperties.ContainsKey('StorageSku') -and $Machine.AzureProperties['StorageSku'] -notmatch 'ultra')
                            {
                                $Machine.AzureProperties['StorageSku']
                            }
                            elseif ($Machine.AzureProperties.ContainsKey('StorageSku') -and $Machine.AzureProperties['StorageSku'] -match 'ultra')
                            {
                                Write-ScreenInfo -Type Warning -Message "Ultra_SSD SKU selected, defaulting to Premium_LRS for OS disk."
                                'Premium_LRS'
                            }
                            else
                            {
                                'Standard_LRS'
                            }
                        }
                    }                    
                    imageReference = $imageRef
                    dataDisks      = @()
                }
                networkProfile  = @{
                    networkInterfaces = @()
                }
                osProfile       = @{
                    adminPassword            = $machine.GetLocalCredential($true).GetNetworkCredential().Password
                    computerName             = $machine.Name
                    allowExtensionOperations = $true
                    adminUsername            = if ($machine.OperatingSystemType -eq 'Linux') { 'automatedlab' } else { ($machine.GetLocalCredential($true).UserName -split '\\')[-1] }
                }
                hardwareProfile = @{
                    vmSize = $vmSize.Name
                }
            }
            type       = "Microsoft.Compute/virtualMachines"
            apiVersion = $apiVersions['VmApi']
            location   = "[resourceGroup().location]"
        }

        if ($machine.OperatingSystem.OperatingSystemName -like 'Kali*')
        {
            # This is a marketplace offer, so we have to do redundant stuff for no good reason
            $machTemplate.plan = @{
                name      = $imageRef.sku # Otherwise known as sku
                product   = $imageRef.offer # Otherwise known as offer
                publisher = $imageRef.publisher # publisher
            }
        }

        if ($machine.OperatingSystemType -eq 'Windows')
        {
            $machTemplate.properties.osProfile.windowsConfiguration = @{
                enableAutomaticUpdates = $true
                provisionVMAgent       = $true
                winRM                  = @{
                    listeners = @(
                        @{
                            protocol = "Http"
                        }
                    )
                }
            }
        }

        if ($machine.OperatingSystemType -eq 'Linux')
        {
            if ($machine.SshPublicKey)
            {
                $machTemplate.properties.osProfile.linuxConfiguration = @{
                    disablePasswordAuthentication = $true
                    enableVMAgentPlatformUpdates  = $true
                    provisionVMAgent              = $true
                    ssh                           = @{
                        publicKeys = [hashtable[]]@(@{
                                keyData = $machine.SshPublicKey
                                path    = "/home/automatedlab/.ssh/authorized_keys"
                            }
                        )
                    }
                }
            }
        }
        
        if ($machine.AzureProperties['EnableSecureBoot'] -and -not $lab.AzureSettings.IsAzureStack) # Available only in public regions
        {            
            $machTemplate.properties.securityProfile = @{
                securityType = 'TrustedLaunch'
                uefiSettings = @{
                    secureBootEnabled = $true
                    vTpmEnabled       = $Machine.AzureProperties['EnableTpm'] -match '1|true|yes'
                }
            }
        }

        $luncount = 0
        foreach ($disk in $machine.Disks)
        {
            if (-not $disk) { continue } # Due to an issue with the disk collection being enumerated even if it is empty
            Write-ScreenInfo -Type Verbose -Message ('Adding disk {0} to machine template' -f $disk.Name)
            $machTemplate.properties.storageProfile.dataDisks += @{
                lun          = $luncount
                name         = $disk.Name
                createOption = "attach"
                managedDisk  = @{
                    id = "[resourceId('Microsoft.Compute/disks/', '$($disk.Name)')]"
                }
            }
            $luncount++
        }

        $niccount = 0
        foreach ($nic in $machine.NetworkAdapters)
        {
            Write-ScreenInfo -Type Verbose -Message ('Adding NIC {0} to template' -f $nic.InterfaceName)
            $machtemplate.dependsOn += "[resourceId('Microsoft.Network/networkInterfaces', '$($machine.ResourceName)nic$($niccount)')]"
            $machTemplate.properties.networkProfile.networkInterfaces += @{
                id         = "[resourceId('Microsoft.Network/networkInterfaces', '$($machine.ResourceName)nic$($niccount)')]"
                properties = @{
                    primary = $niccount -eq 0
                }
            }
            $niccount++
        }
        
        $template.resources += $machTemplate
    }

    $rgDeplParam = @{
        TemplateObject    = $template
        ResourceGroupName = $lab.AzureSettings.DefaultResourceGroup.ResourceGroupName
        Force             = $true
    }

    $templatePath = Join-Path -Path (Get-LabConfigurationItem -Name LabAppDataRoot) -ChildPath "Labs/$($Lab.Name)/armtemplate.json"
    ConvertTo-Json -InputObject $template -Depth 42 | Set-Content -Path $templatePath

    Write-ScreenInfo -Message "Deploying new resource group with template $templatePath"
    # Without wait - unable to catch exception
    if ($Wait.IsPresent)
    {
        $azureRetryCount = Get-LabConfigurationItem -Name AzureRetryCount
        $count = 1
        while ($count -le $azureRetryCount -and -not $deployment)
        {
            try
            {
                $deployment = New-AzResourceGroupDeployment @rgDeplParam -ErrorAction Stop
            }
            catch
            {
                if ($_.Exception.Message -match 'Code:NoRegisteredProviderFound')
                {
                    $count++
                }
                else
                {
                    Write-Error -Message 'Unrecoverable error during resource group deployment' -Exception $_.Exception
                    return
                }
            }
        }
        if ($count -gt $azureRetryCount)
        {
            Write-Error -Message 'Unrecoverable error during resource group deployment'
            return
        }
    }
    else
    {
        $deployment = New-AzResourceGroupDeployment @rgDeplParam -AsJob # Splatting AsJob did not work
    }
    

    if ($PassThru.IsPresent)
    {
        $deployment
    }

    Write-LogFunctionExit
}


function Remove-LWAzureRecoveryServicesVault
{
    [CmdletBinding()]
    param
    (
        [int]
        $RetryCount = 0
    )

    $lab = Get-Lab -ErrorAction SilentlyContinue
    if (-not $lab) { return }

    $rsVault = Get-AzResource -ResourceGroupName $lab.AzureSettings.DefaultResourceGroup.ResourceGroupName -ResourceType Microsoft.RecoveryServices/vaults -ErrorAction SilentlyContinue
    if (-not $rsVault) { return }

    if (-not (Get-Module -ListAvailable -Name Az.RecoveryServices | Where-Object Version -ge '5.3.0'))
    {
        try
        {
            Install-Module -Force -Name Az.RecoveryServices -Repository PSGallery -MinimumVersion 5.3.0 -ErrorAction Stop
        }
        catch
        {
            Write-ScreenInfo -Type Error -Message "Unable to install Az.RecoveryServices, 5.3.0+. Please delete your RecoveryServices Vault $($rsVault.Id) yourself."
            return
        }
    }

    Write-LogFunctionEntry
    Write-ScreenInfo -Message "Removing recovery services vault $($rsVault.Id) in $($rsVault.ResourceGroupName) so that the resource group can be deleted properly. This takes a while."
    $vaultToDelete = Get-AzRecoveryServicesVault -Name $rsVault.ResourceName -ResourceGroupName $rsVault.ResourceGroupName
    $null = Set-AzRecoveryServicesAsrVaultContext -Vault $vaultToDelete

    $null = Set-AzRecoveryServicesVaultProperty -Vault $vaultToDelete.ID -SoftDeleteFeatureState Disable #disable soft delete
    $containerSoftDelete = Get-AzRecoveryServicesBackupItem -BackupManagementType AzureVM -WorkloadType AzureVM -VaultId $vaultToDelete.ID | Where-Object { $_.DeleteState -eq "ToBeDeleted" } #fetch backup items in soft delete state
    foreach ($softitem in $containerSoftDelete)
    {
        $null = Undo-AzRecoveryServicesBackupItemDeletion -Item $softitem -VaultId $vaultToDelete.ID -Force #undelete items in soft delete state
    }
    
    if ((Get-Command Set-AzRecoveryServicesVaultProperty).Parameters.ContainsKey('DisableHybridBackupSecurityFeature'))
    {
        $null = Set-AzRecoveryServicesVaultProperty -VaultId $vaultToDelete.ID -DisableHybridBackupSecurityFeature $true
    }

    #Fetch all protected items and servers
    # Collection of try/catches since some enum values might be invalid
    $backupItemsVM = try { Get-AzRecoveryServicesBackupItem -BackupManagementType AzureVM -WorkloadType AzureVM -VaultId $vaultToDelete.ID -ErrorAction Stop } catch {}
    $backupItemsSQL = try { Get-AzRecoveryServicesBackupItem -BackupManagementType AzureWorkload -WorkloadType MSSQL -VaultId $vaultToDelete.ID -ErrorAction Stop } catch {}
    $backupItemsAFS = try { Get-AzRecoveryServicesBackupItem -BackupManagementType AzureStorage -WorkloadType AzureFiles -VaultId $vaultToDelete.ID -ErrorAction Stop } catch {}
    $backupItemsSAP = try { Get-AzRecoveryServicesBackupItem -BackupManagementType AzureWorkload -WorkloadType SAPHanaDatabase -VaultId $vaultToDelete.ID -ErrorAction Stop } catch {}
    $backupContainersSQL = try { Get-AzRecoveryServicesBackupContainer -ContainerType AzureVMAppContainer -Status Registered -VaultId $vaultToDelete.ID -ErrorAction Stop | Where-Object { $_.ExtendedInfo.WorkloadType -eq "SQL" } } catch {}
    $protectableItemsSQL = try { Get-AzRecoveryServicesBackupProtectableItem -WorkloadType MSSQL -VaultId $vaultToDelete.ID -ErrorAction Stop | Where-Object { $_.IsAutoProtected -eq $true } } catch {}
    $backupContainersSAP = try { Get-AzRecoveryServicesBackupContainer -ContainerType AzureVMAppContainer -Status Registered -VaultId $vaultToDelete.ID -ErrorAction Stop | Where-Object { $_.ExtendedInfo.WorkloadType -eq "SAPHana" } } catch {}
    $StorageAccounts = try { Get-AzRecoveryServicesBackupContainer -ContainerType AzureStorage -Status Registered -VaultId $vaultToDelete.ID -ErrorAction Stop } catch {}
    $backupServersMARS = try { Get-AzRecoveryServicesBackupContainer -ContainerType "Windows" -BackupManagementType MAB -VaultId $vaultToDelete.ID -ErrorAction Stop } catch {}
    $backupServersMABS = try { Get-AzRecoveryServicesBackupManagementServer -VaultId $vaultToDelete.ID -ErrorAction Stop | Where-Object { $_.BackupManagementType -eq "AzureBackupServer" } } catch {}
    $backupServersDPM = try { Get-AzRecoveryServicesBackupManagementServer -VaultId $vaultToDelete.ID -ErrorAction Stop | Where-Object { $_.BackupManagementType -eq "SCDPM" } } catch {}
    $pvtendpoints = try { Get-AzPrivateEndpointConnection -PrivateLinkResourceId $vaultToDelete.ID -ErrorAction Stop } catch {}

    $pool = New-RunspacePool -Variable (Get-Variable vaultToDelete) -ThrottleLimit 20
    $jobs = [system.Collections.ArrayList]::new()

    foreach ($item in $backupItemsVM)
    {
        $null = $jobs.Add((Start-RunspaceJob -ScriptBlock { param ($item) Disable-AzRecoveryServicesBackupProtection -Item $item -VaultId $vaultToDelete.ID -RemoveRecoveryPoints -Force } -RunspacePool $pool -Argument $item))
    }

    foreach ($item in $backupItemsSQL)
    {
        $null = $jobs.Add((Start-RunspaceJob -ScriptBlock { param ($item) Disable-AzRecoveryServicesBackupProtection -Item $item -VaultId $vaultToDelete.ID -RemoveRecoveryPoints -Force } -RunspacePool $pool -Argument $item))
    }

    foreach ($item in $protectableItems)
    {
        $null = $jobs.Add((Start-RunspaceJob -ScriptBlock { param ($item) Disable-AzRecoveryServicesBackupAutoProtection -BackupManagementType AzureWorkload -WorkloadType MSSQL -InputItem $item -VaultId $vaultToDelete.ID } -RunspacePool $pool -Argument $item))
    }

    foreach ($item in $backupContainersSQL)
    {
        $null = $jobs.Add((Start-RunspaceJob -ScriptBlock { param ($item) Unregister-AzRecoveryServicesBackupContainer -Container $item -Force -VaultId $vaultToDelete.ID } -RunspacePool $pool -Argument $item))
    }

    foreach ($item in $backupItemsSAP)
    {
        $null = $jobs.Add((Start-RunspaceJob -ScriptBlock { param ($item) Disable-AzRecoveryServicesBackupProtection -Item $item -VaultId $vaultToDelete.ID -RemoveRecoveryPoints -Force } -RunspacePool $pool -Argument $item))
    }

    foreach ($item in $backupContainersSAP)
    {
        $null = $jobs.Add((Start-RunspaceJob -ScriptBlock { param ($item) Unregister-AzRecoveryServicesBackupContainer -Container $item -Force -VaultId $vaultToDelete.ID } -RunspacePool $pool -Argument $item))
    }

    foreach ($item in $backupItemsAFS)
    {
        $null = $jobs.Add((Start-RunspaceJob -ScriptBlock { param ($item) Disable-AzRecoveryServicesBackupProtection -Item $item -VaultId $vaultToDelete.ID -RemoveRecoveryPoints -Force } -RunspacePool $pool -Argument $item))
    }

    foreach ($item in $StorageAccounts)
    {
        $null = $jobs.Add((Start-RunspaceJob -ScriptBlock { param ($item) Unregister-AzRecoveryServicesBackupContainer -container $item -Force -VaultId $vaultToDelete.ID } -RunspacePool $pool -Argument $item))
    }

    foreach ($item in $backupServersMARS)
    {
        $null = $jobs.Add((Start-RunspaceJob -ScriptBlock { param ($item) Unregister-AzRecoveryServicesBackupContainer -Container $item -Force -VaultId $vaultToDelete.ID } -RunspacePool $pool -Argument $item))
    }

    foreach ($item in $backupServersMABS)
    {
        $null = $jobs.Add((Start-RunspaceJob -ScriptBlock { param ($item) Unregister-AzRecoveryServicesBackupManagementServer -AzureRmBackupManagementServer $item -VaultId $vaultToDelete.ID } -RunspacePool $pool -Argument $item))
    }

    foreach ($item in $backupServersDPM)
    {
        $null = $jobs.Add((Start-RunspaceJob -ScriptBlock { param ($item) Unregister-AzRecoveryServicesBackupManagementServer -AzureRmBackupManagementServer $item -VaultId $vaultToDelete.ID } -RunspacePool $pool -Argument $item))
    }

    $null = Wait-RunspaceJob -RunspaceJob $jobs
    Remove-RunspacePool -RunspacePool $pool

    #Deletion of ASR Items
    $fabricObjects = Get-AzRecoveryServicesAsrFabric
    # First DisableDR all VMs.
    foreach ($fabricObject in $fabricObjects)
    {
        $containerObjects = Get-AzRecoveryServicesAsrProtectionContainer -Fabric $fabricObject -ErrorAction SilentlyContinue
        foreach ($containerObject in $containerObjects)
        {
            $protectedItems = Get-AzRecoveryServicesAsrReplicationProtectedItem -ProtectionContainer $containerObject -ErrorAction SilentlyContinue
            # DisableDR all protected items
            foreach ($protectedItem in $protectedItems)
            {
                $null = Remove-AzRecoveryServicesAsrReplicationProtectedItem -InputObject $protectedItem -Force
            }

            $containerMappings = Get-AzRecoveryServicesAsrProtectionContainerMapping -ProtectionContainer $containerObject
            # Remove all Container Mappings
            foreach ($containerMapping in $containerMappings)
            {
                $null = Remove-AzRecoveryServicesAsrProtectionContainerMapping -ProtectionContainerMapping $containerMapping -Force
            }
        }
        $networkObjects = Get-AzRecoveryServicesAsrNetwork -Fabric $fabricObject
        foreach ($networkObject in $networkObjects)
        {
            #Get the PrimaryNetwork
            $PrimaryNetwork = Get-AzRecoveryServicesAsrNetwork -Fabric $fabricObject -FriendlyName $networkObject
            $NetworkMappings = Get-AzRecoveryServicesAsrNetworkMapping -Network $PrimaryNetwork
            foreach ($networkMappingObject in $NetworkMappings)
            {
                #Get the Neetwork Mappings
                $NetworkMapping = Get-AzRecoveryServicesAsrNetworkMapping -Name $networkMappingObject.Name -Network $PrimaryNetwork
                $null = Remove-AzRecoveryServicesAsrNetworkMapping -InputObject $NetworkMapping
            }
        }
        # Remove Fabric
        $null = Remove-AzRecoveryServicesAsrFabric -InputObject $fabricObject -Force
    }

    foreach ($item in $pvtendpoints)
    {
        $penamesplit = $item.Name.Split(".")
        $pename = $penamesplit[0]
        $null = Remove-AzPrivateEndpointConnection -ResourceId $item.PrivateEndpoint.Id -Force #remove private endpoint connections
        $null = Remove-AzPrivateEndpoint -Name $pename -ResourceGroupName $lab.AzureSettings.DefaultResourceGroup.ResourceGroupName -Force #remove private endpoints
    }

    try
    {
        $null = Remove-AzRecoveryServicesVault -Vault $vaultToDelete -Confirm:$false -ErrorAction Stop
    }
    catch
    {
        if ($RetryCount -le 2)
        {
            Remove-LWAzureRecoveryServicesVault -RetryCount ($RetryCount + 1)
        }
    }
    Write-LogFunctionExit
}


function Remove-LWAzureVM
{
    Param (
        [Parameter(Mandatory)]
        [string]$Name,

        [switch]$AsJob,

        [switch]$PassThru
    )

    Test-LabHostConnected -Throw -Quiet

    Write-LogFunctionEntry

    $azureRetryCount = Get-LabConfigurationItem -Name AzureRetryCount

    $Lab = Get-Lab
    $vm = Get-AzVM -ResourceGroupName $Lab.AzureSettings.DefaultResourceGroup.ResourceGroupName -Name $Name -ErrorAction SilentlyContinue
    $null = $vm | Remove-AzVM -Force
    foreach ($loadBalancer in (Get-AzLoadBalancer -ResourceGroupName $Lab.AzureSettings.DefaultResourceGroup.ResourceGroupName))
    {
        $rules = $loadBalancer | Get-AzLoadBalancerInboundNatRuleConfig | Where-Object Name -like "$($Name)*"
        foreach ($rule in $rules)
        {
            $null = Remove-AzLoadBalancerInboundNatRuleConfig -LoadBalancer $loadBalancer -Name $rule.Name -Confirm:$false
        }
    }

    $vmResources = Get-AzResource -ResourceGroupName $Lab.AzureSettings.DefaultResourceGroup.ResourceGroupName -Name "$($name)*"
    $jobs = $vmResources | Remove-AzResource -AsJob -Force -Confirm:$false

    if (-not $AsJob.IsPresent)
    {
        $null = $jobs | Wait-Job
    }

    if ($PassThru.IsPresent)
    {
        $jobs
    }

    Write-LogFunctionExit
}


function Remove-LWAzureVmSnapshot
{
    [Cmdletbinding()]
    Param
    (
        [Parameter(Mandatory, ParameterSetName = 'BySnapshotName')]
        [Parameter(Mandatory, ParameterSetName = 'AllSnapshots')]
        [string[]]$ComputerName,

        [Parameter(Mandatory, ParameterSetName = 'BySnapshotName')]
        [string]$SnapshotName,

        [Parameter(ParameterSetName = 'AllSnapshots')]
        [switch]$All
    )

    Test-LabHostConnected -Throw -Quiet

    Write-LogFunctionEntry

    $lab = Get-Lab

    $snapshots = Get-AzSnapshot -ResourceGroupName $lab.AzureSettings.DefaultResourceGroup.ResourceGroupName -ErrorAction SilentlyContinue

    if ($PSCmdlet.ParameterSetName -eq 'BySnapshotName')
    {
        $snapshotsToRemove = $ComputerName.Foreach( { '{0}_{1}' -f $_, $SnapshotName })
        $snapshots = $snapshots | Where-Object -Property Name -in $snapshotsToRemove
    }

    $null = $snapshots | Remove-AzSnapshot -Force -Confirm:$false

    Write-LogFunctionExit
}


function Restore-LWAzureVmSnapshot
{
    [Cmdletbinding()]
    Param
    (
        [Parameter(Mandatory)]
        [string[]]$ComputerName,

        [Parameter(Mandatory)]
        [string]$SnapshotName
    )

    Test-LabHostConnected -Throw -Quiet

    Write-LogFunctionEntry

    $lab = Get-Lab
    $resourceGroupName = $lab.AzureSettings.DefaultResourceGroup.ResourceGroupName

    $runningMachines = Get-LabVM -IsRunning -ComputerName $ComputerName -IncludeLinux
    if ($runningMachines)
    {
        Stop-LWAzureVM -ComputerName $runningMachines -StayProvisioned $true
        Wait-LabVMShutdown -ComputerName $runningMachines
    }

    $vms = Get-AzVM -ResourceGroupName $resourceGroupName | Where-Object Name -In $ComputerName
    $machineStatus = @{}
    $ComputerName.ForEach( { $machineStatus[$_] = @{ Stage1 = $null; Stage2 = $null; Stage3 = $null } })

    foreach ($machine in $ComputerName)
    {
        $vm = $vms | Where-Object Name -eq $machine
        $vmSnapshotName = '{0}_{1}' -f $machine, $SnapshotName
        if (-not $vm)
        {
            Write-ScreenInfo -Message "$machine could not be found in $($resourceGroupName). Skipping snapshot." -type Warning
            continue
        }

        $snapshot = Get-AzSnapshot -SnapshotName $vmSnapshotName -ResourceGroupName $resourceGroupName -ErrorAction SilentlyContinue
        if (-not $snapshot)
        {
            Write-ScreenInfo -Message "No snapshot named $vmSnapshotName found for $machine. Skipping restore." -Type Warning
            continue
        }

        $osDiskName = $vm.StorageProfile.OsDisk.name
        $oldOsDisk = Get-AzDisk -Name $osDiskName -ResourceGroupName $resourceGroupName
        $disksToRemove += $oldOsDisk.Name
        $storageType = $oldOsDisk.sku.name
        $diskconf = New-AzDiskConfig -AccountType $storagetype -Location $oldOsdisk.Location -SourceResourceId $snapshot.Id -CreateOption Copy

        $machineStatus[$machine].Stage1 = @{
            VM      = $vm
            OldDisk = $oldOsDisk.Name
            Job     = New-AzDisk -Disk $diskconf -ResourceGroupName $resourceGroupName -DiskName "$($vm.Name)-$((New-Guid).ToString())" -AsJob
        }
    }

    if ($machineStatus.Values.Stage1.Job)
    {
        $null = $machineStatus.Values.Stage1.Job | Wait-Job
    }

    $failedStage1 = $($machineStatus.GetEnumerator() | Where-Object -FilterScript { $_.Value.Stage1.Job.State -eq 'Failed' }).Name
    if ($failedStage1) { Write-ScreenInfo -Type Error -Message "The following machines failed to create a new disk from the snapshot: $($failedStage1 -join ',')" }

    $ComputerName = $($machineStatus.GetEnumerator() | Where-Object -FilterScript { $_.Value.Stage1.Job.State -eq 'Completed' }).Name

    foreach ($machine in $ComputerName)
    {
        $vm = $vms | Where-Object Name -eq $machine
        $newDisk = $machineStatus[$machine].Stage1.Job | Receive-Job -Keep
        $null = Set-AzVMOSDisk -VM $vm -ManagedDiskId $newDisk.Id -Name $newDisk.Name
        $machineStatus[$machine].Stage2 = @{
            Job = Update-AzVM -ResourceGroupName $resourceGroupName -VM $vm -AsJob
        }
    }

    if ($machineStatus.Values.Stage2.Job)
    {
        $null = $machineStatus.Values.Stage2.Job | Wait-Job
    }

    $failedStage2 = $($machineStatus.GetEnumerator() | Where-Object -FilterScript { $_.Value.Stage2.Job.State -eq 'Failed' }).Name
    if ($failedStage2) { Write-ScreenInfo -Type Error -Message "The following machines failed to update with the new OS disk created from a snapshot: $($failedStage2 -join ',')" }

    $ComputerName = $($machineStatus.GetEnumerator() | Where-Object -FilterScript { $_.Value.Stage2.Job.State -eq 'Completed' }).Name

    foreach ($machine in $ComputerName)
    {
        $disk = $machineStatus[$machine].Stage1.OldDisk
        $machineStatus[$machine].Stage3 = @{
            Job = Remove-AzDisk -ResourceGroupName $resourceGroupName -DiskName $disk -Confirm:$false -Force -AsJob
        }
    }
    if ($machineStatus.Values.Stage3.Job)
    {
        $null = $machineStatus.Values.Stage3.Job | Wait-Job
    }

    $failedStage3 = $($machineStatus.GetEnumerator() | Where-Object -FilterScript { $_.Value.Stage3.Job.State -eq 'Failed' }).Name
    if ($failedStage3)
    {
        $failedDisks = $failedStage3.ForEach( { $machineStatus[$_].Stage1.OldDisk })
        Write-ScreenInfo -Type Warning -Message "The following machines failed to remove their old OS disk in a background job: $($failedStage3 -join ','). Trying to remove the disks again synchronously."

        foreach ($machine in $failedStage3)
        {
            $disk = $machineStatus[$machine].Stage1.OldDisk
            $null = Remove-AzDisk -ResourceGroupName $resourceGroupName -DiskName $disk -Confirm:$false -Force
        }
    }

    if ($runningMachines)
    {
        Start-LWAzureVM -ComputerName $runningMachines
        Wait-LabVM -ComputerName $runningMachines
    }

    if ($machineStatus.Values.Values.Job)
    {
        $machineStatus.Values.Values.Job | Remove-Job
    }

    Write-LogFunctionExit
}


function Start-LWAzureVM
{
    param (
        [Parameter(Mandatory = $true)]
        [string[]]$ComputerName,

        [int]$DelayBetweenComputers = 0,

        [int]$ProgressIndicator = 15,

        [switch]$NoNewLine
    )

    Test-LabHostConnected -Throw -Quiet

    Write-LogFunctionEntry

    $azureRetryCount = Get-LabConfigurationItem -Name AzureRetryCount
    $machines = Get-LabVm -ComputerName $ComputerName -IncludeLinux

    $azureVms = Get-LWAzureVm -ComputerName $ComputerName

    $stoppedAzureVms = $azureVms | Where-Object { $_.PowerState -ne 'VM running' -and $_.Name -in $machines.ResourceName }

    $lab = Get-Lab

    $machinesToJoin = @()

    if ($stoppedAzureVms)
    {
        $jobs = foreach ($name in $machines.ResourceName)
        {
            $vm = $azureVms | Where-Object Name -eq $name
            $vm | Start-AzVM -AsJob
        }

        Wait-LWLabJob -Job $jobs -NoDisplay -ProgressIndicator $ProgressIndicator
    }

    # Refresh status
    $azureVms = Get-LWAzureVm -ComputerName $ComputerName

    $azureVms = $azureVms | Where-Object { $_.Name -in $machines.ResourceName }

    foreach ($machine in $machines)
    {
        $vm = $azureVms | Where-Object Name -eq $machine.ResourceName

        if ($vm.PowerState -ne 'VM Running')
        {
            throw "Could not start machine '$machine'"
        }
        else
        {
            if ($machine.IsDomainJoined -and -not $machine.HasDomainJoined -and ($machine.Roles.Name -notcontains 'RootDC' -and $machine.Roles.Name -notcontains 'FirstChildDC' -and $machine.Roles.Name -notcontains 'DC'))
            {
                $machinesToJoin += $machine
            }
        }
    }

    if ($machinesToJoin)
    {
        Write-PSFMessage -Message "Waiting for machines '$($machinesToJoin -join ', ')' to come online"
        Wait-LabVM -ComputerName $machinesToJoin -ProgressIndicator $ProgressIndicator -NoNewLine:$NoNewLine

        Write-PSFMessage -Message 'Start joining the machines to the respective domains'
        Join-LabVMDomain -Machine $machinesToJoin
    }

    Write-LogFunctionExit
}


function Stop-LWAzureVM
{
    param (
        [Parameter(Mandatory)]
        [string[]]
        $ComputerName,

        [ValidateRange(0, 300)]
        [int]$ProgressIndicator = (Get-LabConfigurationItem -Name DefaultProgressIndicator),

        [switch]
        $NoNewLine,

        [switch]
        $ShutdownFromOperatingSystem,

        [bool]
        $StayProvisioned = $false
    )

    Test-LabHostConnected -Throw -Quiet

    Write-LogFunctionEntry

    $azureRetryCount = Get-LabConfigurationItem -Name AzureRetryCount

    if (-not $PSBoundParameters.ContainsKey('ProgressIndicator')) { $PSBoundParameters.Add('ProgressIndicator', $ProgressIndicator) } #enables progress indicator

    $lab = Get-Lab
    $machines = Get-LabVm -ComputerName $ComputerName -IncludeLinux
    $azureVms = Get-AzVM -ResourceGroupName (Get-LabAzureDefaultResourceGroup).ResourceGroupName

    $azureVms = $azureVms | Where-Object { $_.Name -in $machines.ResourceName }

    if ($ShutdownFromOperatingSystem)
    {
        $jobs = @()
        $linux, $windows = $machines.Where( { $_.OperatingSystemType -eq 'Linux' }, 'Split')

        $jobs += Invoke-LabCommand -ComputerName $windows -NoDisplay -AsJob -PassThru -ScriptBlock {
            Stop-Computer -Force -ErrorAction Stop
        }

        $jobs += Invoke-LabCommand -UseLocalCredential -ComputerName $linux -NoDisplay -AsJob -PassThru -ScriptBlock {
            #Sleep as background process so that job does not fail.
            [void] (Start-Job {
                    Start-Sleep -Seconds 5
                    shutdown -P now
                })
        }

        Wait-LWLabJob -Job $jobs -NoDisplay -ProgressIndicator $ProgressIndicator
        $failedJobs = $jobs | Where-Object { $_.State -eq 'Failed' }
        if ($failedJobs)
        {
            Write-ScreenInfo -Message "Could not stop Azure VM(s): '$($failedJobs.Location)'" -Type Error
        }
    }
    else
    {
        $jobs = foreach ($name in $machines.ResourceName)
        {
            $vm = $azureVms | Where-Object Name -eq $name
            $vm | Stop-AzVM -Force -StayProvisioned:$StayProvisioned -AsJob
        }

        Wait-LWLabJob -Job $jobs -NoDisplay -ProgressIndicator $ProgressIndicator
        $failedJobs = $jobs | Where-Object { $_.State -eq 'Failed' }
        if ($failedJobs)
        {
            $jobNames = ($failedJobs | ForEach-Object {
                    if ($_.Name.StartsWith("StopAzureVm_"))
                    {
                        ($_.Name -split "_")[1]
                    }
                    elseif ($_.Name -match "Long Running Operation for 'Stop-AzVM' on resource '(?<MachineName>[\w-]+)'")
                    {
                        $Matches.MachineName
                    }
                }) -join ", "

            Write-ScreenInfo -Message "Could not stop Azure VM(s): '$jobNames'" -Type Error
        }
    }

    Write-ProgressIndicatorEnd

    Write-LogFunctionExit
}


function Wait-LWAzureRestartVM
{
    param (
        [Parameter(Mandatory)]
        [string[]]$ComputerName,

        [switch]$DoNotUseCredSsp,

        [double]$TimeoutInMinutes = 15,

        [int]$ProgressIndicator,

        [switch]$NoNewLine,

        [Parameter(Mandatory)]
        [datetime]
        $MonitoringStartTime
    )

    Test-LabHostConnected -Throw -Quiet

    #required to suporess verbose messages, warnings and errors
    Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState

    Write-LogFunctionEntry

    $azureRetryCount = Get-LabConfigurationItem -Name AzureRetryCount

    $start = $MonitoringStartTime.ToUniversalTime()

    Write-PSFMessage -Message "Starting monitoring the servers at '$start'"

    $machines = Get-LabVM -ComputerName $ComputerName

    $cmd = {
        param (
            [datetime]$Start
        )

        $Start = $Start.ToLocalTime()

        (Get-CimInstance -ClassName Win32_OperatingSystem).LastBootupTime -ge $Start
    }

    $ProgressIndicatorTimer = (Get-Date)

    do
    {
        $machines = foreach ($machine in $machines)
        {
            if (((Get-Date) - $ProgressIndicatorTimer).TotalSeconds -ge $ProgressIndicator)
            {
                Write-ProgressIndicator
                $ProgressIndicatorTimer = (Get-Date)
            }

            $hasRestarted = Invoke-LabCommand -ComputerName $machine -ActivityName WaitForRestartEvent -ScriptBlock $cmd -ArgumentList $start.Ticks -UseLocalCredential -DoNotUseCredSsp:$DoNotUseCredSsp -PassThru -Verbose:$false -NoDisplay -ErrorAction SilentlyContinue -WarningAction SilentlyContinue

            if (-not $hasRestarted)
            {
                $events = Invoke-LabCommand -ComputerName $machine -ActivityName WaitForRestartEvent -ScriptBlock $cmd -ArgumentList $start.Ticks -DoNotUseCredSsp:$DoNotUseCredSsp -PassThru -Verbose:$false -NoDisplay -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
            }

            if ($hasRestarted)
            {
                Write-PSFMessage -Message "VM '$machine' has been restarted"
            }
            else
            {
                Start-Sleep -Seconds 10
                $machine
            }
        }
    }
    until ($machines.Count -eq 0 -or (Get-Date).ToUniversalTime().AddMinutes( - $TimeoutInMinutes) -gt $start)

    if (-not $NoNewLine)
    {
        Write-ProgressIndicatorEnd
    }

    if ((Get-Date).ToUniversalTime().AddMinutes( - $TimeoutInMinutes) -gt $start)
    {
        foreach ($machine in ($machines))
        {
            Write-Error -Message "Timeout while waiting for computers to restart. Computers '$machine' not restarted" -TargetObject $machine
        }
    }

    Write-PSFMessage -Message "Finished monitoring the servers at '$(Get-Date)'"

    Write-LogFunctionExit
}


function Get-LWHypervNetworkSwitchDescription
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseCompatibleCmdlets", "", Justification="Not relevant on Linux")]
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$NetworkSwitchName
    )

    Write-LogFunctionEntry

    if (-not (Get-Lab -ErrorAction SilentlyContinue))
    {
        return
    }
    
    $notePath = Join-Path -Path (Get-Lab).LabPath -ChildPath "Network_$NetworkSwitchName.xml"
    if (-not (Test-Path -Path $notePath))
    {
        Write-Error "The file '$notePath' did not exist. Cannot import metadata of network switch '$NetworkSwitchName'"
        return
    }
    
    $type = Get-Type -GenericType AutomatedLab.DictionaryXmlStore -T string, string

    $dictionary = New-Object $type
    try
    {
        $importMethodInfo = $type.GetMethod('Import', [System.Reflection.BindingFlags]::Public -bor [System.Reflection.BindingFlags]::Static)
        $dictionary = $importMethodInfo.Invoke($null, $notePath)
        $dictionary
    }
    catch
    {
        Write-ScreenInfo -Message "The metadata of the network switch '$ComputerName' could not be read as XML" -Type Warning
    }

    Write-LogFunctionExit
}


function New-LWHypervNetworkSwitch
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseCompatibleCmdlets", "", Justification="Not relevant on Linux")]
    param (
        [Parameter(Mandatory)]
        [AutomatedLab.VirtualNetwork[]]$VirtualNetwork,

        [switch]$PassThru
    )

    Write-LogFunctionEntry

    foreach ($network in $VirtualNetwork)
    {
        if (-not $network.ResourceName)
        {
            throw 'No name specified for virtual network to be created'
        }

        Write-ScreenInfo -Message "Creating Hyper-V virtual network '$($network.ResourceName)'" -TaskStart

        if (Get-VMSwitch -Name $network.ResourceName -ErrorAction SilentlyContinue)
        {
            Write-ScreenInfo -Message "The network switch '$($network.ResourceName)' already exists, no changes will be made to configuration" -Type Warning
            continue
        }

        if ((Get-NetIPAddress -AddressFamily IPv4) -contains $network.AddressSpace.FirstUsable)
        {
            Write-ScreenInfo -Message "The IP '$($network.AddressSpace.FirstUsable)' Address for network switch '$($network.ResourceName)' is already in use" -Type Error
            return
        }

        try
        {
            $switchCreation = Get-LabConfigurationItem -Name SwitchDeploymentInProgressPath
            while (Test-Path -Path $switchCreation)
            {
                Start-Sleep -Milliseconds 250
            }

            $null = New-Item -Path $switchCreation -ItemType File -Value (Get-Lab).Name
            if ($network.SwitchType -eq 'External')
            {
                $adapterMac = (Get-NetAdapter -Name $network.AdapterName).MacAddress
                $adapterCountWithSameMac = (Get-NetAdapter | Where-Object { $_.MacAddress -eq $adapterMac -and $_.DriverDescription -ne 'Microsoft Network Adapter Multiplexor Driver' } | Group-Object -Property MacAddress).Count
                if ($adapterCountWithSameMac -gt 1)
                {
                    if (Get-NetLbfoTeam -Name $network.AdapterName -ErrorAction SilentlyContinue)
                    {
                        Write-ScreenInfo -Message "Network Adapter ($($network.AdapterName)) is a teamed interface, ignoring duplicate MAC checking" -Type Warning
                    }
                    else
                    {
                        throw "The given network adapter ($($network.AdapterName)) for the external virtual switch ($($network.ResourceName)) is already part of a network bridge and cannot be used."
                    }
                }

                $switch = New-VMSwitch -NetAdapterName $network.AdapterName -Name $network.ResourceName -AllowManagementOS $network.EnableManagementAdapter -ErrorAction Stop
            }
            else
            {
                try
                {
                    $switch = New-VMSwitch -Name $network.ResourceName -SwitchType ([string]$network.SwitchType) -ErrorAction Stop
                }
                catch
                {
                    Start-Sleep -Seconds 2
                    $switch = New-VMSwitch -Name $network.ResourceName -SwitchType ([string]$network.SwitchType) -ErrorAction Stop
                }

                if ($network.UseNat) {
                    $null = New-NetNat -Name $network.ResourceName -InternalIPInterfaceAddressPrefix $network.AddressSpace
                }

                Set-LWHypervNetworkSwitchDescription -NetworkSwitchName $network.ResourceName -Hashtable @{
                    CreatedBy = '{0} ({1})' -f $PSCmdlet.MyInvocation.MyCommand.Module.Name, $PSCmdlet.MyInvocation.MyCommand.Module.Version
                    CreationTime = Get-Date
                    LabName = (Get-Lab).Name
                }
            }
        }
        finally
        {
            Remove-Item -Path $switchCreation -ErrorAction SilentlyContinue
        }

        Start-Sleep -Seconds 1

        if ($network.EnableManagementAdapter) {

            $config = Get-NetAdapter | Where-Object Name -Match "^vEthernet \($($network.ResourceName)\) ?(\d{1,2})?"
            if (-not $config)
            {
                throw "The network adapter for network switch '$network' could not be found. Cannot set up address hence will not be able to contact the machines"
            }

            if ($null -ne $network.ManagementAdapter.InterfaceName)
            {
                #A management adapter was defined, use its provided IP settings
                $adapterIpAddress = if ($network.ManagementAdapter.ipv4Address.IpAddress -eq $network.ManagementAdapter.ipv4Address.Network)
                {
                    $network.ManagementAdapter.ipv4Address.FirstUsable
                }
                else
                {
                    $network.ManagementAdapter.ipv4Address.IpAddress
                }

                $adapterCidr = if ($network.ManagementAdapter.ipv4Address.Cidr)
                {
                    $network.ManagementAdapter.ipv4Address.Cidr
                }
                else
                {
                    #default to a class C (255.255.255.0) CIDR if one wasnt supplied
                    24
                }

                #Assign the IP address to the interface, implementing a default gateway if one was supplied
                if ($network.ManagementAdapter.ipv4Gateway) {
                    $null = New-NetIPAddress -InterfaceAlias "vEthernet ($($network.ResourceName))" -IPAddress $adapterIpAddress.AddressAsString -AddressFamily IPv4 -PrefixLength $adapterCidr -DefaultGateway $network.ManagementAdapter.ipv4Gateway.AddressAsString
                }
                else
                {
                    $null = New-NetIPAddress -InterfaceAlias "vEthernet ($($network.ResourceName))" -IPAddress $adapterIpAddress.AddressAsString -AddressFamily IPv4 -PrefixLength $adapterCidr
                }

                if (-not $network.ManagementAdapter.AccessVLANID -eq 0) {
                    #VLANID has been specified for the vEthernet Adapter, so set it
                    Set-VMNetworkAdapterVlan -ManagementOS -VMNetworkAdapterName $network.ResourceName -Access -VlanId $network.ManagementAdapter.AccessVLANID
                }
            }
            else
            {
                #if no address space has been defined, the management adapter will just be left as a DHCP-enabled interface
                if ($null -ne $network.AddressSpace)
                {
                    #if the network address was defined, get the first usable IP for the network adapter
                    $adapterIpAddress = if ($network.AddressSpace.IpAddress -eq $network.AddressSpace.Network)
                    {
                        $network.AddressSpace.FirstUsable
                    }
                    else
                    {
                        $network.AddressSpace.IpAddress
                    }

                    while ($adapterIpAddress -in (Get-LabMachineDefinition).IpAddress.IpAddress)
                    {
                        $adapterIpAddress = $adapterIpAddress.Increment()
                    }

                    $null = $config | Set-NetIPInterface -Dhcp Disabled
                    $null = $config | Remove-NetIPAddress -Confirm:$false
                    $null = $config | New-NetIPAddress -IPAddress $adapterIpAddress.AddressAsString -AddressFamily IPv4 -PrefixLength $network.AddressSpace.Cidr
                }
                else
                {
                    Write-ScreenInfo -Message "Management Interface for switch '$($network.ResourceName)' on Network Adapter '$($network.AdapterName)' has no defined AddressSpace and will remain DHCP enabled, ensure this is desired behaviour." -Type Warning
                }
            }
        }
        Write-ScreenInfo -Message "Done" -TaskEnd

        if ($PassThru)
        {
            $switch
        }
    }

    Write-LogFunctionExit
}


function Remove-LWNetworkSwitch
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseCompatibleCmdlets", "", Justification="Not relevant on Linux")]
    param (
        [Parameter(Mandatory)]
        [string]$Name
    )

    Write-LogFunctionEntry

    if (-not (Get-VMSwitch -Name $Name -ErrorAction SilentlyContinue))
    {
        Write-ScreenInfo 'The network switch does not exist' -Type Warning
        return
    }

    if ((Get-LWHypervVM -ErrorAction SilentlyContinue | Get-VMNetworkAdapter | Where-Object {$_.SwitchName -eq $Name} | Measure-Object).Count -eq 0)
    {
        try {
            $config = Get-NetAdapter | Where-Object Name -Match "^vEthernet \($($Name)\) ?(\d{1,2})?"
            if ($config.InterfaceIndex) {
                Get-NetIpAddress -IfIndex $config.InterfaceIndex -ErrorAction SilentlyContinue | Remove-NetIPAddress -Confirm:$false -ErrorAction Stop
            }
        }
        catch {
            Write-ScreenInfo -Type Error -Message "Unable to remove NAT Gateway IP, $($_.Exception.Message)"
        }

        try {
            Get-NetNat -Name $Name -ErrorAction SilentlyContinue | Remove-NetNat -Confirm:$false -ErrorAction Stop
        }
        catch {
            Write-ScreenInfo -Type Error -Message "Unable to remove NAT Gateway, $($_.Exception.Message)"
        }

        try
        {
            Remove-VMSwitch -Name $Name -Force -ErrorAction Stop
        }
        catch
        {
            Start-Sleep -Seconds 2
            Remove-VMSwitch -Name $Name -Force

            $networkDescription = Join-Path -Path (Get-Lab).LabPath -ChildPath "Network_$Name.xml"
            if (Test-Path -Path $networkDescription) {
                Remove-Item -Path $networkDescription
            }
        }

        Write-PSFMessage "Network switch '$Name' removed"
    }
    else
    {
        Write-ScreenInfo "Network switch '$Name' is still in use, skipping removal" -Type Warning
    }

    Write-LogFunctionExit
}


function Set-LWHypervNetworkSwitchDescription
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseCompatibleCmdlets", "", Justification="Not relevant on Linux")]
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [hashtable]$Hashtable,

        [Parameter(Mandatory)]
        [string]$NetworkSwitchName
    )

    Write-LogFunctionEntry

    $notePath = Join-Path -Path (Get-Lab).LabPath -ChildPath "Network_$NetworkSwitchName.xml"

    $type = Get-Type -GenericType AutomatedLab.DictionaryXmlStore -T string, string
    $dictionary = New-Object $type

    foreach ($kvp in $Hashtable.GetEnumerator())
    {
        $dictionary.Add($kvp.Key, $kvp.Value)
    }

    $dictionary.Export($notePath)

    Write-LogFunctionExit
}


function Get-LWAzureWindowsFeature
{
    [cmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [AutomatedLab.Machine[]]$Machine,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string[]]$FeatureName,

        [switch]$UseLocalCredential,

        [switch]$AsJob
    )

    Write-LogFunctionEntry

    $activityName = "Get Windows Feature(s): '$($FeatureName -join ', ')'"

    $result = @()
    foreach ($m in $Machine)
    {
        if ($m.OperatingSystem.Version -ge [System.Version]'6.2')
        {
            if ($m.OperatingSystem.Installation -eq 'Client')
            {
                if ($FeatureName.Count -gt 1)
                {
                    foreach ($feature in $FeatureName)
                    {
                        $cmd = [scriptblock]::Create("Get-WindowsOptionalFeature -Online -FeatureName $($feature) -WarningAction SilentlyContinue")
                        $result += Invoke-LabCommand -ComputerName $m -ActivityName $activityName -NoDisplay -ScriptBlock $cmd -UseLocalCredential:$UseLocalCredential -AsJob:$AsJob -PassThru
                    }
                }
                else
                {
                    $cmd = [scriptblock]::Create("Get-WindowsOptionalFeature -Online -FeatureName $($FeatureName -join ', ') -WarningAction SilentlyContinue")
                    $result += Invoke-LabCommand -ComputerName $m -ActivityName $activityName -NoDisplay -ScriptBlock $cmd -UseLocalCredential:$UseLocalCredential -AsJob:$AsJob -PassThru
                }
            }
            else
            {
                $cmd = [scriptblock]::Create("Get-WindowsFeature $($FeatureName -join ', ')  -WarningAction SilentlyContinue | Select-Object Description, DisplayName,FeatureType,Installed,InstallState,Name, Path, PostConfigurationNeeded")
                $result += Invoke-LabCommand -ComputerName $m -ActivityName $activityName -NoDisplay -ScriptBlock $cmd -UseLocalCredential:$UseLocalCredential -AsJob:$AsJob -PassThru
            }
        }
        else
        {
            if ($m.OperatingSystem.Installation -eq 'Client')
            {
                if ($FeatureName.Count -gt 1)
                {
                    foreach ($feature in $FeatureName)
                    {
                        $cmd = [scriptblock]::Create("DISM /online /get-featureinfo /featurename:$($feature)")
                        $featureList = Invoke-LabCommand -ComputerName $m -ActivityName $activityName -NoDisplay -ScriptBlock $cmd -UseLocalCredential:$UseLocalCredential -AsJob:$AsJob -PassThru

                        $parseddismOutput = $featureList | Select-String -Pattern "Feature Name :", "State :", "Restart Required :"
                        [string]$featureNamedismOutput = $parseddismOutput[0]
                        [string]$featureRRdismOutput = $parseddismOutput[1]
                        [string]$featureStatedismOutput = $parseddismOutput[2]


                        $result += [PSCustomObject]@{
                            FeatureName     = $featureNamedismOutput.Split(":")[1].Trim()
                            RestartRequired = $featureRRdismOutput.Split(":")[1].Trim()
                            State           = $featureStatedismOutput.Split(":")[1].Trim()
                        }
                    }
                }
                else
                {
                    $cmd = [scriptblock]::Create("DISM /online /get-featureinfo /featurename:$($FeatureName)")
                    $featureList = Invoke-LabCommand -ComputerName $m -ActivityName $activityName -NoDisplay -ScriptBlock $cmd -UseLocalCredential:$UseLocalCredential -AsJob:$AsJob -PassThru
                    $parseddismOutput = $featureList | Select-String -Pattern "Feature Name :", "State :", "Restart Required :"

                    [string]$featureNamedismOutput = $parseddismOutput[0]
                    [string]$featureRRdismOutput = $parseddismOutput[1]
                    [string]$featureStatedismOutput = $parseddismOutput[2]


                    $result += [PSCustomObject]@{
                        FeatureName     = $featureNamedismOutput.Split(":")[1].Trim()
                        RestartRequired = $featureRRdismOutput.Split(":")[1].Trim()
                        State           = $featureStatedismOutput.Split(":")[1].Trim()
                    }
                }
            }
            else
            {
                $cmd = [scriptblock]::Create("`$null;Import-Module -Name ServerManager; Get-WindowsFeature $($FeatureName -join ', ') -WarningAction SilentlyContinue | Select-Object Description, DisplayName,FeatureType,Installed,InstallState,Name, Path, PostConfigurationNeeded")
                $result += Invoke-LabCommand -ComputerName $m -ActivityName $activityName -NoDisplay -ScriptBlock $cmd -UseLocalCredential:$UseLocalCredential -AsJob:$AsJob -PassThru
            }
        }
    }

    if ($PassThru)
    {
        $result
    }

    Write-LogFunctionExit
}


function Get-LWHypervWindowsFeature
{
    [cmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [AutomatedLab.Machine[]]$Machine,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string[]]$FeatureName,

        [switch]$UseLocalCredential,

        [switch]$AsJob
    )

    Write-LogFunctionEntry

    $activityName = "Get Windows Feature(s): '$($FeatureName -join ', ')'"

    $result = @()
    foreach ($m in $Machine)
    {
        if ($m.OperatingSystem.Version -ge [System.Version]'6.2')
        {
            if ($m.OperatingSystem.Installation -eq 'Client')
            {
                if ($FeatureName.Count -gt 1)
                {
                    foreach ($feature in $FeatureName)
                    {
                        $cmd = [scriptblock]::Create("Get-WindowsOptionalFeature -Online -FeatureName $($feature) -WarningAction SilentlyContinue")
                        $result += Invoke-LabCommand -ComputerName $m -ActivityName $activityName -NoDisplay -ScriptBlock $cmd -UseLocalCredential:$UseLocalCredential -AsJob:$AsJob -PassThru
                    }
                }
                else
                {
                    $cmd = [scriptblock]::Create("Get-WindowsOptionalFeature -Online -FeatureName $($FeatureName) -WarningAction SilentlyContinue")
                    $result += Invoke-LabCommand -ComputerName $m -ActivityName $activityName -NoDisplay -ScriptBlock $cmd -UseLocalCredential:$UseLocalCredential -AsJob:$AsJob -PassThru
                }
            }
            else
            {
                $cmd = [scriptblock]::Create("Get-WindowsFeature $($FeatureName -join ', ') -WarningAction SilentlyContinue")
                $result += Invoke-LabCommand -ComputerName $m -ActivityName $activityName -NoDisplay -ScriptBlock $cmd -UseLocalCredential:$UseLocalCredential -AsJob:$AsJob -PassThru
            }
        }
        else
        {
            if ($m.OperatingSystem.Installation -eq 'Client')
            {
                if ($FeatureName.Count -gt 1)
                {
                    foreach ($feature in $FeatureName)
                    {
                        $cmd = [scriptblock]::Create("DISM /online /get-featureinfo /featurename:$($feature)")
                        $featureList = Invoke-LabCommand -ComputerName $m -ActivityName $activityName -NoDisplay -ScriptBlock $cmd -UseLocalCredential:$UseLocalCredential -AsJob:$AsJob -PassThru

                        $parseddismOutput = $featureList | Select-String -Pattern "Feature Name :", "State :", "Restart Required :"
                        [string]$featureNamedismOutput = $parseddismOutput[0]
                        [string]$featureRRdismOutput = $parseddismOutput[1]
                        [string]$featureStatedismOutput = $parseddismOutput[2]


                        $result += [PSCustomObject]@{
                            FeatureName     = $featureNamedismOutput.Split(":")[1].Trim()
                            RestartRequired = $featureRRdismOutput.Split(":")[1].Trim()
                            State           = $featureStatedismOutput.Split(":")[1].Trim()
                        }
                    }
                }
                else
                {
                    $cmd = [scriptblock]::Create("DISM /online /get-featureinfo /featurename:$($FeatureName)")
                    $featureList = Invoke-LabCommand -ComputerName $m -ActivityName $activityName -NoDisplay -ScriptBlock $cmd -UseLocalCredential:$UseLocalCredential -AsJob:$AsJob -PassThru
                    $parseddismOutput = $featureList | Select-String -Pattern "Feature Name :", "State :", "Restart Required :"

                    [string]$featureNamedismOutput = $parseddismOutput[0]
                    [string]$featureRRdismOutput = $parseddismOutput[1]
                    [string]$featureStatedismOutput = $parseddismOutput[2]


                    $result += [PSCustomObject]@{
                        FeatureName     = $featureNamedismOutput.Split(":")[1].Trim()
                        RestartRequired = $featureRRdismOutput.Split(":")[1].Trim()
                        State           = $featureStatedismOutput.Split(":")[1].Trim()
                    }
                }
            }
            else
            {
                $cmd = [scriptblock]::Create("`$null;Import-Module -Name ServerManager; Get-WindowsFeature $($FeatureName -join ', ') -WarningAction SilentlyContinue")
                $result += Invoke-LabCommand -ComputerName $m -ActivityName $activityName -NoDisplay -ScriptBlock $cmd -UseLocalCredential:$UseLocalCredential -AsJob:$AsJob -PassThru
            }
        }
    }

    $result

    Write-LogFunctionExit
}


function Install-LWAzureWindowsFeature
{
    [cmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [AutomatedLab.Machine[]]$Machine,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string[]]$FeatureName,

        [switch]$IncludeAllSubFeature,

        [switch]$IncludeManagementTools,

        [switch]$UseLocalCredential,

        [switch]$AsJob,

        [switch]$PassThru
    )

    Write-LogFunctionEntry

    $activityName = "Install Windows Feature(s): '$($FeatureName -join ', ')'"

    $result = @()
    foreach ($m in $Machine)
    {
        if ($m.OperatingSystem.Version -ge [System.Version]'6.2')
        {
            if ($m.OperatingSystem.Installation -eq 'Client')
            {
                $cmd = [scriptblock]::Create("Enable-WindowsOptionalFeature -Online -FeatureName $($FeatureName -join ', ') -Source 'C:\Windows\WinSXS' -All:`$$IncludeAllSubFeature -NoRestart -WarningAction SilentlyContinue")
                $result += Invoke-LabCommand -ComputerName $m -ActivityName $activityName -NoDisplay -ScriptBlock $cmd -UseLocalCredential:$UseLocalCredential -AsJob:$AsJob -PassThru:$PassThru
            }
            else
            {
                $cmd = [scriptblock]::Create("Install-WindowsFeature $($FeatureName -join ', ') -Source 'C:\Windows\WinSXS' -IncludeAllSubFeature:`$$IncludeAllSubFeature -IncludeManagementTools:`$$IncludeManagementTools -WarningAction SilentlyContinue")
                $result += Invoke-LabCommand -ComputerName $m -ActivityName $activityName -NoDisplay -ScriptBlock $cmd -UseLocalCredential:$UseLocalCredential -AsJob:$AsJob -PassThru:$PassThru
            }
        }
        else
        {
            if ($m.OperatingSystem.Installation -eq 'Client')
            {
                if ($FeatureName.Count -gt 1)
                {
                    foreach ($feature in $FeatureName)
                    {
                        $cmd = [scriptblock]::Create("DISM /online /enable-feature /featurename:$($feature)")
                        $result += Invoke-LabCommand -ComputerName $m -ActivityName $activityName -NoDisplay -ScriptBlock $cmd -UseLocalCredential:$UseLocalCredential -AsJob:$AsJob -PassThru:$PassThru
                    }
                }
                else
                {
                    $cmd = [scriptblock]::Create("DISM /online /enable-feature /featurename:$($feature)")
                    $result += Invoke-LabCommand -ComputerName $m -ActivityName $activityName -NoDisplay -ScriptBlock $cmd -UseLocalCredential:$UseLocalCredential -AsJob:$AsJob -PassThru:$PassThru
                }
            }
            else
            {
                $cmd = [scriptblock]::Create("`$null;Import-Module -Name ServerManager; Add-WindowsFeature $($FeatureName -join ', ') -IncludeAllSubFeature:`$$IncludeAllSubFeature -WarningAction SilentlyContinue")
                $result += Invoke-LabCommand -ComputerName $m -ActivityName $activityName -NoDisplay -ScriptBlock $cmd -UseLocalCredential:$UseLocalCredential -AsJob:$AsJob -PassThru:$PassThru
            }
        }
    }

    if ($PassThru)
    {
        $result
    }

    Write-LogFunctionExit
}


function Install-LWHypervWindowsFeature
{
    [cmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [AutomatedLab.Machine[]]$Machine,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string[]]$FeatureName,

        [switch]$IncludeAllSubFeature,

        [switch]$IncludeManagementTools,

        [switch]$UseLocalCredential,

        [switch]$AsJob,

        [switch]$PassThru
    )

    Write-LogFunctionEntry

    $activityName = "Install Windows Feature(s): '$($FeatureName -join ', ')'"

    $result = @()
    foreach ($m in $Machine)
    {
        if ($m.OperatingSystem.Version -ge [System.Version]'6.2')
        {
            if ($m.OperatingSystem.Installation -eq 'Client')
            {
                $cmd = [scriptblock]::Create("Enable-WindowsOptionalFeature -Online -FeatureName $($FeatureName -join ', ') -Source ""`$(@(Get-WmiObject -Class Win32_CDRomDrive)[-1].Drive)\sources\sxs"" -All:`$$IncludeAllSubFeature -NoRestart -WarningAction SilentlyContinue")
                $result += Invoke-LabCommand -ComputerName $m -ActivityName $activityName -NoDisplay -ScriptBlock $cmd -UseLocalCredential:$UseLocalCredential -AsJob:$AsJob -PassThru:$PassThru
            }
            else
            {
                $cmd = [scriptblock]::Create("Install-WindowsFeature $($FeatureName -join ', ') -Source ""`$(@(Get-WmiObject -Class Win32_CDRomDrive)[-1].Drive)\sources\sxs"" -IncludeAllSubFeature:`$$IncludeAllSubFeature -IncludeManagementTools:`$$IncludeManagementTools -WarningAction SilentlyContinue")
                $result += Invoke-LabCommand -ComputerName $m -ActivityName $activityName -NoDisplay -ScriptBlock $cmd -UseLocalCredential:$UseLocalCredential -AsJob:$AsJob -PassThru:$PassThru
            }
        }
        else
        {
            if ($m.OperatingSystem.Installation -eq 'Client')
            {
                if ($FeatureName.Count -gt 1)
                {
                    foreach ($feature in $FeatureName)
                    {
                        $cmd = [scriptblock]::Create("DISM /online /enable-feature /featurename:$($feature)")
                        $result += Invoke-LabCommand -ComputerName $m -ActivityName $activityName -NoDisplay -ScriptBlock $cmd -UseLocalCredential:$UseLocalCredential -AsJob:$AsJob -PassThru:$PassThru
                    }
                }
                else
                {
                    $cmd = [scriptblock]::Create("DISM /online /enable-feature /featurename:$($feature)")
                    $result += Invoke-LabCommand -ComputerName $m -ActivityName $activityName -NoDisplay -ScriptBlock $cmd -UseLocalCredential:$UseLocalCredential -AsJob:$AsJob -PassThru:$PassThru
                }
            }
            else
            {
                $cmd = [scriptblock]::Create("`$null;Import-Module -Name ServerManager; Add-WindowsFeature $($FeatureName -join ', ') -IncludeAllSubFeature:`$$IncludeAllSubFeature -WarningAction SilentlyContinue")
                $result += Invoke-LabCommand -ComputerName $m -ActivityName $activityName -NoDisplay -ScriptBlock $cmd -UseLocalCredential:$UseLocalCredential -AsJob:$AsJob -PassThru:$PassThru
            }
        }
    }

    if ($PassThru)
    {
        $result
    }

    Write-LogFunctionExit
}


function Invoke-LWCommand
{
    param (
        [Parameter(Mandatory)]
        [string[]]$ComputerName,

        [Parameter(Mandatory)]
        [System.Management.Automation.Runspaces.PSSession[]]$Session,

        [string]$ActivityName,

        [Parameter(Mandatory, ParameterSetName = 'FileContentDependencyLocalScript')]
        [Parameter(Mandatory, ParameterSetName = 'FileContentDependencyRemoteScript')]
        [Parameter(Mandatory, ParameterSetName = 'FileContentDependencyScriptBlock')]
        [string]$DependencyFolderPath,

        [Parameter(Mandatory, ParameterSetName = 'FileContentDependencyLocalScript')]
        [Parameter(Mandatory, ParameterSetName = 'IsoImageDependencyLocalScript')]
        [Parameter(Mandatory, ParameterSetName = 'NoDependencyLocalScript')]
        [string]$ScriptFilePath,

        [Parameter(Mandatory, ParameterSetName = 'FileContentDependencyRemoteScript')]
        [string]$ScriptFileName,

        [Parameter(Mandatory, ParameterSetName = 'IsoImageDependencyScriptBlock')]
        [Parameter(Mandatory, ParameterSetName = 'FileContentDependencyScriptBlock')]
        [Parameter(Mandatory, ParameterSetName = 'NoDependencyScriptBlock')]
        [scriptblock]$ScriptBlock,

        [Parameter(ParameterSetName = 'FileContentDependencyRemoteScript')]
        [Parameter(ParameterSetName = 'FileContentDependencyLocalScript')]
        [Parameter(ParameterSetName = 'FileContentDependencyScriptBlock')]
        [switch]$KeepFolder,

        [Parameter(Mandatory, ParameterSetName = 'IsoImageDependencyScriptBlock')]
        [Parameter(Mandatory, ParameterSetName = 'IsoImageDependencyLocalScript')]
        [Parameter(Mandatory, ParameterSetName = 'IsoImageDependencyScript')]
        [string]$IsoImagePath,

        [object[]]$ArgumentList,

        [string]$ParameterVariableName,

        [Parameter(ParameterSetName = 'IsoImageDependencyScriptBlock')]
        [Parameter(ParameterSetName = 'FileContentDependencyScriptBlock')]
        [Parameter(ParameterSetName = 'NoDependencyScriptBlock')]
        [Parameter(ParameterSetName = 'FileContentDependencyRemoteScript')]
        [Parameter(Mandatory, ParameterSetName = 'FileContentDependencyLocalScript')]
        [Parameter(Mandatory, ParameterSetName = 'IsoImageDependencyLocalScript')]
        [Parameter(Mandatory, ParameterSetName = 'NoDependencyLocalScript')]
        [int]$Retries,

        [Parameter(ParameterSetName = 'IsoImageDependencyScriptBlock')]
        [Parameter(ParameterSetName = 'FileContentDependencyScriptBlock')]
        [Parameter(ParameterSetName = 'NoDependencyScriptBlock')]
        [Parameter(ParameterSetName = 'FileContentDependencyRemoteScript')]
        [Parameter(Mandatory, ParameterSetName = 'FileContentDependencyLocalScript')]
        [Parameter(Mandatory, ParameterSetName = 'IsoImageDependencyLocalScript')]
        [Parameter(Mandatory, ParameterSetName = 'NoDependencyLocalScript')]
        [int]$RetryIntervalInSeconds,

        [int]$ThrottleLimit = 32,

        [switch]$AsJob,

        [switch]$PassThru
    )

    Write-LogFunctionEntry

    #required to supress verbose messages, warnings and errors
    Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState

    if ($DependencyFolderPath)
    {
        $result = if ((Get-Lab).DefaultVirtualizationEngine -eq 'Azure' -and (Test-LabPathIsOnLabAzureLabSourcesStorage -Path $DependencyFolderPath) )
        { 
            Test-LabPathIsOnLabAzureLabSourcesStorage -Path $DependencyFolderPath
        }
        else
        {
            Test-Path -Path $DependencyFolderPath
        }
        
        if (-not $result)
        {
            Write-Error "The DependencyFolderPath '$DependencyFolderPath' could not be found"
            return
        }
    }

    if ($ScriptFilePath)
    {
        $result = if ((Get-Lab).DefaultVirtualizationEngine -eq 'Azure' -and (Test-LabPathIsOnLabAzureLabSourcesStorage -Path $ScriptFilePath))
        {
            Test-LabPathIsOnLabAzureLabSourcesStorage -Path $ScriptFilePath
        }
        else
        {
            Test-Path -Path $ScriptFilePath
        }
        
        if (-not $result)
        {
            Write-Error "The ScriptFilePath '$ScriptFilePath' could not be found"
            return
        }
    }

    $internalSession = New-Object System.Collections.ArrayList
    $internalSession.AddRange(
        @($Session | Foreach-Object {
                if ($_.State -eq 'Broken')
                {
                    New-LabPSSession -Session $_ -ErrorAction SilentlyContinue
                }
                else
                {
                    $_
                }
        } | Where-Object {$_}) # Remove empty values. Invoke-LWCommand fails too early if AsJob is present and a broken session cannot be recreated
    )

    if (-not $ActivityName)
    {
        $ActivityName = '<unnamed>'
    }
    Write-PSFMessage -Message "Starting Activity '$ActivityName'"

    #if the image path is set we mount the image to the VM
    if ($PSCmdlet.ParameterSetName -like 'FileContentDependency*')
    {
        Write-PSFMessage -Message "Copying files from '$DependencyFolderPath' to $ComputerName..."

        if ((Get-Lab).DefaultVirtualizationEngine -eq 'Azure' -and (Test-LabPathIsOnLabAzureLabSourcesStorage -Path $DependencyFolderPath))
        {
            Invoke-Command -Session $Session -ScriptBlock { Copy-Item -Path $args[0] -Destination / -Recurse -Force } -ArgumentList $DependencyFolderPath
        }
        else
        {
            try
            {
                Copy-LabFileItem -Path $DependencyFolderPath -ComputerName $ComputerName -ErrorAction Stop
            }
            catch
            {
                if ((Get-Item -Path $DependencyFolderPath).PSIsContainer)
                {
                    Send-Directory -SourceFolderPath $DependencyFolderPath -DestinationFolder (Join-Path -Path (Get-LabConfigurationItem -Name OsRoot) -ChildPath (Split-Path -Path $DependencyFolderPath -Leaf)) -Session $internalSession
                }
                else
                {
                    Send-File -SourceFilePath $DependencyFolderPath -DestinationFolderPath (Get-LabConfigurationItem -Name OsRoot) -Session $internalSession
                }
            }
        }

        if ($PSCmdlet.ParameterSetName -eq 'FileContentDependencyRemoteScript')
        {
            $cmd = ''
            if ($ScriptFileName)
            {
                $cmd += "& '$(Join-Path -Path / -ChildPath (Split-Path $DependencyFolderPath -Leaf))\$ScriptFileName'"
            }
            if ($ParameterVariableName)
            {
                $cmd += " @$ParameterVariableName"
            }
            $cmd += "`n"
            if (-not $KeepFolder)
            {
                $cmd += "Remove-Item '$(Join-Path -Path C:\ -ChildPath (Split-Path $DependencyFolderPath -Leaf))' -Recurse -Force"
            }

            Write-PSFMessage -Message "Invoking script '$ScriptFileName'"

            $parameters = @{ }
            $parameters.Add('Session', $internalSession)
            $parameters.Add('ScriptBlock', [scriptblock]::Create($cmd))
            $parameters.Add('ArgumentList', $ArgumentList)
            if ($AsJob)
            {
                $parameters.Add('AsJob', $AsJob)
                $parameters.Add('JobName', $ActivityName)
            }
            if ($PSBoundParameters.ContainsKey('ThrottleLimit'))
            {
                $parameters.Add('ThrottleLimit', $ThrottleLimit)
            }
        }
        else
        {
            $parameters = @{ }
            $parameters.Add('Session', $internalSession)
            if ($ScriptFilePath)
            {
                $parameters.Add('FilePath', (Join-Path -Path $DependencyFolderPath -ChildPath $ScriptFilePath))
            }
            if ($ScriptBlock)
            {
                $parameters.Add('ScriptBlock', $ScriptBlock)
            }
            $parameters.Add('ArgumentList', $ArgumentList)
            if ($AsJob)
            {
                $parameters.Add('AsJob', $AsJob)
                $parameters.Add('JobName', $ActivityName)
            }
            if ($PSBoundParameters.ContainsKey('ThrottleLimit'))
            {
                $parameters.Add('ThrottleLimit', $ThrottleLimit)
            }
        }
    }
    elseif ($PSCmdlet.ParameterSetName -like 'NoDependency*')
    {
        $parameters = @{ }
        $parameters.Add('Session', $internalSession)
        if ($ScriptFilePath)
        {
            $parameters.Add('FilePath', $ScriptFilePath)
        }
        if ($ScriptBlock)
        {
            $parameters.Add('ScriptBlock', $ScriptBlock)
        }
        $parameters.Add('ArgumentList', $ArgumentList)
        if ($AsJob)
        {
            $parameters.Add('AsJob', $AsJob)
            $parameters.Add('JobName', $ActivityName)
        }
        if ($PSBoundParameters.ContainsKey('ThrottleLimit'))
        {
            $parameters.Add('ThrottleLimit', $ThrottleLimit)
        }
    }

    if ($VerbosePreference -eq 'Continue') { $parameters.Add('Verbose', $VerbosePreference) }
    if ($DebugPreference -eq 'Continue') { $parameters.Add('Debug', $DebugPreference) }

    [System.Collections.ArrayList]$result = New-Object System.Collections.ArrayList

    if (-not $AsJob -and $parameters.ScriptBlock)
    {
        Write-Debug 'Adding LABHOSTNAME to scriptblock'
        #in some situations a retry makes sense. In order to know which machines have done the job, the scriptblock must return the hostname
        $parameters.ScriptBlock = [scriptblock]::Create($parameters.ScriptBlock.ToString() + "`n;`"LABHOSTNAME:`$([System.Net.Dns]::GetHostName())`"`n")
    }

    if ($AsJob)
    {
        $job = Invoke-Command @parameters -ErrorAction SilentlyContinue
    }
    else
    {
        while ($Retries -gt 0 -and $internalSession.Count -gt 0)
        {
            $nonAvailableSessions = @($internalSession | Where-Object State -ne Opened)
            foreach ($nonAvailableSession in $nonAvailableSessions)
            {
                Write-PSFMessage "Re-creating unavailable session for machine '$($nonAvailableSessions.ComputerName)'"
                $internalSession.Add((New-LabPSSession -Session $nonAvailableSession)) | Out-Null
                Write-PSFMessage "removing unavailable session for machine '$($nonAvailableSessions.ComputerName)'"
                $internalSession.Remove($nonAvailableSession)
            }

            $result.AddRange(@(Invoke-Command @parameters))

            #remove all sessions for machines successfully invoked the command
            foreach ($machineFinished in ($result | Where-Object { $_ -like 'LABHOSTNAME*' }))
            {
                $machineFinishedName = $machineFinished.Substring($machineFinished.IndexOf(':') + 1)
                $internalSession.Remove(($internalSession | Where-Object LabMachineName -eq $machineFinishedName))
            }
            $result = @($result | Where-Object { $_ -notlike 'LABHOSTNAME*' })

            $Retries--

            if ($Retries -gt 0 -and $internalSession.Count -gt 0)
            {
                Write-PSFMessage "Scriptblock did not run on all machines, retrying (Retries = $Retries)"
                Start-Sleep -Seconds $RetryIntervalInSeconds
            }
        }
    }

    if ($PassThru)
    {
        if ($AsJob)
        {
            $job
        }
        else
        {
            $result
        }
    }
    else
    {
        $resultVariable = New-Variable -Name ("AL_$([guid]::NewGuid().Guid)") -Scope Global -PassThru
        $resultVariable.Value = $result
        Write-PSFMessage "The Output of the task on machine '$($ComputerName)' will be available in the variable '$($resultVariable.Name)'"
    }

    Write-PSFMessage -Message "Finished Installation Activity '$ActivityName'"

    Write-LogFunctionExit -ReturnValue $resultVariable
}


function Uninstall-LWAzureWindowsFeature
{
    [cmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [AutomatedLab.Machine[]]$Machine,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string[]]$FeatureName,

        [switch]$IncludeManagementTools,

        [switch]$UseLocalCredential,

        [switch]$AsJob,

        [switch]$PassThru
    )

    Write-LogFunctionEntry

    $activityName = "Uninstall Windows Feature(s): '$($FeatureName -join ', ')'"

    $result = @()
    foreach ($m in $Machine)
    {
        if ($m.OperatingSystem.Version -ge [System.Version]'6.2')
        {
            if ($m.OperatingSystem.Installation -eq 'Client')
            {
                $cmd = [scriptblock]::Create("Disable-WindowsOptionalFeature -Online -FeatureName $($FeatureName -join ', ') -NoRestart -WarningAction SilentlyContinue")
                $result += Invoke-LabCommand -ComputerName $m -ActivityName $activityName -NoDisplay -ScriptBlock $cmd -UseLocalCredential:$UseLocalCredential -AsJob:$AsJob -PassThru:$PassThru
            }
            else
            {
                $cmd = [scriptblock]::Create("Uninstall-WindowsFeature $($FeatureName -join ', ') -IncludeManagementTools:`$$IncludeManagementTools -WarningAction SilentlyContinue")
                $result += Invoke-LabCommand -ComputerName $m -ActivityName $activityName -NoDisplay -ScriptBlock $cmd -UseLocalCredential:$UseLocalCredential -AsJob:$AsJob -PassThru:$PassThru
            }
        }
        else
        {
            if ($m.OperatingSystem.Installation -eq 'Client')
            {
                if ($FeatureName.Count -gt 1)
                {
                    foreach ($feature in $FeatureName)
                    {
                        $cmd = [scriptblock]::Create("DISM /online /disable-feature /featurename:$($feature)")
                        $result += Invoke-LabCommand -ComputerName $m -ActivityName $activityName -NoDisplay -ScriptBlock $cmd -UseLocalCredential:$UseLocalCredential -AsJob:$AsJob -PassThru:$PassThru
                    }
                }
                else
                {
                    $cmd = [scriptblock]::Create("DISM /online /disable-feature /featurename:$($feature)")
                    $result += Invoke-LabCommand -ComputerName $m -ActivityName $activityName -NoDisplay -ScriptBlock $cmd -UseLocalCredential:$UseLocalCredential -AsJob:$AsJob -PassThru:$PassThru
                }
            }
            else
            {
                $cmd = [scriptblock]::Create("`$null;Import-Module -Name ServerManager; Remove-WindowsFeature $($FeatureName -join ', ') -WarningAction SilentlyContinue")
                $result += Invoke-LabCommand -ComputerName $m -ActivityName $activityName -NoDisplay -ScriptBlock $cmd -UseLocalCredential:$UseLocalCredential -AsJob:$AsJob -PassThru:$PassThru
            }
        }
    }

    if ($PassThru)
    {
        $result
    }

    Write-LogFunctionExit
}


function Uninstall-LWHypervWindowsFeature
{
    [cmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [AutomatedLab.Machine[]]$Machine,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string[]]$FeatureName,

        [switch]$IncludeManagementTools,

        [switch]$UseLocalCredential,

        [switch]$AsJob,

        [switch]$PassThru
    )

    Write-LogFunctionEntry

    $activityName = "Uninstall Windows Feature(s): '$($FeatureName -join ', ')'"

    $result = @()
    foreach ($m in $Machine)
    {
        if ($m.OperatingSystem.Version -ge [System.Version]'6.2')
        {
            if ($m.OperatingSystem.Installation -eq 'Client')
            {
                $cmd = [scriptblock]::Create("Disable-WindowsOptionalFeature -Online -FeatureName $($FeatureName -join ', ') -NoRestart -WarningAction SilentlyContinue")
                $result += Invoke-LabCommand -ComputerName $m -ActivityName $activityName -NoDisplay -ScriptBlock $cmd -UseLocalCredential:$UseLocalCredential -AsJob:$AsJob -PassThru:$PassThru
            }
            else
            {
                $cmd = [scriptblock]::Create("Uninstall-WindowsFeature $($FeatureName -join ', ') -IncludeManagementTools:`$$IncludeManagementTools -WarningAction SilentlyContinue")
                $result += Invoke-LabCommand -ComputerName $m -ActivityName $activityName -NoDisplay -ScriptBlock $cmd -UseLocalCredential:$UseLocalCredential -AsJob:$AsJob -PassThru:$PassThru
            }
        }
        else
        {
            if ($m.OperatingSystem.Installation -eq 'Client')
            {
                if ($FeatureName.Count -gt 1)
                {
                    foreach ($feature in $FeatureName)
                    {
                        $cmd = [scriptblock]::Create("DISM /online /disable-feature /featurename:$($feature)")
                        $result += Invoke-LabCommand -ComputerName $m -ActivityName $activityName -NoDisplay -ScriptBlock $cmd -UseLocalCredential:$UseLocalCredential -AsJob:$AsJob -PassThru:$PassThru
                    }
                }
                else
                {
                    $cmd = [scriptblock]::Create("DISM /online /disable-feature /featurename:$($feature)")
                    $result += Invoke-LabCommand -ComputerName $m -ActivityName $activityName -NoDisplay -ScriptBlock $cmd -UseLocalCredential:$UseLocalCredential -AsJob:$AsJob -PassThru:$PassThru
                }
            }
            else
            {
                $cmd = [scriptblock]::Create("`$null;Import-Module -Name ServerManager; Remove-WindowsFeature $($FeatureName -join ', ') -WarningAction SilentlyContinue")
                $result += Invoke-LabCommand -ComputerName $m -ActivityName $activityName -NoDisplay -ScriptBlock $cmd -UseLocalCredential:$UseLocalCredential -AsJob:$AsJob -PassThru:$PassThru
            }
        }
    }

    if ($PassThru)
    {
        $result
    }

    Write-LogFunctionExit
}


function Wait-LWLabJob
{
    Param
    (
        [Parameter(Mandatory, ParameterSetName = 'ByJob')]
        [AllowNull()]
        [AllowEmptyCollection()]
        [System.Management.Automation.Job[]]$Job,

        [Parameter(Mandatory, ParameterSetName = 'ByName')]
        [string[]]$Name,

        [ValidateRange(0, 300)]
        [int]$ProgressIndicator = (Get-LabConfigurationItem -Name DefaultProgressIndicator),

        [int]$Timeout = 120,

        [switch]$NoNewLine,

        [switch]$NoDisplay,

        [switch]$PassThru
    )

    if (-not $PSBoundParameters.ContainsKey('ProgressIndicator')) { $PSBoundParameters.Add('ProgressIndicator', $ProgressIndicator) } #enables progress indicator

    Write-LogFunctionEntry

    Write-ProgressIndicator

    if (-not $Job -and -not $Name)
    {
        Write-PSFMessage 'There is no job to wait for'
        Write-LogFunctionExit
        return
    }

    $start = (Get-Date)

    if ($Job)
    {
        $jobs = Get-Job -Id $Job.ID
    }
    else
    {
        $jobs = Get-Job -Name $Name
    }

    Write-ScreenInfo -Message "Waiting for job(s) to complete with ID(s): $($jobs.Id -join ', ')" -TaskStart

    if ($jobs -and ($jobs.State -contains 'Running' -or $jobs.State -contains 'AtBreakpoint'))
    {
        $jobs = Get-Job -Id $jobs.ID
        $ProgressIndicatorTimer = Get-Date
        do
        {
            Start-Sleep -Seconds 1
            if (((Get-Date) - $ProgressIndicatorTimer).TotalSeconds -ge $ProgressIndicator)
            {
                Write-ProgressIndicator
                $ProgressIndicatorTimer = Get-Date
            }
        }
        until (($jobs.State -notcontains 'Running' -and $jobs.State -notcontains 'AtBreakPoint') -or ((Get-Date) -gt ($Start.AddMinutes($Timeout))))
    }

    Write-ProgressIndicatorEnd

    if ((Get-Date) -gt ($Start.AddMinutes($Timeout)))
    {
        $jobs = Get-Job -Id $jobs.Id | Where-Object State -eq Running
        Write-Error -Message "Timeout while waiting for job $($jobs.ID -join ', ')"
    }
    else
    {
        Write-ScreenInfo -Message 'Job(s) no longer running' -TaskEnd

        if ($PassThru)
        {
            $result = $jobs | Receive-Job -ErrorAction SilentlyContinue -ErrorVariable jobErrors
            $result
            #PSRemotingTransportException are very likely due to restarts or problems AL cannot recover
            $jobErrors = $jobErrors | Where-Object { $_.Exception -isnot [System.Management.Automation.Remoting.PSRemotingTransportException] }
            foreach ($jobError in $jobErrors)
            {
                Write-Error -ErrorRecord $jobError
            }
        }
    }

    Write-LogFunctionExit
}


function Add-LWAzureLoadBalancedPort
{
    param
    (
        [Parameter(Mandatory)]
        [uint16]
        $Port,

        [Parameter(Mandatory)]
        [uint16]
        $DestinationPort,

        [Parameter(Mandatory)]
        [string]
        $ComputerName
    )

    Test-LabHostConnected -Throw -Quiet

    if (Get-LabAzureLoadBalancedPort @PSBoundParameters)
    {
        Write-PSFMessage -Message ('Port {0} -> {1} already configured for {2}' -f $Port, $DestinationPort, $ComputerName)
        return
    }

    $lab = Get-Lab
    $resourceGroup = (Get-LabAzureDefaultResourceGroup).ResourceGroupName
    $machine = Get-LabVm -ComputerName $ComputerName
    $net = $lab.VirtualNetworks.Where({ $_.Name -eq $machine.Network[0] })

    $lb = Get-AzLoadBalancer -ResourceGroupName $resourceGroup | Where-Object {$_.Tag['Vnet'] -eq $net.ResourceName}
    if (-not $lb)
    {
        Write-PSFMessage "No load balancer found to add port rules to"
        return
    }

    $frontendConfig = $lb | Get-AzLoadBalancerFrontendIpConfig

    $lb = Add-AzLoadBalancerInboundNatRuleConfig -LoadBalancer $lb -Name "$($machine.ResourceName.ToLower())-$Port-$DestinationPort" -FrontendIpConfiguration $frontendConfig -Protocol Tcp -FrontendPort $Port -BackendPort $DestinationPort
    $lb = $lb | Set-AzLoadBalancer

    $vm = Get-AzVM -ResourceGroupName $resourceGroup -Name $machine.ResourceName
    $nic = $vm.NetworkProfile.NetworkInterfaces | Get-AzResource | Get-AzNetworkInterface
    $rules = Get-LWAzureLoadBalancedPort -ComputerName $ComputerName
    $nic.IpConfigurations[0].LoadBalancerInboundNatRules = $rules
    [void] ($nic | Set-AzNetworkInterface)

    # Extend NSG
    $nsg = Get-AzNetworkSecurityGroup -Name "nsg" -ResourceGroupName $resourceGroup

    $rule = $nsg | Get-AzNetworkSecurityRuleConfig -Name NecessaryPorts
    if (-not $rule.DestinationPortRange.Contains($DestinationPort))
    {
        $rule.DestinationPortRange.Add($DestinationPort)
        
        # Update the NSG.
        $nsg = $nsg | Set-AzNetworkSecurityRuleConfig -Name $rule.Name -DestinationPortRange $rule.DestinationPortRange -Protocol $rule.Protocol -SourcePortRange $rule.SourcePortRange -SourceAddressPrefix $rule.SourceAddressPrefix -DestinationAddressPrefix $rule.DestinationAddressPrefix -Access Allow -Priority $rule.Priority -Direction $rule.Direction
        $null = $nsg | Set-AzNetworkSecurityGroup
    }

    if (-not $machine.InternalNotes."AdditionalPort-$Port-$DestinationPort")
    {
        $machine.InternalNotes.Add("AdditionalPort-$Port-$DestinationPort", $DestinationPort)
    }

    $machine.InternalNotes."AdditionalPort-$Port-$DestinationPort" = $DestinationPort

    Export-Lab
}


function Get-LabAzureLoadBalancedPort
{
    param
    (
        [Parameter()]
        [uint16]
        $Port,

        [uint16]
        $DestinationPort,

        [Parameter(Mandatory)]
        [string]
        $ComputerName
    )

    $lab = Get-Lab -ErrorAction SilentlyContinue

    if (-not $lab)
    {
        Write-ScreenInfo -Type Warning -Message 'Lab data not available. Cannot list ports. Use Import-Lab to import an existing lab'
        return
    }

    $machine = Get-LabVm -ComputerName $ComputerName

    if (-not $machine)
    {
        Write-PSFMessage -Message "$ComputerName not found. Cannot list ports."
        return
    }

    $ports = if ($DestinationPort -and $Port)
    {
        $machine.InternalNotes.GetEnumerator() | Where-Object -Property Key -eq "AdditionalPort-$Port-$DestinationPort"
    }
    elseif ($DestinationPort)
    {
        $machine.InternalNotes.GetEnumerator() | Where-Object -Property Key -like "AdditionalPort-*-$DestinationPort"
    }
    elseif ($Port)
    {
        $machine.InternalNotes.GetEnumerator() | Where-Object -Property Key -like "AdditionalPort-$Port-*"
    }
    else
    {
        $machine.InternalNotes.GetEnumerator() | Where-Object -Property Key -like 'AdditionalPort*'
    }

    $ports | Foreach-Object {
        [pscustomobject]@{
            Port = ($_.Key -split '-')[1]
            DestinationPort = ($_.Key -split '-')[2]
            ComputerName = $machine.ResourceName
        }
    }
}


function Get-LWAzureLoadBalancedPort
{
    param
    (
        [Parameter()]
        [uint16]
        $Port,

        [Parameter()]
        [uint16]
        $DestinationPort,

        [Parameter(Mandatory)]
        [string]
        $ComputerName
    )

    Test-LabHostConnected -Throw -Quiet

    $lab = Get-Lab
    $resourceGroup = $lab.Name
    $machine = Get-LabVm -ComputerName $ComputerName
    $net = $lab.VirtualNetworks.Where({ $_.Name -eq $machine.Network[0] })

    $lb = Get-AzLoadBalancer -ResourceGroupName $resourceGroup | Where-Object {$_.Tag['Vnet'] -eq $net.ResourceName}
    if (-not $lb)
    {
        Write-PSFMessage "No load balancer found to list port rules of"
        return
    }

    $existingConfiguration = $lb | Get-AzLoadBalancerInboundNatRuleConfig

    # Port müssen unique sein, destination port + computername müssen unique sein
    if ($Port)
    {
        $filteredRules = $existingConfiguration | Where-Object -Property FrontendPort -eq $Port

        if (($filteredRules | Where-Object Name -notlike "$($machine.ResourceName)*"))
        {
            $err = ($filteredRules | Where-Object Name -notlike "$($machine.ResourceName)*")[0].Name
            $existingComputer = $err.Substring(0, $err.IndexOf('-'))
            Write-Error -Message ("Incoming port {0} is already mapped to {1}!" -f $Port, $existingComputer)
            return
        }

        return $filteredRules
    }

    if ($DestinationPort)
    {
        return ($existingConfiguration | Where-Object {$_.BackendPort -eq $DestinationPort -and $_.Name -like "$($machine.ResourceName)*"})
    }

    return ($existingConfiguration | Where-Object -Property Name -like "$($machine.ResourceName)*")
}


function Get-LWAzureNetworkSwitch
{
    param
    (
        [Parameter(Mandatory)]
        [AutomatedLab.VirtualNetwork[]]
        $virtualNetwork
    )

    Test-LabHostConnected -Throw -Quiet

    $lab = Get-Lab
    $jobs = @()

    foreach ($network in $VirtualNetwork)
    {
        Write-PSFMessage -Message "Locating Azure virtual network '$($network.ResourceName)'"

        $azureNetworkParameters = @{
            Name              = $network.ResourceName
            ResourceGroupName = (Get-LabAzureDefaultResourceGroup)
            ErrorAction       = 'SilentlyContinue'
            WarningAction     = 'SilentlyContinue'
        }

        Get-AzVirtualNetwork @azureNetworkParameters
    }
}


function Set-LWAzureDnsServer
{
    param
    (
        [Parameter(Mandatory)]
        [AutomatedLab.VirtualNetwork[]]
        $VirtualNetwork,

        [switch]
        $PassThru
    )

    Test-LabHostConnected -Throw -Quiet

    Write-LogFunctionEntry

    foreach ($network in $VirtualNetwork)
    {
        if ($network.DnsServers.Count -eq 0)
        {
            Write-PSFMessage -Message "Skipping $($network.ResourceName) because no DNS servers are configured"
            continue
        }

        Write-ScreenInfo -Message "Setting DNS servers for $($network.ResourceName)" -TaskStart
        $azureVnet = Get-LWAzureNetworkSwitch -VirtualNetwork $network -ErrorAction SilentlyContinue
        if (-not $azureVnet)
        {
            Write-Error "$($network.ResourceName) does not exist"
            continue
        }

        $azureVnet.DhcpOptions.DnsServers = New-Object -TypeName System.Collections.Generic.List[string]
        $network.DnsServers.AddressAsString | ForEach-Object { $azureVnet.DhcpOptions.DnsServers.Add($PSItem)}
        $null = $azureVnet | Set-AzVirtualNetwork -ErrorAction Stop

        if ($PassThru)
        {
            $azureVnet
        }

        Write-ScreenInfo -Message "Successfully set DNS servers for $($network.ResourceName)" -TaskEnd
    }

    Write-LogFunctionExit
}


Function Test-IpInSameSameNetwork
{
	param
    (
		[AutomatedLab.IPNetwork]$Ip1,
		[AutomatedLab.IPNetwork]$Ip2
	)

    $ip1Decimal = $Ip1.SerializationNetworkAddress
    $ip2Decimal = $Ip2.SerializationNetworkAddress
    $ip1Total   = $Ip1.Total
    $ip2Total   = $Ip2.Total

    if (($ip1Decimal -ge $ip2Decimal) -and ($ip1Decimal -lt ([long]$ip2Decimal+[long]$ip2Total)))
    {
        return $true
    }

    if (($ip2Decimal -ge $ip1Decimal) -and ($ip2Decimal -lt ([long]$ip1Decimal+[long]$ip1Total)))
    {
        return $true
    }

    return $false
}


function Add-LWVMVHDX
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseCompatibleCmdlets", "", Justification="Not relevant on Linux")]
    [Cmdletbinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [string]$VMName,

        [Parameter(Mandatory = $true)]
        [string]$VhdxPath
    )

    Write-LogFunctionEntry

    if (-not (Test-Path -Path $VhdxPath))
    {
        Write-Error 'VHDX cannot be found'
        return
    }

    $vm = Get-LWHypervVM -Name $VMName -ErrorAction SilentlyContinue
    if (-not $vm)
    {
        Write-Error 'VM cannot be found'
        return
    }

    Add-VMHardDiskDrive -VM $vm -Path $VhdxPath

    Write-LogFunctionExit
}


function New-LWReferenceVHDX
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseCompatibleCmdlets", "", Justification="Not relevant on Linux")]
    [Cmdletbinding()]
    Param (
        #ISO of OS
        [Parameter(Mandatory = $true)]
        [string]$IsoOsPath,

        #Path to reference VHD
        [Parameter(Mandatory = $true)]
        [string]$ReferenceVhdxPath,

        #Path to reference VHD
        [Parameter(Mandatory = $true)]
        [string]$OsName,

        #Real image name in ISO file
        [Parameter(Mandatory = $true)]
        [string]$ImageName,

        #Size of the reference VHD
        [Parameter(Mandatory = $true)]
        [int]$SizeInGB,

        [Parameter(Mandatory = $true)]
        [ValidateSet('MBR', 'GPT')]
        [string]$PartitionStyle
    )

    Write-LogFunctionEntry

    # Get start time
    $start = Get-Date
    Write-PSFMessage "Beginning at $start"

    try
    {
        $FDVDenyWriteAccess = (Get-ItemProperty -Path HKLM:\SYSTEM\CurrentControlSet\Policies\Microsoft\FVE -Name FDVDenyWriteAccess -ErrorAction SilentlyContinue).FDVDenyWriteAccess

        $imageList = Get-LabAvailableOperatingSystem -Path $IsoOsPath
        Write-PSFMessage "The Windows Image list contains $($imageList.Count) items"

        Write-PSFMessage "Mounting ISO image '$IsoOsPath'"
        [void] (Mount-DiskImage -ImagePath $IsoOsPath)

        Write-PSFMessage 'Getting disk image of the ISO'
        $isoImage = Get-DiskImage -ImagePath $IsoOsPath | Get-Volume
        Write-PSFMessage "Got disk image '$($isoImage.DriveLetter)'"

        $isoDrive = "$($isoImage.DriveLetter):"
        Write-PSFMessage "OS ISO mounted on drive letter '$isoDrive'"

        $image = $imageList | Where-Object OperatingSystemName -eq $OsName

        if (-not $image)
        {
            throw "The specified image ('$OsName') could not be found on the ISO '$(Split-Path -Path $IsoOsPath -Leaf)'. Please specify one of the following values: $($imageList.ImageName -join ', ')"
        }

        $imageIndex = $image.ImageIndex
        Write-PSFMessage "Selected image index '$imageIndex' with name '$($image.ImageName)'"

        $vmDisk = New-VHD -Path $ReferenceVhdxPath -SizeBytes ($SizeInGB * 1GB) -ErrorAction Stop
        Write-PSFMessage "Created VHDX file '$($vmDisk.Path)'"

        Write-ScreenInfo -Message "Creating base image for operating system '$OsName'" -NoNewLine -TaskStart

        [void] (Mount-DiskImage -ImagePath $ReferenceVhdxPath)
        $vhdDisk = Get-DiskImage -ImagePath $ReferenceVhdxPath | Get-Disk
        $vhdDiskNumber = [string]$vhdDisk.Number
        Write-PSFMessage "Reference image is on disk number '$vhdDiskNumber'"

        Initialize-Disk -Number $vhdDiskNumber -PartitionStyle $PartitionStyle | Out-Null
        if ($PartitionStyle -eq 'MBR')
        {
            if ($FDVDenyWriteAccess) {
                Set-ItemProperty -Path HKLM:\SYSTEM\CurrentControlSet\Policies\Microsoft\FVE -Name FDVDenyWriteAccess -Value 0
            }
            $vhdWindowsDrive = New-Partition -DiskNumber $vhdDiskNumber -UseMaximumSize -IsActive -AssignDriveLetter |
            Format-Volume -FileSystem NTFS -NewFileSystemLabel 'System' -Confirm:$false
        }
        else
        {
            $vhdRecoveryPartition = New-Partition -DiskNumber $vhdDiskNumber -GptType '{de94bba4-06d1-4d40-a16a-bfd50179d6ac}' -Size 300MB
            $vhdRecoveryDrive = $vhdRecoveryPartition | Format-Volume -FileSystem NTFS -NewFileSystemLabel 'Windows RE Tools' -Confirm:$false

            $recoveryPartitionNumber = (Get-Disk -Number $vhdDiskNumber | Get-Partition | Where-Object Type -eq Recovery).PartitionNumber
            $diskpartCmd = @"
select disk $vhdDiskNumber
select partition $recoveryPartitionNumber
gpt attributes=0x8000000000000001
exit
"@
            $diskpartCmd | diskpart.exe | Out-Null

            $systemPartition = New-Partition -DiskNumber $vhdDiskNumber -GptType '{c12a7328-f81f-11d2-ba4b-00a0c93ec93b}' -Size 100MB
            #does not work, seems to be a bug. Using diskpart as a workaround
            #$systemPartition | Format-Volume -FileSystem FAT32 -NewFileSystemLabel 'System' -Confirm:$false

            $diskpartCmd = @"
select disk $vhdDiskNumber
select partition $($systemPartition.PartitionNumber)
format quick fs=fat32 label=System
exit
"@
            $diskpartCmd | diskpart.exe | Out-Null

            $reservedPartition = New-Partition -DiskNumber $vhdDiskNumber -GptType '{e3c9e316-0b5c-4db8-817d-f92df00215ae}' -Size 128MB

            if ($FDVDenyWriteAccess) {
                Set-ItemProperty -Path HKLM:\SYSTEM\CurrentControlSet\Policies\Microsoft\FVE -Name FDVDenyWriteAccess -Value 0
            }
            $vhdWindowsDrive = New-Partition -DiskNumber $vhdDiskNumber -UseMaximumSize -AssignDriveLetter |
            Format-Volume -FileSystem NTFS -NewFileSystemLabel 'System' -Confirm:$false
        }

        $vhdWindowsVolume = "$($vhdWindowsDrive.DriveLetter):"
        Write-PSFMessage "VHD drive '$vhdWindowsDrive', Vhd volume '$vhdWindowsVolume'"

        Write-PSFMessage "Disabling Bitlocker Drive Encryption on drive $vhdWindowsVolume"
        if (Test-Path -Path C:\Windows\System32\manage-bde.exe)
        {
            manage-bde.exe -off $vhdWindowsVolume | Out-Null #without this on some devices (for exmaple Surface 3) the VHD was auto-encrypted
        }

        Write-PSFMessage 'Applying image to the volume...'

        $installFilePath = Get-Item -Path "$isoDrive\Sources\install.*" | Where-Object Name -Match '.*\.(esd|wim)'

        $job = Start-Job -ScriptBlock {
            $output = Dism.exe /English /apply-Image /ImageFile:$using:installFilePath /index:$using:imageIndex /ApplyDir:$using:vhdWindowsVolume\
            New-Object PSObject -Property @{
                Outout = $output
                LastExitCode = $LASTEXITCODE
            }
        }

        $dismResult = Wait-LWLabJob -Job $job -NoDisplay -ProgressIndicator 20 -Timeout 60 -PassThru
        if ($dismResult.LastExitCode)
        {
            throw (New-Object System.ComponentModel.Win32Exception($dismResult.LastExitCode,
            "The base image for operating system '$OsName' could not be created. The error is $($dismResult.LastExitCode)"))
        }
        Start-Sleep -Seconds 10

        Write-PSFMessage 'Setting BCDBoot'
        if ($PartitionStyle -eq 'MBR')
        {
            bcdboot.exe $vhdWindowsVolume\Windows /s $vhdWindowsVolume /f BIOS | Out-Null
        }
        else
        {
            $possibleDrives = [char[]](65..90)
            $drives = (Get-PSDrive -PSProvider FileSystem).Name
            $freeDrives = Compare-Object -ReferenceObject $possibleDrives -DifferenceObject $drives | Where-Object { $_.SideIndicator -eq '<=' }
            $freeDrive = ($freeDrives | Select-Object -First 1).InputObject

            $diskpartCmd = @"
    select disk $vhdDiskNumber
    select partition $($systemPartition.PartitionNumber)
    assign letter=$freeDrive
    exit
"@
            $diskpartCmd | diskpart.exe | Out-Null

            bcdboot.exe $vhdWindowsVolume\Windows /s "$($freeDrive):" /f UEFI | Out-Null

            $diskpartCmd = @"
    select disk $vhdDiskNumber
    select partition $($systemPartition.PartitionNumber)
    remove letter=$freeDrive
    exit
"@
            $diskpartCmd | diskpart.exe | Out-Null
        }
    }
    catch
    {
        Write-PSFMessage 'Dismounting ISO and new disk'
        [void] (Dismount-DiskImage -ImagePath $ReferenceVhdxPath)
        [void] (Dismount-DiskImage -ImagePath $IsoOsPath)
        Remove-Item -Path $ReferenceVhdxPath -Force #removing as the creation did not succeed
        if ($FDVDenyWriteAccess) {
            Set-ItemProperty -Path HKLM:\SYSTEM\CurrentControlSet\Policies\Microsoft\FVE -Name FDVDenyWriteAccess -Value $FDVDenyWriteAccess
        }

        throw $_.Exception
    }

    Write-PSFMessage 'Dismounting ISO and new disk'
    [void] (Dismount-DiskImage -ImagePath $ReferenceVhdxPath)
    [void] (Dismount-DiskImage -ImagePath $IsoOsPath)
    if ($FDVDenyWriteAccess) {
        Set-ItemProperty -Path HKLM:\SYSTEM\CurrentControlSet\Policies\Microsoft\FVE -Name FDVDenyWriteAccess -Value $FDVDenyWriteAccess
    }
    Write-ScreenInfo -Message 'Finished creating base image' -TaskEnd

    $end = Get-Date
    Write-PSFMessage "Runtime: '$($end - $start)'"

    Write-LogFunctionExit
}


function New-LWVHDX
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseCompatibleCmdlets", "", Justification="Not relevant on Linux")]
    [Cmdletbinding()]
    Param (
        #Path to reference VHD
        [Parameter(Mandatory = $true)]
        [string]$VhdxPath,

        #Size of the reference VHD
        [Parameter(Mandatory = $true)]
        [int]$SizeInGB,

        [string]$Label,

        [switch]$UseLargeFRS,

        [char]$DriveLetter,

        [long]$AllocationUnitSize,

        [string]$PartitionStyle,

        [switch]$SkipInitialize
    )

    Write-LogFunctionEntry

    $PSBoundParameters.Add('ProgressIndicator', 1) #enables progress indicator

    $VmDisk = New-VHD -Path $VhdxPath -SizeBytes ($SizeInGB * 1GB) -ErrorAction Stop
    Write-ProgressIndicator
    Write-PSFMessage "Created VHDX file '$($vmDisk.Path)'"

    if ($SkipInitialize)
    {
        Write-PSFMessage -Message "Skipping the initialization of '$($vmDisk.Path)'"
        Write-LogFunctionExit
        return
    }

    $mountedVhd = $VmDisk | Mount-VHD -PassThru
    Write-ProgressIndicator

    if ($DriveLetter)
    {
        $Label += "_AL_$DriveLetter"
    }

    $formatParams = @{
        FileSystem = 'NTFS'
        NewFileSystemLabel = 'Data'
        Force = $true
        Confirm = $false
        UseLargeFRS = $UseLargeFRS
        AllocationUnitSize = $AllocationUnitSize
    }
    if ($Label)
    {
        $formatParams.NewFileSystemLabel = $Label
    }

    $mountedVhd | Initialize-Disk -PartitionStyle $PartitionStyle
    $mountedVhd | New-Partition -UseMaximumSize -AssignDriveLetter |
    Format-Volume @formatParams |
    Out-Null

    Write-ProgressIndicator

    $VmDisk | Dismount-VHD

    Write-LogFunctionExit
}


function Remove-LWVHDX
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseCompatibleCmdlets", "", Justification="Not relevant on Linux")]
    [Cmdletbinding()]
    Param (
        #Path to reference VHD
        [Parameter(Mandatory = $true)]
        [string]$VhdxPath
    )

    Write-LogFunctionEntry

    $VmDisk = Get-VHD -Path $VhdxPath -ErrorAction SilentlyContinue
    if (-not $VmDisk)
    {
        Write-ScreenInfo -Message "VHDX '$VhdxPath' does not exist, cannot remove it" -Type Warning
    }
    else
    {
        $VmDisk | Remove-Item
        Write-PSFMessage "VHDX '$($vmDisk.Path)' removed"
    }

    Write-LogFunctionExit
}


function Checkpoint-LWHypervVM
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseCompatibleCmdlets", "", Justification="Not relevant on Linux")]
    [Cmdletbinding()]
    Param (
        [Parameter(Mandatory)]
        [string[]]$ComputerName,

        [Parameter(Mandatory)]
        [string]$SnapshotName
    )

    Write-LogFunctionEntry

    $step1 = {
        param ($Name, $DisableClusterCheck)
        $vm = Get-LWHypervVM -Name $Name -DisableClusterCheck $DisableClusterCheck -ErrorAction SilentlyContinue
        if ($vm.State -eq 'Running' -and -not ($vm | Get-VMSnapshot -Name $SnapshotName -ErrorAction SilentlyContinue))
        {
            $vm | Hyper-V\Suspend-VM -ErrorAction SilentlyContinue
            $vm | Hyper-V\Save-VM -ErrorAction SilentlyContinue

            Write-Verbose -Message "'$Name' was running"
            $Name
        }
    }
    $step2 = {
        param ($Name, $DisableClusterCheck)
        $vm = Get-LWHypervVM -Name $Name -DisableClusterCheck $DisableClusterCheck -ErrorAction SilentlyContinue
        if (-not ($vm | Get-VMSnapshot -Name $SnapshotName -ErrorAction SilentlyContinue))
        {
            $vm | Hyper-V\Checkpoint-VM -SnapshotName $SnapshotName
        }
        else
        {
            Write-Error "A snapshot with the name '$SnapshotName' already exists for machine '$Name'"
        }
    }
    $step3 = {
        param ($Name, $RunningMachines, $DisableClusterCheck)
        if ($Name -in $RunningMachines)
        {
            Write-Verbose -Message "Machine '$Name' was running, starting it."
            Get-LWHypervVM -Name $Name -DisableClusterCheck $DisableClusterCheck -ErrorAction SilentlyContinue | Hyper-V\Start-VM -ErrorAction SilentlyContinue
        }
        else
        {
            Write-Verbose -Message "Machine '$Name' was NOT running."
        }
    }

    $pool = New-RunspacePool -ThrottleLimit 20 -Variable (Get-Variable -Name SnapshotName) -Function (Get-Command Get-LWHypervVM)

    $jobsStep1 = foreach ($Name in $ComputerName)
    {
        Start-RunspaceJob -RunspacePool $pool -ScriptBlock $step1 -Argument $Name,(Get-LabConfigurationItem -Name DoNotAddVmsToCluster -Default $false)
    }

    $runningMachines = $jobsStep1 | Receive-RunspaceJob

    $jobsStep2 = foreach ($Name in $ComputerName)
    {
        Start-RunspaceJob -RunspacePool $pool -ScriptBlock $step2 -Argument $Name,(Get-LabConfigurationItem -Name DoNotAddVmsToCluster -Default $false)
    }

    [void] ($jobsStep2 | Wait-RunspaceJob)

    $jobsStep3 = foreach ($Name in $ComputerName)
    {
        Start-RunspaceJob -RunspacePool $pool -ScriptBlock $step3 -Argument $Name, $runningMachines,(Get-LabConfigurationItem -Name DoNotAddVmsToCluster -Default $false)
    }

    [void] ($jobsStep3 | Wait-RunspaceJob)

    $pool | Remove-RunspacePool

    Write-LogFunctionExit
}


function Dismount-LWIsoImage
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseCompatibleCmdlets", "", Justification="Not relevant on Linux")]
    param(
        [Parameter(Mandatory, Position = 0)]
        [string[]]$ComputerName
    )

    $machines = Get-LabVM -ComputerName $ComputerName

    foreach ($machine in $machines)
    {
        $vm = Get-LWHypervVM -Name $machine.ResourceName -ErrorAction SilentlyContinue
        if ($machine.OperatingSystem.Version -ge [System.Version]'6.2')
        {
            Write-PSFMessage -Message "Removing DVD drive for machine '$machine'"
            $vm | Get-VMDvdDrive | Remove-VMDvdDrive
        }
        else
        {
            Write-PSFMessage -Message "Setting DVD drive for machine '$machine' to null"
            $vm | Get-VMDvdDrive | Set-VMDvdDrive -Path $null
        }
    }
}


function Enable-LWHypervVMRemoting
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseCompatibleCmdlets", "", Justification="Not relevant on Linux")]
    param(
        [Parameter(Mandatory, Position = 0)]
        [string[]]$ComputerName
    )

    $machines = Get-LabVM -ComputerName $ComputerName

    $script = {
        param ($DomainName, $UserName, $Password)

        $RegPath = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon'

        Set-ItemProperty -Path $RegPath -Name AutoAdminLogon -Value 1 -ErrorAction SilentlyContinue
        Set-ItemProperty -Path $RegPath -Name DefaultUserName -Value $UserName -ErrorAction SilentlyContinue
        Set-ItemProperty -Path $RegPath -Name DefaultPassword -Value $Password -ErrorAction SilentlyContinue
        Set-ItemProperty -Path $RegPath -Name DefaultDomainName -Value $DomainName -ErrorAction SilentlyContinue

        Enable-WSManCredSSP -Role Server -Force | Out-Null
    }

    foreach ($machine in $machines)
    {
        $cred = $machine.GetCredential((Get-Lab))
        try
        {
            Invoke-LabCommand -ComputerName $machine -ActivityName SetLabVMRemoting -ScriptBlock $script -DoNotUseCredSsp -NoDisplay  `
            -ArgumentList $machine.DomainName, $cred.UserName, $cred.GetNetworkCredential().Password -ErrorAction Stop
        }
        catch
        {
            Connect-WSMan -ComputerName $machine -Credential $cred
            Set-Item -Path "WSMan:\$machine\Service\Auth\CredSSP" -Value $true
            Disconnect-WSMan -ComputerName $machine
        }
    }
}


function Get-LWHypervVM
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseCompatibleCmdlets", "", Justification = "Not relevant on Linux")]
    [CmdletBinding()]
    Param
    (
        [Parameter()]
        [string[]]
        $Name,

        [Parameter()]
        [bool]
        $DisableClusterCheck = (Get-LabConfigurationItem -Name DisableClusterCheck -Default $false),

        [switch]
        $NoError
    )

    Write-LogFunctionEntry

    $param = @{
        ErrorAction = 'SilentlyContinue'
    }

    if ($Name.Count -gt 0)
    {        
        $param['Name'] = $Name
    }

    [object[]]$vm = Hyper-V\Get-VM @param
    $vm = $vm | Sort-Object -Unique -Property Name

    if ($Name.Count -gt 0 -and $vm.Count -eq $Name.Count)
    {
        return $vm
    }

    if (-not $script:clusterDetected -and (Get-Command -Name Get-Cluster -Module FailoverClusters -CommandType Cmdlet -ErrorAction SilentlyContinue)) { $script:clusterDetected = Get-Cluster -ErrorAction SilentlyContinue -WarningAction SilentlyContinue}

    if (-not $DisableClusterCheck -and $script:clusterDetected)
    {
        $vm += Get-ClusterResource | Where-Object -Property ResourceType -eq 'Virtual Machine' | Get-VM
        if ($Name.Count -gt 0)
        {
            $vm = $vm | Where Name -in $Name
        }
    }

    # In case VM was in cluster and has now been added a second time
    $vm = $vm | Sort-Object -Unique -Property Name

    if (-not $NoError.IsPresent -and $Name.Count -gt 0 -and -not $vm)
    {
        Write-Error -Message "No virtual machine $Name found"
        return
    }

    if ($vm.Count -eq 0) { return } # Get-VMNetworkAdapter does not take kindly to $null
    
    $vm

    Write-LogFunctionExit
}


function Get-LWHypervVMDescription
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseCompatibleCmdlets", "", Justification="Not relevant on Linux")]
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$ComputerName
    )

    Write-LogFunctionEntry
    
    $notePath = Join-Path -Path (Get-Lab).LabPath -ChildPath "$ComputerName.xml"
    $type = Get-Type -GenericType AutomatedLab.DictionaryXmlStore -T string, string

    if (-not (Test-Path $notePath))
    {
        # Old labs still use the previous, slow method
        $vm = Get-LWHypervVM -Name $ComputerName -ErrorAction SilentlyContinue
        if (-not $vm)
        {
            return
        }

        $prefix = '#AL<#'
        $suffix = '#>AL#'
        $pattern = '{0}(?<ALNotes>[\s\S]+){1}' -f [regex]::Escape($prefix), [regex]::Escape($suffix)

        $notes = if ($vm.Notes -match $pattern) {
            $Matches.ALNotes
        }
        else {
            $vm.Notes
        }

        try
        {
            $dictionary = New-Object $type
            $importMethodInfo = $type.GetMethod('ImportFromString', [System.Reflection.BindingFlags]::Public -bor [System.Reflection.BindingFlags]::Static)
            $dictionary = $importMethodInfo.Invoke($null, $notes.Trim())
            return $dictionary
        }
        catch
        {
            Write-ScreenInfo -Message "The notes field of the virtual machine '$ComputerName' could not be read as XML" -Type Warning
            return
        }
    }

    $dictionary = New-Object $type
    try
    {
        $importMethodInfo = $type.GetMethod('Import', [System.Reflection.BindingFlags]::Public -bor [System.Reflection.BindingFlags]::Static)
        $dictionary = $importMethodInfo.Invoke($null, $notePath)
        $dictionary
    }
    catch
    {
        Write-ScreenInfo -Message "The notes field of the virtual machine '$ComputerName' could not be read as XML" -Type Warning
    }

    Write-LogFunctionExit
}


function Get-LWHypervVMSnapshot
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseCompatibleCmdlets", "", Justification="Not relevant on Linux")]
    [Cmdletbinding()]
    Param
    (
        [string[]]$VMName,

        [string]$Name
    )

    Write-LogFunctionEntry

    (Hyper-V\Get-VMSnapshot @PSBoundParameters).ForEach({
            [AutomatedLab.Snapshot]::new($_.Name, $_.VMName, $_.CreationTime)
    })

    Write-LogFunctionExit
}


function Get-LWHypervVMStatus
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseCompatibleCmdlets", "", Justification="Not relevant on Linux")]
    param (
        [Parameter(Mandatory)]
        [string[]]$ComputerName
    )

    Write-LogFunctionEntry

    $result = @{ }
    $vms = Get-LWHypervVM -Name $ComputerName -ErrorAction SilentlyContinue
    $vmTable = @{ }
    Get-LabVm -IncludeLinux | Where-Object FriendlyName -in $ComputerName | ForEach-Object {$vmTable[$_.FriendlyName] = $_.Name}

    foreach ($vm in $vms)
    {
        $vmName = if ($vmTable[$vm.Name]) {$vmTable[$vm.Name]} else {$vm.Name}
        if ($vm.State -eq 'Running')
        {
            $result.Add($vmName, 'Started')
        }
        elseif ($vm.State -eq 'Off')
        {
            $result.Add($vmName, 'Stopped')
        }
        else
        {
            $result.Add($vmName, 'Unknown')
        }
    }

    $result

    Write-LogFunctionExit
}


function Mount-LWIsoImage
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseCompatibleCmdlets", "", Justification="Not relevant on Linux")]
    param(
        [Parameter(Mandatory, Position = 0)]
        [string[]]$ComputerName,

        [Parameter(Mandatory, Position = 1)]
        [string]$IsoPath,

        [switch]$PassThru
    )

    if (-not (Test-Path -Path $IsoPath -PathType Leaf))
    {
        Write-Error "The path '$IsoPath' could not be found or is pointing to a folder"
        return
    }

    $IsoPath = (Resolve-Path -Path $IsoPath).Path
    $machines = Get-LabVM -ComputerName $ComputerName

    foreach ($machine in $machines)
    {
        Write-PSFMessage -Message "Adding DVD drive '$IsoPath' to machine '$machine'"
        $start = (Get-Date)
        $done = $false
        $delayBeforeCheck = 5, 10, 15, 30, 45, 60
        $delayIndex = 0

        $dvdDrivesBefore = Invoke-LabCommand -ComputerName $machine -ScriptBlock {
            Get-WmiObject -Class Win32_LogicalDisk -Filter 'DriveType = 5 AND FileSystem LIKE "%"' | Select-Object -ExpandProperty DeviceID
        } -PassThru -NoDisplay

        #this is required as Compare-Object cannot work with a null object
        if (-not $dvdDrivesBefore) { $dvdDrivesBefore = @() }

        while ((-not $done) -and ($delayIndex -le $delayBeforeCheck.Length))
        {
            try
            {
                $vm = Get-LWHypervVM -Name $machine.ResourceName
                if ($machine.OperatingSystem.Version -ge '6.2')
                {
                    $drive = $vm | Add-VMDvdDrive -Path $IsoPath -ErrorAction Stop -Passthru -AllowUnverifiedPaths
                }
                else
                {
                    if (-not ($vm | Get-VMDvdDrive))
                    {
                        throw "No DVD drive exist for machine '$machine'. Machine is generation 1 and DVD drive needs to be crate in advance (during creation of the machine). Cannot continue."
                    }
                    $drive = $vm | Set-VMDvdDrive -Path $IsoPath -ErrorAction Stop -Passthru -AllowUnverifiedPaths
                }

                Start-Sleep -Seconds $delayBeforeCheck[$delayIndex]

                if (($vm | Get-VMDvdDrive).Path -contains $IsoPath)
                {
                    $done = $true
                }
                else
                {
                    Write-ScreenInfo -Message "DVD drive '$IsoPath' was NOT successfully added to machine '$machine'. Retrying." -Type Error
                    $delayIndex++
                }
            }
            catch
            {
                Write-ScreenInfo -Message "Could not add DVD drive '$IsoPath' to machine '$machine'. Retrying." -Type Warning
                Start-Sleep -Seconds $delayBeforeCheck[$delayIndex]
            }
        }

        $dvdDrivesAfter = Invoke-LabCommand -ComputerName $machine -ScriptBlock {
            Get-WmiObject -Class Win32_LogicalDisk -Filter 'DriveType = 5 AND FileSystem LIKE "%"' | Select-Object -ExpandProperty DeviceID
        } -PassThru -NoDisplay

        $driveLetter = (Compare-Object -ReferenceObject $dvdDrivesBefore -DifferenceObject $dvdDrivesAfter).InputObject
        $drive | Add-Member -Name DriveLetter -MemberType NoteProperty -Value $driveLetter
        $drive | Add-Member -Name InternalComputerName -MemberType NoteProperty -Value $machine.Name

        if ($PassThru) { $drive }

        if (-not $done)
        {
            throw "Could not add DVD drive '$IsoPath' to machine '$machine' after repeated attempts."
        }
    }
}


function New-LWHypervVM
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseCompatibleCmdlets", "", Justification = "Not relevant on Linux")]
    [Cmdletbinding()]
    Param (
        [Parameter(Mandatory)]
        [AutomatedLab.Machine]$Machine
    )

    $PSBoundParameters.Add('ProgressIndicator', 1) #enables progress indicator
    if ($Machine.SkipDeployment) { return }

    Write-LogFunctionEntry

    $script:lab = Get-Lab

    if (Get-LWHypervVM -Name $Machine.ResourceName -ErrorAction SilentlyContinue)
    {
        Write-ProgressIndicatorEnd
        Write-ScreenInfo -Message "The machine '$Machine' does already exist" -Type Warning
        return $false
    }

    if ($PSDefaultParameterValues.ContainsKey('*:IsKickstart')) { $PSDefaultParameterValues.Remove('*:IsKickstart') }
    if ($PSDefaultParameterValues.ContainsKey('*:IsAutoYast')) { $PSDefaultParameterValues.Remove('*:IsAutoYast') }
    if ($PSDefaultParameterValues.ContainsKey('*:IsCloudInit')) { $PSDefaultParameterValues.Remove('*:IsCloudInit') }

    if ($Machine.OperatingSystemType -eq 'Linux' -and $Machine.LinuxType -eq 'RedHat')
    {
        $PSDefaultParameterValues['*:IsKickstart'] = $true
    }
    if ($Machine.OperatingSystemType -eq 'Linux' -and $Machine.LinuxType -eq 'Suse')
    {
        $PSDefaultParameterValues['*:IsAutoYast'] = $true
    }
    if ($Machine.OperatingSystemType -eq 'Linux' -and $Machine.LinuxType -eq 'Ubuntu')
    {
        $PSDefaultParameterValues['*:IsCloudInit'] = $true
    }

    Write-PSFMessage "Creating machine with the name '$($Machine.ResourceName)' in the path '$VmPath'"

    #region Unattend XML settings
    if (-not $Machine.ProductKey)
    {
        $Machine.ProductKey = $Machine.OperatingSystem.ProductKey
    }

    $unattendContent = $Machine.UnattendedXmlContent
    if ($Machine.LinuxType -eq 'Suse' -and $Machine.OperatingSystem.OperatingSystemName -match 'Leap') {
        $unattendContent = $unattendContent -replace 'SUSEVERSION', "$($Machine.OperatingSystem.Version.Major).$($Machine.OperatingSystem.Version.Minor)"
    }

    Import-UnattendedContent -Content $unattendContent

    # Ensure package selection works
    if ($Machine.LinuxType -eq 'Suse' -and $Machine.OperatingSystem.OperatingSystemName -match 'Tumbleweed') {
        $nsm = [System.Xml.XmlNamespaceManager]::new((Get-UnattendedContent).NameTable)
        $nsm.AddNamespace('un', "http://www.suse.com/1.0/yast2ns")
        $nsm.AddNamespace('config', "http://www.suse.com/1.0/configns" )
        $addOnNode = (Get-UnattendedContent).SelectSingleNode('/un:profile/un:add-on/un:add_on_others', $nsm)
        $addOnNode.RemoveAll()

        # Restore attribute after clearing the node
        $listAttr = (Get-UnattendedContent).CreateAttribute('t')
        $listAttr.InnerText = 'list'
        $null = $addOnNode.Attributes.Append($listAttr)

        $listNodeUpdate = (Get-UnattendedContent).CreateElement('listentry', $nsm.LookupNamespace('un'))
        $mapAttr = (Get-UnattendedContent).CreateAttribute('t')
        $mapAttr.InnerText = 'map'
        $aliasNode = (Get-UnattendedContent).CreateElement('alias', $nsm.LookupNamespace('un'))
        $aliasNode.InnerText = 'repo-update'
        $mediaUrlNode = (Get-UnattendedContent).CreateElement('media_url', $nsm.LookupNamespace('un')) 
        $mediaUrlNode.InnerText = 'http://download.opensuse.org/update/tumbleweed/'
        $nameNode = (Get-UnattendedContent).CreateElement('name', $nsm.LookupNamespace('un'))
        $nameNode.InnerText = 'Update'
        $priorityNode = (Get-UnattendedContent).CreateElement('priority', $nsm.LookupNamespace('un'))
        $priorityNode.InnerText = '1'
        $null = $listNodeUpdate.AppendChild($aliasNode)
        $null = $listNodeUpdate.AppendChild($mediaUrlNode)
        $null = $listNodeUpdate.AppendChild($nameNode)
        $null = $listNodeUpdate.AppendChild($priorityNode)
        $null = $listNodeUpdate.Attributes.Append($mapAttr)
        $null = $addOnNode.AppendChild($listNodeUpdate)


        $listNodeNonOss = (Get-UnattendedContent).CreateElement('listentry', $nsm.LookupNamespace('un'))
        $mapAttr = (Get-UnattendedContent).CreateAttribute('t')
        $mapAttr.InnerText = 'map'
        $aliasNode = (Get-UnattendedContent).CreateElement('alias', $nsm.LookupNamespace('un'))
        $aliasNode.InnerText = 'repo-update'
        $mediaUrlNode = (Get-UnattendedContent).CreateElement('media_url', $nsm.LookupNamespace('un')) 
        $mediaUrlNode.InnerText = 'http://download.opensuse.org/tumbleweed/repo/non-oss/'
        $nameNode = (Get-UnattendedContent).CreateElement('name', $nsm.LookupNamespace('un'))
        $nameNode.InnerText = 'Update'
        $priorityNode = (Get-UnattendedContent).CreateElement('priority', $nsm.LookupNamespace('un'))
        $priorityNode.InnerText = '2'
        $null = $listNodeNonOss.AppendChild($aliasNode)
        $null = $listNodeNonOss.AppendChild($mediaUrlNode)
        $null = $listNodeNonOss.AppendChild($nameNode)
        $null = $listNodeNonOss.AppendChild($priorityNode)
        $null = $listNodeNonOss.Attributes.Append($mapAttr)
        $null = $addOnNode.AppendChild($listNodeNonOss)
    }
    #endregion

    #region Ubuntu Desktop configuration
    # ref: https://github.com/canonical/autoinstall-desktop
    if ($Machine.LinuxType -eq 'Ubuntu' -and $Machine.OperatingSystem.Edition -eq 'Desktop') {
        Set-UnattendedPackage -Package 'ubuntu-desktop'
        Set-UnattendedPackage -Package firefox -IsSnap $true
        if ($Machine.OperatingSystem.Version.Major -eq 20) {
            Set-UnattendedPackage -Package gnome-3-38-2004 -IsSnap $true
            Add-UnattendedPreinstallationCommand -Description 'Enable Hardware Experience' -Command "echo 'linux-generic-hwe-20.04' > /run/kernel-meta-package"
        } elseif ($Machine.OperatingSystem.Version.Major -eq 22) {
            Set-UnattendedPackage -Package gnome-42-2204 -IsSnap $true # Same package is installed as verified on installed OS
            Add-UnattendedPreinstallationCommand -Description 'Enable Hardware Experience' -Command "echo 'linux-generic-hwe-22.04' > /run/kernel-meta-package"
        } else {
            Set-UnattendedPackage -Package gnome-42-2204 -IsSnap $true # Same package is installed as verified on installed OS
            Add-UnattendedPreinstallationCommand -Description 'Enable Hardware Experience' -Command "echo 'linux-generic-hwe-24.04' > /run/kernel-meta-package"
        }
        Set-UnattendedPackage -Package gtk-common-themes -IsSnap $true
        Set-UnattendedPackage -Package snap-store -IsSnap $true
        Set-UnattendedPackage -Package snapd-desktop-integration -IsSnap $true
        Add-UnattendedSynchronousCommand -Description 'Configure grub' -Command "sed -i /etc/default/grub -e 's/GRUB_CMDLINE_LINUX_DEFAULT=`".*/GRUB_CMDLINE_LINUX_DEFAULT=`"quiet splash`"/'"
        Add-UnattendedSynchronousCommand -Description 'Update grub' -Command 'update-grub'
        Add-UnattendedSynchronousCommand -Description 'Keep cloud init' -Command 'apt-get install -y cloud-init'
        Add-UnattendedSynchronousCommand -Description 'Remove unused packages' -Command 'apt-get autoremove -y'
    }
    #endregion

    #region network adapter settings
    $macAddressPrefix = Get-LabConfigurationItem -Name MacAddressPrefix
    $macAddressesInUse = @(Get-LWHypervVM | Get-VMNetworkAdapter | Select-Object -ExpandProperty MacAddress)
    $macAddressesInUse += (Get-LabVm -IncludeLinux).NetworkAdapters.MacAddress

    $macIdx = 0
    $prefixlength = 12 - $macAddressPrefix.Length
    while ("$macAddressPrefix{0:X$prefixLength}" -f $macIdx -in $macAddressesInUse) { $macIdx++ }

    $type = Get-Type -GenericType AutomatedLab.ListXmlStore -T AutomatedLab.NetworkAdapter
    $adapters = New-Object $type
    $Machine.NetworkAdapters | ForEach-Object { $adapters.Add($_) }

    if ($Machine.IsDomainJoined)
    {
        #move the adapter that connects the machine to the domain to the top
        $dc = Get-LabVM -Role RootDC, FirstChildDC | Where-Object { $_.DomainName -eq $Machine.DomainName }

        if ($dc)
        {
            #the first adapter that has an IP address in the same IP range as the RootDC or FirstChildDC in the same domain will be used on top of
            #the network ordering
            $domainAdapter = $adapters | Where-Object { $_.Ipv4Address[0] } |
            Where-Object { [AutomatedLab.IPNetwork]::Contains($_.Ipv4Address[0], $dc.IpAddress[0]) } |
            Select-Object -First 1

            if ($domainAdapter)
            {
                $adapters.Remove($domainAdapter)
                $adapters.Insert(0, $domainAdapter)
            }
        }
    }

    $adapterCount = 0
    foreach ($adapter in $adapters)
    {
        $ipSettings = @{}
        $openSuseLinuxRcNetwork = [System.Text.StringBuilder]::new()
        $null = $openSuseLinuxRcNetwork.Append("ifcfg=`"eth$($adapterCount)`"=")

        $prefixlength = 12 - $macAddressPrefix.Length
        $mac = "$macAddressPrefix{0:X$prefixLength}" -f $macIdx++

        if (-not $adapter.MacAddress)
        {
            $adapter.MacAddress = $mac
        }
        
        #$ipSettings.Add('MacAddress', $adapter.MacAddress)
        $macWithDash = '{0}-{1}-{2}-{3}-{4}-{5}' -f (Get-StringSection -SectionSize 2 -String $adapter.MacAddress)

        $ipSettings.Add('InterfaceName', $macWithDash)
        $ipSettings.Add('IpAddresses', @())
        if ($adapter.Ipv4Address.Count -ge 1)
        {
            foreach ($ipv4Address in $adapter.Ipv4Address)
            {
                $ipSettings.IpAddresses += "$($ipv4Address.IpAddress)/$($ipv4Address.Cidr)"
            }
        }
        if ($adapter.Ipv6Address.Count -ge 1)
        {
            foreach ($ipv6Address in $adapter.Ipv6Address)
            {
                $ipSettings.IpAddresses += "$($ipv6Address.IpAddress)/$($ipv6Address.Cidr)"
            }
        }

        $ipSettings.Add('Gateways', ($adapter.Ipv4Gateway + $adapter.Ipv6Gateway))
        $ipSettings.Add('DNSServers', ($adapter.Ipv4DnsServers + $adapter.Ipv6DnsServers))
        
        $null = $openSuseLinuxRcNetwork.Append($ipSettings.IpAddresses -join ' ')
        $null = $openSuseLinuxRcNetwork.Append(' ')
        $null = $openSuseLinuxRcNetwork.Append($ipSettings.Gateways -join ' ')
        $null = $openSuseLinuxRcNetwork.Append(' ')
        $null = $openSuseLinuxRcNetwork.Append($ipSettings.DNSServers -join ' ')

        if (-not $Machine.IsDomainJoined -and (-not $adapter.ConnectionSpecificDNSSuffix))
        {
            $rootDomainName = Get-LabVM -Role RootDC | Select-Object -First 1 | Select-Object -ExpandProperty DomainName
            $ipSettings.Add('DnsDomain', $rootDomainName)
            $null = $openSuseLinuxRcNetwork.Append(" $rootDomainName")
        }

        if ($adapter.ConnectionSpecificDNSSuffix)
        {
            $ipSettings.Add('DnsDomain', $adapter.ConnectionSpecificDNSSuffix)
            $null = $openSuseLinuxRcNetwork.Append(" $($adapter.ConnectionSpecificDNSSuffix)")
        }

        $ipSettings.Add('UseDomainNameDevolution', (([string]($adapter.AppendParentSuffixes)) = 'true'))
        if ($adapter.AppendDNSSuffixes)
        {
            $ipSettings.Add('DNSSuffixSearchOrder', $adapter.AppendDNSSuffixes -join ',')
            $null = $openSuseLinuxRcNetwork.Append(" $($adapter.AppendDNSSuffixes -join ' ')")
        }

        $ipSettings.Add('EnableAdapterDomainNameRegistration', ([string]($adapter.DnsSuffixInDnsRegistration)).ToLower())
        $ipSettings.Add('DisableDynamicUpdate', ([string](-not $adapter.RegisterInDNS)).ToLower())

        if ($machine.OperatingSystemType -eq 'Linux' -and $machine.LinuxType -eq 'RedHat')
        {
            $ipSettings.Add('IsKickstart', $true)
        }
        if ($machine.OperatingSystemType -eq 'Linux' -and $machine.LinuxType -eq 'Suse')
        {
            $ipSettings.Add('IsAutoYast', $true)
        }
        if ($machine.OperatingSystemType -eq 'Linux' -and $machine.LinuxType -eq 'Ubuntu')
        {
            $ipSettings.Add('IsCloudInit', $true)
        }

        switch ($Adapter.NetbiosOptions)
        {
            'Default' { $ipSettings.Add('NetBIOSOptions', '0') }
            'Enabled' { $ipSettings.Add('NetBIOSOptions', '1') }
            'Disabled' { $ipSettings.Add('NetBIOSOptions', '2') }
        }

        Add-UnattendedNetworkAdapter @ipSettings
        $adapterCount++
    }

    $Machine.NetworkAdapters = $adapters

    if ($Machine.OperatingSystemType -eq 'Windows')
    {
        Add-UnattendedRenameNetworkAdapters
    }
    #endregion network adapter settings

    Set-UnattendedComputerName -ComputerName $Machine.Name
    Set-UnattendedAdministratorName -Name $Machine.InstallationUser.UserName
    Set-UnattendedAdministratorPassword -Password $Machine.InstallationUser.Password

    if ($Machine.ProductKey)
    {
        Set-UnattendedProductKey -ProductKey $Machine.ProductKey
    }

    if ($Machine.UserLocale)
    {
        Set-UnattendedUserLocale -UserLocale $Machine.UserLocale
    }

    #if the time zone is specified we use it, otherwise we take the timezone from the host machine
    if ($Machine.TimeZone)
    {
        Set-UnattendedTimeZone -TimeZone $Machine.TimeZone
    }
    else
    {
        Set-UnattendedTimeZone -TimeZone ([System.TimeZoneInfo]::Local.Id)
    }

    #if domain-joined and not a DC
    if ($Machine.IsDomainJoined -eq $true -and -not ($Machine.Roles.Name -contains 'RootDC' -or $Machine.Roles.Name -contains 'FirstChildDC' -or $Machine.Roles.Name -contains 'DC'))
    {
        Set-UnattendedAutoLogon -DomainName $Machine.DomainName -Username $Machine.InstallationUser.Username -Password $Machine.InstallationUser.Password
    }
    else
    {
        Set-UnattendedAutoLogon -DomainName $Machine.Name -Username $Machine.InstallationUser.Username -Password $Machine.InstallationUser.Password
    }

    $disableWindowsDefender = Get-LabConfigurationItem -Name DisableWindowsDefender
    if (-not $disableWindowsDefender)
    {
        Set-UnattendedAntiMalware -Enabled $false
    }

    $setLocalIntranetSites = Get-LabConfigurationItem -Name SetLocalIntranetSites
    if ($setLocalIntranetSites -ne 'None' -or $null -ne $setLocalIntranetSites)
    {
        if ($setLocalIntranetSites -eq 'All')
        {
            $localIntranetSites = $lab.Domains
        }
        elseif ($setLocalIntranetSites -eq 'Forest' -and $Machine.DomainName)
        {
            $forest = $lab.GetParentDomain($Machine.DomainName)
            $localIntranetSites = $lab.Domains | Where-Object { $lab.GetParentDomain($_) -eq $forest }
        }
        elseif ($setLocalIntranetSites -eq 'Domain' -and $Machine.DomainName)
        {
            $localIntranetSites = $Machine.DomainName
        }

        $localIntranetSites = $localIntranetSites | ForEach-Object {
            "http://$($_)"
            "https://$($_)"
        }

        #removed the call to Set-LocalIntranetSites as setting the local intranet zone in the unattended file does not work due to bugs in Windows
        #Set-LocalIntranetSites -Values $localIntranetSites
    }

    Set-UnattendedFirewallState -State $Machine.EnableWindowsFirewall

    if ($Machine.LinuxType -eq 'Suse') {
        try {
            $repoContent = (Invoke-RestMethod -Method Get -Uri "https://packages.microsoft.com/config/rhel/$Version/prod.repo" -ErrorAction Stop) -split "`n"
        }
        catch { }

        $pwshRelease = ((Invoke-RestMethod -Uri 'https://api.github.com/repos/PowerShell/PowerShell/releases/latest' -ErrorAction SilentlyContinue).assets | Where-Object Name -match 'rh\.x86_64\.rpm').browser_download_url
        if (-not $pwshRelease) {
            $pwshRelease = 'https://github.com/PowerShell/PowerShell/releases/download/v7.5.2/powershell-7.5.2-1.rh.x86_64.rpm'
        }

        Add-UnattendedSynchronousCommand -Command "sudo zypper update -y && sudo zypper install -y libicu libopenssl3`nsudo rpm -i --nodeps $pwshRelease`necho `"Subsystem powershell /usr/bin/pwsh -sshs -NoLogo`" >> /etc/ssh/sshd_config`nsystemctl restart sshd`n" -Description 'Install PowerShell'
    }
    
    if (-not [string]::IsNullOrEmpty($Machine.SshPublicKey) -and $Machine.LinuxType -in 'Ubuntu', 'Suse')
    {
        Add-UnattendedSshPublicKey -PublicKey $Machine.SshPublicKey
    }
    elseif ($Machine.OperatingSystemType -eq 'Linux' -and -not [string]::IsNullOrEmpty($Machine.SshPublicKey))
    {
        $command = @"
mkdir -p /root/.ssh
mkdir -p /home/$($Machine.InstallationUser.UserName)/.ssh
echo "$($Machine.SshPublicKey)" > /root/.ssh/authorized_keys
echo "$($Machine.SshPublicKey)" > /home/$($Machine.InstallationUser.UserName)/.ssh/authorized_keys
chown -R root:root /root/.ssh
chown -R $($Machine.InstallationUser.UserName):$($Machine.InstallationUser.UserName) /home/$($Machine.InstallationUser.UserName)/.ssh
chmod 700 /root/.ssh && chmod 600 /root/.ssh/authorized_keys
chmod 700 /home/$($Machine.InstallationUser.UserName)/.ssh && chmod 600 /home/$($Machine.InstallationUser.UserName)/.ssh/authorized_keys
sed -i 's|[#]*GSSAPIAuthentication yes|GSSAPIAuthentication yes|g' /etc/ssh/sshd_config
sed -i 's|[#]*PasswordAuthentication yes|PasswordAuthentication no|g' /etc/ssh/sshd_config
sed -i 's|[#]*PubkeyAuthentication yes|PubkeyAuthentication yes|g' /etc/ssh/sshd_config
restorecon -R /$($Machine.InstallationUser.UserName)/.ssh/
restorecon -R /root/.ssh/
"@
        Add-UnattendedSynchronousCommand -Command $command -Description 'SSH'
    }

    if ($Machine.Roles.Name -contains 'RootDC' -or
        $Machine.Roles.Name -contains 'FirstChildDC' -or
        $Machine.Roles.Name -contains 'DC')
    {
        #machine will not be added to domain or workgroup
    }
    else
    {
        if (-not [string]::IsNullOrEmpty($Machine.WorkgroupName))
        {
            Set-UnattendedWorkgroup -WorkgroupName $Machine.WorkgroupName
        }

        if (-not [string]::IsNullOrEmpty($Machine.DomainName))
        {
            $domain = $lab.Domains | Where-Object Name -eq $Machine.DomainName

            $parameters = @{
                DomainName = $Machine.DomainName
                Username   = $domain.Administrator.UserName
                Password   = $domain.Administrator.Password
            }
            if ($Machine.OrganizationalUnit) {
                $parameters['OrganizationalUnit'] = $machine.OrganizationalUnit
            }

            Set-UnattendedDomain @parameters

            if ($Machine.OperatingSystemType -eq 'Linux' -and $Machine.LinuxType -ne 'Ubuntu')
            {
                if ($Machine.LinuxType -eq 'Suse')
                {
                    Set-UnattendedPackage -Package sssd, samba
                }

                $sudoParam = @{
                    Command     = "sed -i '/^%wheel.*/a %$($Machine.DomainName.ToUpper())\\\\domain\\ admins ALL=(ALL) NOPASSWD: ALL' /etc/sudoers"
                    Description = 'Enable domain admin as sudoer without password'
                }

                Add-UnattendedSynchronousCommand @sudoParam
                [System.Collections.Generic.List[string]] $commands = @(
                    "mkdir -p /home/$($domain.Administrator.UserName)@$($Machine.DomainName)/.ssh"
                    "chown -R $($domain.Administrator.UserName)@$($Machine.DomainName):domain\ users@$($Machine.DomainName) /home/$($domain.Administrator.UserName)@$($Machine.DomainName)/.ssh" 
                    "chmod 700 /home/$($domain.Administrator.UserName)@$($Machine.DomainName)/.ssh && chmod 600 /home/$($domain.Administrator.UserName)@$($Machine.DomainName)/.ssh/authorized_keys"
                    "echo `"$($Machine.SshPublicKey)`" > /home/$($domain.Administrator.UserName)@$($Machine.DomainName)/.ssh/authorized_keys"
                    "restorecon -R /home/$($domain.Administrator.UserName)@$($Machine.DomainName)/.ssh/"
                )

                if (-not [string]::IsNullOrEmpty($Machine.SshPublicKey))
                {
                    $command = @"
mkdir -p /home/$($Machine.InstallationUser.UserName.ToLower())@$($Machine.DomainName.ToLower())/.ssh
chown -R "$($Machine.InstallationUser.UserName)@$($Machine.DomainName):domain users@$($Machine.DomainName)" /home/$($Machine.InstallationUser.UserName.ToLower())@$($Machine.DomainName.ToLower())/.ssh
chmod 700 /home/$($Machine.InstallationUser.UserName.ToLower())@$($Machine.DomainName.ToLower())/.ssh && chmod 600 /home/$($Machine.InstallationUser.UserName.ToLower())@$($Machine.DomainName.ToLower())/.ssh/authorized_keys
echo "$($Machine.SshPublicKey)" > /home/$($Machine.InstallationUser.UserName.ToLower())@$($Machine.DomainName.ToLower())/.ssh/authorized_keys
restorecon -R /$($domain.Administrator.UserName)@$($Machine.DomainName)/.ssh/
"@
                    Add-UnattendedSynchronousCommand -Command $command -Description 'SSH'
                }
            }
            elseif ($Machine.OperatingSystemType -eq 'Linux' -and $Machine.LinuxType -eq 'Ubuntu')
            {
                Write-UnattendedFile -Content $Machine.SshPublicKey -DestinationPath "/home/$($domain.Administrator.UserName.ToLower())@$($Machine.DomainName)/.ssh/authorized_keys"
                Write-UnattendedFile -Content @"
#!/bin/bash
mkdir -p /home/$($domain.Administrator.UserName.ToLower())@$($Machine.DomainName)/.ssh
chown -R $($domain.Administrator.UserName.ToLower())@$($Machine.DomainName):domain\ users@$($Machine.DomainName) /home/$($domain.Administrator.UserName.ToLower())@$($Machine.DomainName)
chmod 700 /home/$($domain.Administrator.UserName.ToLower())@$($Machine.DomainName)/.ssh && chmod 600 /home/$($domain.Administrator.UserName.ToLower())@$($Machine.DomainName)/.ssh/authorized_keys
restorecon -R /home/$($domain.Administrator.UserName.ToLower())@$($Machine.DomainName)/.ssh/
rm -rf /etc/cron.d/postconf
"@ -DestinationPath '/postconf.sh'
                Write-UnattendedFile -Content '@reboot root sleep 30; bash /postconf.sh' -DestinationPath '/etc/cron.d/10postconf'
            }
        }
    }

    #set the Generation for the VM depending on SupportGen2VMs, host OS version and VM OS version
    $hostOsVersion = [System.Environment]::OSVersion.Version

    $generation = if (Get-LabConfigurationItem -Name SupportGen2VMs)
    {
        if ($Machine.VmGeneration -ne 1 -and $hostOsVersion -ge [System.Version]6.3 -and $Machine.Gen2VmSupported)
        {
            2
        }
        else
        {
            1
        }
    }
    else
    {
        1
    }

    $vmPath = $lab.GetMachineTargetPath($Machine.ResourceName)
    $path = "$vmPath\$($Machine.ResourceName).vhdx"
    Write-PSFMessage "`tVM Disk path is '$path'"

    if (Test-Path -Path $path)
    {
        Write-ScreenInfo -Message "The disk $path does already exist. Disk cannot be created" -Type Warning
        return $false
    }

    Write-ProgressIndicator

    if ($Machine.OperatingSystemType -eq 'Linux')
    {
        $nextDriveLetter = [char[]](67..90) |
        Where-Object { (Get-CimInstance -Class Win32_LogicalDisk |
                Select-Object -ExpandProperty DeviceID) -notcontains "$($_):" } |
        Select-Object -First 1
        $systemDisk = New-Vhd -Path $path -SizeBytes ($lab.Target.ReferenceDiskSizeInGB * 1GB) -BlockSizeBytes 1MB
        $mountedOsDisk = $systemDisk | Mount-VHD -Passthru

            $mountedOsDisk | Initialize-Disk -PartitionStyle GPT
            $size = 6GB
            if ($Machine.LinuxType -in 'RedHat', 'Ubuntu')
            {
                $size = 100MB
            }
            $label = if ($Machine.LinuxType -eq 'RedHat') { 'OEMDRV' } else { 'CIDATA' }
            $unattendPartition = $mountedOsDisk | New-Partition -Size $size

            # Use a small FAT32 partition to hold AutoYAST and Kickstart configuration
            $diskpartCmd = "@
                select disk $($mountedOsDisk.DiskNumber)
                select partition $($unattendPartition.PartitionNumber)
                format quick fs=fat32 label=$label
                exit
            @"
            $diskpartCmd | diskpart.exe | Out-Null

            $unattendPartition | Set-Partition -NewDriveLetter $nextDriveLetter
            $unattendPartition = $unattendPartition | Get-Partition
            $drive = [System.IO.DriveInfo][string]$unattendPartition.DriveLetter

        if ($machine.LinuxPackageGroup )
        {
            Set-UnattendedPackage -Package $machine.LinuxPackageGroup
        }
        elseif ($machine.LinuxType -eq 'RedHat')
        {
            Set-UnattendedPackage -Package '@^server-product-environment'
        }

        # Copy Unattend-Stuff here
        if ($machine.LinuxType -eq 'RedHat')
        {
            Export-UnattendedFile -Path (Join-Path -Path $drive.RootDirectory -ChildPath ks.cfg) -Version $machine.OperatingSystem.Version.Major
            Copy-Item -Path (Join-Path -Path $drive.RootDirectory -ChildPath ks.cfg) -Destination (Join-Path -Path $script:lab.Sources.UnattendedXml.Value -ChildPath "ks_$($Machine.Name).cfg")
        }
        elseif ($Machine.LinuxType -eq 'Suse')
        {
            Export-UnattendedFile -Path (Join-Path -Path $drive.RootDirectory -ChildPath autoinst.xml)
            Export-UnattendedFile -Path (Join-Path -Path $script:lab.Sources.UnattendedXml.Value -ChildPath "autoinst_$($Machine.Name).xml")
            # Mount ISO
            $mountedIso = Mount-DiskImage -ImagePath $Machine.OperatingSystem.IsoPath -PassThru | Get-Volume
            $isoDrive = [System.IO.DriveInfo][string]$mountedIso.DriveLetter
            # Copy data
            Copy-Item -Path "$($isoDrive.RootDirectory.FullName)*" -Destination $drive.RootDirectory.FullName -Recurse -Force -PassThru |
            Where-Object IsReadOnly | Set-ItemProperty -name IsReadOnly -Value $false
            
            # Unmount ISO
            [void] (Dismount-DiskImage -ImagePath $Machine.OperatingSystem.IsoPath)

            # AutoYast XML file is not picked up properly without modifying bootloader config
            # Change grub and isolinux configuration
            $grubFile = Get-ChildItem -Recurse -Path $drive.RootDirectory.FullName -Filter 'grub.cfg'
            $isolinuxFile = Get-ChildItem -Recurse -Path $drive.RootDirectory.FullName -Filter 'isolinux.cfg'

            ($grubFile | Get-Content -Raw) -replace "splash=silent", "splash=silent textmode=1 $openSuseLinuxRcNetwork YAST_SKIP_XML_VALIDATION=1 autoyast=device:///autoinst.xml" | Set-Content -Path $grubFile.FullName
            ($isolinuxFile | Get-Content -Raw) -replace "splash=silent", "splash=silent textmode=1 $openSuseLinuxRcNetwork YAST_SKIP_XML_VALIDATION=1 autoyast=device:///autoinst.xml" | Set-Content -Path $isolinuxFile.FullName
        }
        elseif ($machine.LinuxType -eq 'Ubuntu')
        {
            Export-UnattendedFile -Path $drive.RootDirectory
            $ubuLease = '{0:d2}.{1:d2}' -f $machine.OperatingSystem.Version.Major, $machine.OperatingSystem.Version.Minor # Microsoft Repo does not use $RELEASE but version number instead.
            (Get-Content -Path (Join-Path -Path $drive.RootDirectory -ChildPath user-data)) -replace 'REPLACERELEASE', $ubuLease | Set-Content (Join-Path -Path $drive.RootDirectory -ChildPath user-data)
            (Get-Content -Path (Join-Path -Path $drive.RootDirectory -ChildPath meta-data)) -replace 'REPLACERELEASE', $ubuLease | Set-Content (Join-Path -Path $drive.RootDirectory -ChildPath meta-data)
            

            Copy-Item -Path (Join-Path -Path $drive.RootDirectory -ChildPath user-data) -Destination (Join-Path -Path $script:lab.Sources.UnattendedXml.Value -ChildPath "cloudinit_user_$($Machine.Name).yml")
            Copy-Item -Path (Join-Path -Path $drive.RootDirectory -ChildPath meta-data) -Destination (Join-Path -Path $script:lab.Sources.UnattendedXml.Value -ChildPath "cloudinit_meta_$($Machine.Name).yml")
        }

        $mountedOsDisk | Dismount-VHD

        if ($PSDefaultParameterValues.ContainsKey('*:IsKickstart')) { $PSDefaultParameterValues.Remove('*:IsKickstart') }
        if ($PSDefaultParameterValues.ContainsKey('*:IsAutoYast')) { $PSDefaultParameterValues.Remove('*:IsAutoYast') }
        if ($PSDefaultParameterValues.ContainsKey('*:CloudInit')) { $PSDefaultParameterValues.Remove('*:CloudInit') }
    }
    else
    {
        $referenceDiskPath = if ($Machine.ReferenceDiskPath) { $Machine.ReferenceDiskPath } else { $Machine.OperatingSystem.BaseDiskPath }
        $systemDisk = New-VHD -Path $path -Differencing -ParentPath $referenceDiskPath -ErrorAction Stop
        Write-PSFMessage "`tcreated differencing disk '$($systemDisk.Path)' pointing to '$ReferenceVhdxPath'"

        $mountedOsDisk = Mount-VHD -Path $path -Passthru
        try
        {
            $drive = $mountedosdisk | get-disk | Get-Partition | Get-Volume  | Where { $_.DriveLetter -and $_.FileSystemLabel -eq 'System' }

            $paths = [Collections.ArrayList]::new()
            $alcommon = Get-Module -Name AutomatedLab.Common
            $null = $paths.Add((Split-Path -Path $alcommon.ModuleBase -Parent))
            $null = foreach ($req in $alCommon.RequiredModules.Name)
            {
                $paths.Add((Split-Path -Path (Get-Module -Name $req -ListAvailable)[0].ModuleBase -Parent))
            }

            Copy-Item -Path $paths -Destination "$($drive.DriveLetter):\Program Files\WindowsPowerShell\Modules" -Recurse


            if ($Machine.InitialDscConfigurationMofPath)
            {
                $exportedModules = Get-RequiredModulesFromMOF -Path $Machine.InitialDscConfigurationMofPath
                foreach ($exportedModule in $exportedModules.GetEnumerator())
                {
                    $moduleInfo = Get-Module -ListAvailable -Name $exportedModule.Key | Where-Object Version -eq $exportedModule.Value | Select-Object -First 1
                    if (-not $moduleInfo)
                    {
                        Write-ScreenInfo -Type Warning -Message "Unable to find $($exportedModule.Key). Attempting to download from PSGallery"
                        Save-Module -Path "$($drive.DriveLetter):\Program Files\WindowsPowerShell\Modules" -Name $exportedModule.Key -RequiredVersion $exportedModule.Value -Repository PSGallery -Force -AllowPrerelease
                    }
                    else
                    {
                        $source = Get-ModuleDependency -Module $moduleInfo | Sort-Object -Unique | ForEach-Object { 
                            if ((Get-Item $_).BaseName -match '\d{1,4}\.\d{1,4}\.\d{1,4}' -and $Machine.OperatingSystem.Version -ge 10.0)
                            {
                                #parent folder contains a specific version. In order to copy the module right, the parent of this parent is required
                                Split-Path -Path $_ -Parent
                            }
                            else
                            {
                                $_
                            }    
                        }

                        Copy-Item -Recurse -Path $source -Destination "$($drive.DriveLetter):\Program Files\WindowsPowerShell\Modules"
                    }
                }
                Copy-Item -Path $Machine.InitialDscConfigurationMofPath -Destination "$($drive.DriveLetter):\Windows\System32\configuration\pending.mof"
            }

            if ($Machine.InitialDscLcmConfigurationMofPath)
            {
                Copy-Item -Path $Machine.InitialDscLcmConfigurationMofPath -Destination "$($drive.DriveLetter):\Windows\System32\configuration\MetaConfig.mof"
            }
        }
        finally
        {
            $mountedOsDisk | Dismount-VHD
        }
    }

    Write-ProgressIndicator

    $vmParameter = @{
        Name               = $Machine.ResourceName
        MemoryStartupBytes = ($Machine.Memory)
        VHDPath            = $systemDisk.Path
        Path               = $VmPath
        Generation         = $generation
        ErrorAction        = 'Stop'
    }

    $vm = Hyper-V\New-VM @vmParameter

    Set-LWHypervVMDescription -ComputerName $Machine.ResourceName -Hashtable @{
        CreatedBy    = '{0} ({1})' -f $PSCmdlet.MyInvocation.MyCommand.Module.Name, $PSCmdlet.MyInvocation.MyCommand.Module.Version
        CreationTime = Get-Date
        LabName      = (Get-Lab).Name
        InitState    = [AutomatedLab.LabVMInitState]::Uninitialized
    }

    #Removing this check as this 'Get-SecureBootUEFI' is not supported on Azure VMs for nested virtualization
    #$isUefi = try
    #{
    #    Get-SecureBootUEFI -Name SetupMode
    #}
    #catch { }

    if ($vm.Generation -ge 2)
    {
        $secureBootTemplate = if ($Machine.HypervProperties.SecureBootTemplate)
        {
            $Machine.HypervProperties.SecureBootTemplate
        }
        else
        {
            if ($Machine.LinuxType -eq 'unknown')
            {
                'MicrosoftWindows'
            }
            else
            {
                'MicrosoftUEFICertificateAuthority'
            }
        }

        $vmFirmwareParameters = @{}

        if ($Machine.HypervProperties.EnableSecureBoot)
        {
            $vmFirmwareParameters.EnableSecureBoot = 'On'
            $vmFirmwareParameters.SecureBootTemplate = $secureBootTemplate
        }
        else
        {
            $vmFirmwareParameters.EnableSecureBoot = 'Off'
        }

        $vm | Set-VMFirmware @vmFirmwareParameters

        if ($Machine.HyperVProperties.EnableTpm -match '1|true|yes')
        {
            $vm | Set-VMKeyProtector -NewLocalKeyProtector
            $vm | Enable-VMTPM
        }
    }

    #remove the unconnected default network adapter
    $vm | Remove-VMNetworkAdapter
    foreach ($adapter in $adapters)
    {
        #bind all network adapters to their designated switches, Repair-LWHypervNetworkConfig will change the binding order if necessary
        $parameters = @{
            Name             = $adapter.VirtualSwitch.ResourceName
            SwitchName       = $adapter.VirtualSwitch.ResourceName
            StaticMacAddress = $adapter.MacAddress
            VMName           = $vm.Name
            PassThru         = $true
        }

        if (-not (Get-LabConfigurationItem -Name DisableDeviceNaming -Default $false) -and (Get-Command Add-VMNetworkAdapter).Parameters.Values.Name -contains 'DeviceNaming' -and $vm.Generation -eq 2 -and $Machine.OperatingSystem.Version -ge 10.0)
        {
            $parameters['DeviceNaming'] = 'On'
        }

        $newAdapter = Add-VMNetworkAdapter @parameters

        if (-not $adapter.AccessVLANID -eq 0)
        {

            Set-VMNetworkAdapterVlan -VMNetworkAdapter $newAdapter -Access -VlanId $adapter.AccessVLANID
            Write-PSFMessage "Network Adapter: '$($adapter.VirtualSwitch.ResourceName)' for VM: '$($vm.Name)' created with VLAN ID: '$($adapter.AccessVLANID)', Ensure external routing is configured correctly"
        }
    }

    Write-PSFMessage "`tMachine '$Name' created"

    $automaticStartAction = 'Nothing'
    $automaticStartDelay = 0
    $automaticStopAction = 'ShutDown'

    if ($Machine.HypervProperties.AutomaticStartAction) { $automaticStartAction = $Machine.HypervProperties.AutomaticStartAction }
    if ($Machine.HypervProperties.AutomaticStartDelay)  { $automaticStartDelay  = $Machine.HypervProperties.AutomaticStartDelay  }
    if ($Machine.HypervProperties.AutomaticStopAction)  { $automaticStopAction  = $Machine.HypervProperties.AutomaticStopAction  }
    $vm | Hyper-V\Set-VM -AutomaticStartAction $automaticStartAction -AutomaticStartDelay $automaticStartDelay -AutomaticStopAction $automaticStopAction

    Write-ProgressIndicator

    if ( $Machine.OperatingSystemType -eq 'Linux')
    {
        $dvd = $vm | Add-VMDvdDrive -Path $Machine.OperatingSystem.IsoPath -Passthru
        if ( $Machine.LinuxType -in 'RedHat','Ubuntu') {
            $vm | Set-VMFirmware -FirstBootDevice $dvd
        }
    }

    if ( $Machine.OperatingSystemType -eq 'Windows')
    {
        [void](Mount-DiskImage -ImagePath $path)
        $VhdDisk = Get-DiskImage -ImagePath $path | Get-Disk
        $VhdPartition = Get-Partition -DiskNumber $VhdDisk.Number

        if ($VhdPartition.Count -gt 1)
        {
            #for Generation 2 VMs
            $vhdOsPartition = $VhdPartition | Where-Object Type -eq 'Basic'
            # If no drive letter is assigned, make sure we assign it before continuing
            If ($vhdOsPartition.NoDefaultDriveLetter)
            {
                # Get all available drive letters, and store in a temporary variable.
                $usedDriveLetters = @(Get-Volume | ForEach-Object { "$([char]$_.DriveLetter)" }) + @(Get-CimInstance -ClassName Win32_MappedLogicalDisk | ForEach-Object { $([char]$_.DeviceID.Trim(':')) })
                [char[]]$tempDriveLetters = Compare-Object -DifferenceObject $usedDriveLetters -ReferenceObject $( 67..90 | ForEach-Object { "$([char]$_)" }) -PassThru | Where-Object { $_.SideIndicator -eq '<=' }
                # Sort the available drive letters to get the first available drive letter
                $availableDriveLetters = ($TempDriveLetters | Sort-Object)
                $firstAvailableDriveLetter = $availableDriveLetters[0]
                $vhdOsPartition | Set-Partition -NewDriveLetter $firstAvailableDriveLetter
                $VhdVolume = "$($firstAvailableDriveLetter):"

            }
            Else
            {
                $VhdVolume = "$($vhdOsPartition.DriveLetter):"
            }
        }
        else
        {
            #for Generation 1 VMs
            $VhdVolume = "$($VhdPartition.DriveLetter):"
        }
        Write-PSFMessage "`tDisk mounted to drive $VhdVolume"

        #Get-PSDrive needs to be called to update the PowerShell drive list
        Get-PSDrive | Out-Null

        #copy AL tools to lab machine and optionally the tools folder
        $drive = New-PSDrive -Name $VhdVolume[0] -PSProvider FileSystem -Root $VhdVolume

        Write-PSFMessage 'Copying AL tools to VHD...'
        $tempPath = "$([System.IO.Path]::GetTempPath())$([System.IO.Path]::GetRandomFileName())"
        New-Item -ItemType Directory -Path $tempPath | Out-Null
        Copy-Item -Path "$((Get-Module -Name AutomatedLabCore)[0].ModuleBase)\Tools\HyperV\*" -Destination $tempPath -Recurse
        foreach ($file in (Get-ChildItem -Path $tempPath -Recurse -File))
        {
            # Why???
            if ($PSEdition -eq 'Desktop')
            {
                $file.Decrypt()
            }
        }

        Copy-Item -Path "$tempPath\*" -Destination "$vhdVolume\Windows" -Recurse

        Remove-Item -Path $tempPath -Recurse -ErrorAction SilentlyContinue

        Write-PSFMessage '...done'

        

        if ($Machine.OperatingSystemType -eq 'Windows' -and -not [string]::IsNullOrEmpty($Machine.SshPublicKey))
        {
            Add-UnattendedSynchronousCommand -Command 'PowerShell -File "C:\Program Files\OpenSSH-Win64\install-sshd.ps1"' -Description 'Configure SSH'
            Add-UnattendedSynchronousCommand -Command 'PowerShell -Command "Set-Service -Name sshd -StartupType Automatic"' -Description 'Enable SSH'
            Add-UnattendedSynchronousCommand -Command 'PowerShell -Command "Restart-Service -Name sshd"' -Description 'Restart SSH'

            Write-PSFMessage 'Copying PowerShell 7 and setting up SSH'
            $release = try { Invoke-RestMethod -Uri 'https://api.github.com/repos/powershell/powershell/releases/latest' -UseBasicParsing -ErrorAction Stop } catch {}
            $uri = ($release.assets | Where-Object name -like '*-win-x64.zip').browser_download_url
            if (-not $uri)
            {
                $uri = 'https://github.com/PowerShell/PowerShell/releases/download/v7.2.6/PowerShell-7.2.6-win-x64.zip'
            }
            $psArchive = Get-LabInternetFile -Uri $uri -Path "$labSources/SoftwarePackages/PS7.zip"

        
            $release = try { Invoke-RestMethod -Uri 'https://api.github.com/repos/powershell/win32-openssh/releases/latest' -UseBasicParsing -ErrorAction Stop } catch {}
            $uri = ($release.assets | Where-Object name -like '*-win64.zip').browser_download_url
            if (-not $uri)
            {
                $uri = 'https://github.com/PowerShell/Win32-OpenSSH/releases/download/v8.9.1.0p1-Beta/OpenSSH-Win64.zip'
            }
            $sshArchive = Get-LabInternetFile -Uri $uri -Path "$labSources/SoftwarePackages/ssh.zip"

            $null = New-Item -ItemType Directory -Force -Path (Join-Path -Path $vhdVolume -ChildPath 'Program Files\PowerShell\7')
            Expand-Archive -Path "$labSources/SoftwarePackages/PS7.zip" -DestinationPath (Join-Path -Path $vhdVolume -ChildPath 'Program Files\PowerShell\7')
            Expand-Archive -Path "$labSources/SoftwarePackages/ssh.zip" -DestinationPath (Join-Path -Path $vhdVolume -ChildPath 'Program Files')

            $null = New-Item -ItemType File -Path (Join-Path -Path $vhdVolume -ChildPath '\AL\SSH\keys'), (Join-Path -Path $vhdVolume -ChildPath 'ProgramData\ssh\sshd_config') -Force
        
            $Machine.SshPublicKey | Add-Content -Path (Join-Path -Path $vhdVolume -ChildPath '\AL\SSH\keys')
        
            $sshdConfig = @"
Port 22
PasswordAuthentication no
PubkeyAuthentication yes
GSSAPIAuthentication yes
AllowGroups Users Administrators
AuthorizedKeysFile c:/al/ssh/keys
Subsystem powershell c:/progra~1/powershell/7/pwsh.exe -sshs -NoLogo
"@
            $sshdConfig | Set-Content -Path (Join-Path -Path $vhdVolume -ChildPath 'ProgramData\ssh\sshd_config')
            Write-PSFMessage 'Done'
        }

        if ($Machine.ToolsPath.Value)
        {
            $toolsDestination = "$vhdVolume\Tools"
            if ($Machine.ToolsPathDestination)
            {
                $toolsDestination = "$($toolsDestination[0])$($Machine.ToolsPathDestination.Substring(1,$Machine.ToolsPathDestination.Length - 1))"
            }
            Write-PSFMessage 'Copying tools to VHD...'
            Copy-Item -Path $Machine.ToolsPath -Destination $toolsDestination -Recurse
            Write-PSFMessage '...done'
        }

        $enableWSManRegDump = @'
Windows Registry Editor Version 5.00

[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WSMAN]
"StackVersion"="2.0"
"UpdatedConfig"="857C6BDB-A8AC-4211-93BB-8123C9ECE4E5"

[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WSMAN\Listener\*+HTTP]
"uriprefix"="wsman"
"Port"=dword:00001761

[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WSMAN\Plugin\Event Forwarding Plugin]
"ConfigXML"="<PlugInConfiguration xmlns=\"http://schemas.microsoft.com/wbem/wsman/1/config/PluginConfiguration\" Name=\"Event Forwarding Plugin\" Filename=\"C:\\Windows\\system32\\wevtfwd.dll\" SDKVersion=\"1\" XmlRenderingType=\"text\" UseSharedProcess=\"false\" ProcessIdleTimeoutSec=\"0\" RunAsUser=\"\" RunAsPassword=\"\" AutoRestart=\"false\" Enabled=\"true\" OutputBufferingMode=\"Block\" ><Resources><Resource ResourceUri=\"http://schemas.microsoft.com/wbem/wsman/1/windows/EventLog\" SupportsOptions=\"true\" ><Security Uri=\"\" ExactMatch=\"false\" Sddl=\"O:NSG:BAD:P(A;;GA;;;BA)(A;;GR;;;ER)S:P(AU;FA;GA;;;WD)(AU;SA;GWGX;;;WD)\" /><Capability Type=\"Subscribe\" SupportsFiltering=\"true\" /></Resource></Resources><Quotas MaxConcurrentUsers=\"100\" MaxConcurrentOperationsPerUser=\"15\" MaxConcurrentOperations=\"1500\"/></PlugInConfiguration>"

[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WSMAN\Plugin\Microsoft.PowerShell]
"ConfigXML"="<PlugInConfiguration xmlns=\"http://schemas.microsoft.com/wbem/wsman/1/config/PluginConfiguration\" Name=\"microsoft.powershell\" Filename=\"%windir%\\system32\\pwrshplugin.dll\" SDKVersion=\"2\" XmlRenderingType=\"text\" Enabled=\"true\" Architecture=\"64\" UseSharedProcess=\"false\" ProcessIdleTimeoutSec=\"0\" RunAsUser=\"\" RunAsPassword=\"\" AutoRestart=\"false\" OutputBufferingMode=\"Block\"><InitializationParameters><Param Name=\"PSVersion\" Value=\"3.0\"/></InitializationParameters><Resources><Resource ResourceUri=\"http://schemas.microsoft.com/powershell/microsoft.powershell\" SupportsOptions=\"true\" ExactMatch=\"true\"><Security Uri=\"http://schemas.microsoft.com/powershell/microsoft.powershell\" Sddl=\"O:NSG:BAD:P(A;;GA;;;BA)(A;;GA;;;RM)S:P(AU;FA;GA;;;WD)(AU;SA;GXGW;;;WD)\" ExactMatch=\"False\"/><Capability Type=\"Shell\"/></Resource></Resources><Quotas MaxIdleTimeoutms=\"2147483647\" MaxConcurrentUsers=\"5\" IdleTimeoutms=\"7200000\" MaxProcessesPerShell=\"15\" MaxMemoryPerShellMB=\"1024\" MaxConcurrentCommandsPerShell=\"1000\" MaxShells=\"25\" MaxShellsPerUser=\"25\"/></PlugInConfiguration>"

[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WSMAN\Plugin\Microsoft.PowerShell.Workflow]
"ConfigXML"="<PlugInConfiguration xmlns=\"http://schemas.microsoft.com/wbem/wsman/1/config/PluginConfiguration\" Name=\"microsoft.powershell.workflow\" Filename=\"%windir%\\system32\\pwrshplugin.dll\" SDKVersion=\"2\" XmlRenderingType=\"text\" UseSharedProcess=\"true\" ProcessIdleTimeoutSec=\"28800\" RunAsUser=\"\" RunAsPassword=\"\" AutoRestart=\"false\" Enabled=\"true\" Architecture=\"64\" OutputBufferingMode=\"Block\"><InitializationParameters><Param Name=\"PSVersion\" Value=\"3.0\"/><Param Name=\"AssemblyName\" Value=\"Microsoft.PowerShell.Workflow.ServiceCore, Version=3.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35, processorArchitecture=MSIL\"/><Param Name=\"PSSessionConfigurationTypeName\" Value=\"Microsoft.PowerShell.Workflow.PSWorkflowSessionConfiguration\"/><Param Name=\"SessionConfigurationData\" Value=\"                             &lt;SessionConfigurationData&gt;                                 &lt;Param Name=&quot;ModulesToImport&quot; Value=&quot;%windir%\\system32\\windowspowershell\\v1.0\\Modules\\PSWorkflow&quot;/&gt;                                 &lt;Param Name=&quot;PrivateData&quot;&gt;                                     &lt;PrivateData&gt;                                         &lt;Param Name=&quot;enablevalidation&quot; Value=&quot;true&quot; /&gt;                                     &lt;/PrivateData&gt;                                 &lt;/Param&gt;                             &lt;/SessionConfigurationData&gt;                         \"/></InitializationParameters><Resources><Resource ResourceUri=\"http://schemas.microsoft.com/powershell/microsoft.powershell.workflow\" SupportsOptions=\"true\" ExactMatch=\"true\"><Security Uri=\"http://schemas.microsoft.com/powershell/microsoft.powershell.workflow\" Sddl=\"O:NSG:BAD:P(A;;GA;;;BA)(A;;GA;;;RM)S:P(AU;FA;GA;;;WD)(AU;SA;GXGW;;;WD)\" ExactMatch=\"False\"/><Capability Type=\"Shell\"/></Resource></Resources><Quotas MaxIdleTimeoutms=\"2147483647\" MaxConcurrentUsers=\"5\" IdleTimeoutms=\"7200000\" MaxProcessesPerShell=\"15\" MaxMemoryPerShellMB=\"1024\" MaxConcurrentCommandsPerShell=\"1000\" MaxShells=\"25\" MaxShellsPerUser=\"25\"/></PlugInConfiguration>"

[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WSMAN\Plugin\Microsoft.PowerShell32]
"ConfigXML"="<PlugInConfiguration xmlns=\"http://schemas.microsoft.com/wbem/wsman/1/config/PluginConfiguration\" Name=\"microsoft.powershell32\" Filename=\"%windir%\\system32\\pwrshplugin.dll\" SDKVersion=\"2\" XmlRenderingType=\"text\" Architecture=\"32\" Enabled=\"true\" UseSharedProcess=\"false\" ProcessIdleTimeoutSec=\"0\" RunAsUser=\"\" RunAsPassword=\"\" AutoRestart=\"false\" OutputBufferingMode=\"Block\"><InitializationParameters><Param Name=\"PSVersion\" Value=\"3.0\"/></InitializationParameters><Resources><Resource ResourceUri=\"http://schemas.microsoft.com/powershell/microsoft.powershell32\" SupportsOptions=\"true\" ExactMatch=\"true\"><Security Uri=\"http://schemas.microsoft.com/powershell/microsoft.powershell32\" Sddl=\"O:NSG:BAD:P(A;;GA;;;BA)(A;;GA;;;RM)S:P(AU;FA;GA;;;WD)(AU;SA;GXGW;;;WD)\" ExactMatch=\"False\"/><Capability Type=\"Shell\"/></Resource></Resources><Quotas MaxIdleTimeoutms=\"2147483647\" MaxConcurrentUsers=\"5\" IdleTimeoutms=\"7200000\" MaxProcessesPerShell=\"15\" MaxMemoryPerShellMB=\"1024\" MaxConcurrentCommandsPerShell=\"1000\" MaxShells=\"25\" MaxShellsPerUser=\"25\"/></PlugInConfiguration>"

[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WSMAN\Plugin\WMI Provider]
"ConfigXML"="<PlugInConfiguration xmlns=\"http://schemas.microsoft.com/wbem/wsman/1/config/PluginConfiguration\" Name=\"WMI Provider\" Filename=\"C:\\Windows\\system32\\WsmWmiPl.dll\" SDKVersion=\"1\" XmlRenderingType=\"text\" UseSharedProcess=\"false\" ProcessIdleTimeoutSec=\"0\" RunAsUser=\"\" RunAsPassword=\"\" AutoRestart=\"false\" Enabled=\"true\" OutputBufferingMode=\"Block\" ><Resources><Resource ResourceUri=\"http://schemas.microsoft.com/wbem/wsman/1/wmi\" SupportsOptions=\"true\" ><Security Uri=\"\" ExactMatch=\"false\" Sddl=\"O:NSG:BAD:P(A;;GA;;;BA)(A;;GA;;;IU)(A;;GA;;;RM)S:P(AU;FA;GA;;;WD)(AU;SA;GWGX;;;WD)\" /><Capability Type=\"Identify\" /><Capability Type=\"Get\" SupportsFragment=\"true\" /><Capability Type=\"Put\" SupportsFragment=\"true\" /><Capability Type=\"Invoke\" /><Capability Type=\"Create\" /><Capability Type=\"Delete\" /><Capability Type=\"Enumerate\" SupportsFiltering=\"true\"/><Capability Type=\"Subscribe\" SupportsFiltering=\"true\"/></Resource><Resource ResourceUri=\"http://schemas.dmtf.org/wbem/wscim/1/cim-schema\" SupportsOptions=\"true\" ><Security Uri=\"\" ExactMatch=\"false\" Sddl=\"O:NSG:BAD:P(A;;GA;;;BA)(A;;GA;;;IU)(A;;GA;;;RM)S:P(AU;FA;GA;;;WD)(AU;SA;GWGX;;;WD)\" /><Capability Type=\"Get\" SupportsFragment=\"true\" /><Capability Type=\"Put\" SupportsFragment=\"true\" /><Capability Type=\"Invoke\" /><Capability Type=\"Create\" /><Capability Type=\"Delete\" /><Capability Type=\"Enumerate\"/><Capability Type=\"Subscribe\" SupportsFiltering=\"true\"/></Resource><Resource ResourceUri=\"http://schemas.dmtf.org/wbem/wscim/1/*\" SupportsOptions=\"true\" ExactMatch=\"true\" ><Security Uri=\"\" ExactMatch=\"false\" Sddl=\"O:NSG:BAD:P(A;;GA;;;BA)(A;;GA;;;IU)(A;;GA;;;RM)S:P(AU;FA;GA;;;WD)(AU;SA;GWGX;;;WD)\" /><Capability Type=\"Enumerate\" SupportsFiltering=\"true\"/><Capability Type=\"Subscribe\"SupportsFiltering=\"true\"/></Resource><Resource ResourceUri=\"http://schemas.dmtf.org/wbem/cim-xml/2/cim-schema/2/*\" SupportsOptions=\"true\" ExactMatch=\"true\"><Security Uri=\"\" ExactMatch=\"false\" Sddl=\"O:NSG:BAD:P(A;;GA;;;BA)(A;;GA;;;IU)(A;;GA;;;RM)S:P(AU;FA;GA;;;WD)(AU;SA;GWGX;;;WD)\" /><Capability Type=\"Get\" SupportsFragment=\"false\"/><Capability Type=\"Enumerate\" SupportsFiltering=\"true\"/></Resource></Resources><Quotas MaxConcurrentUsers=\"100\" MaxConcurrentOperationsPerUser=\"100\" MaxConcurrentOperations=\"1500\"/></PlugInConfiguration>"

[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WSMAN\Service]
"allow_remote_requests"=dword:00000001
'@
        #Using the .net class as the PowerShell provider usually does not recognize the new drive
        [System.IO.File]::WriteAllText("$vhdVolume\WSManRegKey.reg", $enableWSManRegDump)

        $additionalDisksOnline = @"
`$deployDebug = (New-Item -ItemType Directory -Path `$ExecutionContext.InvokeCommand.ExpandString("$AL_DeployDebugFolder") -Force).FullName
Start-Transcript -Path `$deployDebug\AdditionalDisksOnline.log
`$diskpartCmd = 'LIST DISK'
`$disks = `$diskpartCmd | diskpart.exe
`$pattern = 'Disk (?<DiskNumber>\d{1,3}) \s+(?<State>Online|Offline)\s+(?<Size>\d+) (KB|MB|GB|TB)\s+(?<Free>\d+) (B|KB|MB|GB|TB)'
foreach (`$line in `$disks)
{
    if (`$line -match `$pattern)
    {
        #`$nextDriveLetter = [char[]](67..90) |
        #Where-Object { (Get-CimInstance -Class Win32_LogicalDisk |
        #Select-Object -ExpandProperty DeviceID) -notcontains "`$(`$_):"} |
        #Select-Object -First 1
        `$diskNumber = `$Matches.DiskNumber
        if (`$Matches.State -eq 'Offline')
        {
            `$diskpartCmd = "@
                SELECT DISK `$diskNumber
                ATTRIBUTES DISK CLEAR READONLY
                ONLINE DISK
                EXIT
            @"
            `$diskpartCmd | diskpart.exe | Out-Null
        }
    }
}
foreach (`$volume in (Get-WmiObject -Class Win32_Volume))
{
    if (`$volume.Label -notmatch '(?<Label>[-_\w\d]+)_AL_(?<DriveLetter>[A-Z])')
    {
        continue
    }
        if (`$volume.DriveLetter -ne "`$(`$Matches.DriveLetter):")
    {
        `$volume.DriveLetter = "`$(`$Matches.DriveLetter):"
    }
        `$volume.Label = `$Matches.Label
    `$volume.Put()
}
Stop-Transcript
"@
        [System.IO.File]::WriteAllText("$vhdVolume\AdditionalDisksOnline.ps1", $additionalDisksOnline)

        $defaultSettings = @{
            WinRmMaxEnvelopeSizeKb              = 500
            WinRmMaxConcurrentOperationsPerUser = 1500
            WinRmMaxConnections                 = 300
        }
    
        $command = 'Start-Service WinRm'
        foreach ($setting in $defaultSettings.GetEnumerator())
        {
            $settingValue = if ((Get-LabConfigurationItem -Name $setting.Key) -ne $setting.Value)
            {
                Get-LabConfigurationItem -Name $setting.Key
            }
            else
            {
                $setting.Value
            }

            $subdir = if ($setting.Key -match 'MaxEnvelope') { $null } else { 'Service\' }
            $command = -join @($command, "`r`nSet-Item WSMAN:\localhost\$subdir$($setting.Key.Replace('WinRm','')) $($settingValue) -Force")
        }

        [System.IO.File]::WriteAllText("$vhdVolume\WinRmCustomization.ps1", $command)
    
        Write-ProgressIndicator
        
        $unattendXmlContent = Get-UnattendedContent
        $unattendXmlContent.Save("$VhdVolume\Unattend.xml")
        Write-PSFMessage "`tUnattended file copied to VM Disk '$vhdVolume\unattend.xml'"
        
        [void](Dismount-DiskImage -ImagePath $path)
        Write-PSFMessage "`tdisk image dismounted"
    }    

    Write-PSFMessage "`tSettings RAM, start and stop actions"
    $param = @{}
    $param.Add('MemoryStartupBytes', $Machine.Memory)
    $param.Add('AutomaticCheckpointsEnabled', $false)
    $param.Add('CheckpointType', 'Production')

    if ($Machine.MaxMemory) { $param.Add('MemoryMaximumBytes', $Machine.MaxMemory) }
    if ($Machine.MinMemory) { $param.Add('MemoryMinimumBytes', $Machine.MinMemory) }

    if ($Machine.MaxMemory -or $Machine.MinMemory)
    {
        $param.Add('DynamicMemory', $true)
        Write-PSFMessage "`tSettings dynamic memory to MemoryStartupBytes $($Machine.Memory), minimum $($Machine.MinMemory), maximum $($Machine.MaxMemory)"
    }
    else
    {
        Write-PSFMessage "`tSettings static memory to $($Machine.Memory)"
        $param.Add('StaticMemory', $true)
    }

    $param = Sync-Parameter -Command (Get-Command Set-Vm) -Parameters $param

    Hyper-V\Set-VM -Name $Machine.ResourceName @param

    Hyper-V\Set-VM -Name $Machine.ResourceName -ProcessorCount $Machine.Processors

    if ($DisableIntegrationServices)
    {
        Disable-VMIntegrationService -VMName $Machine.ResourceName -Name 'Time Synchronization'
    }

    if ($Generation -eq 1)
    {
        Set-VMBios -VMName $Machine.ResourceName -EnableNumLock
    }

    Write-PSFMessage "Creating snapshot named '$($Machine.ResourceName) - post OS Installation'"
    if ($CreateCheckPoints)
    {
        Hyper-V\Checkpoint-VM -VM (Hyper-V\Get-VM -Name $Machine.ResourceName) -SnapshotName 'Post OS Installation'
    }

    if ($Machine.Disks.Name)
    {
        $disks = Get-LabVHDX -Name $Machine.Disks.Name
        foreach ($disk in $disks)
        {
            Add-LWVMVHDX -VMName $Machine.ResourceName -VhdxPath $disk.Path
        }
    }

    Write-ProgressIndicatorEnd

    $writeVmConnectConfigFile = Get-LabConfigurationItem -Name VMConnectWriteConfigFile
    if ($writeVmConnectConfigFile)
    {
        New-LWHypervVmConnectSettingsFile -VmName $Machine.ResourceName
    }

    Write-LogFunctionExit

    return $true
}


function New-LWHypervVmConnectSettingsFile
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseCompatibleCmdlets", "", Justification="Not relevant on Linux")]
    [Cmdletbinding(PositionalBinding = $false)]
    #In the parameter block, 'HelpMessageResourceId' is misused to store the type that is stored in the config file.
    #'HelpMessageResourceId' does not have any effect on the parameter itself.
    param (
        [Parameter(HelpMessageResourceId = 'System.Boolean')]
        [bool]$AudioCaptureRedirectionMode = $false,
        
        [Parameter(HelpMessageResourceId = 'System.Boolean')]
        [bool]$EnablePrinterRedirection = $false,
        
        [Parameter(HelpMessageResourceId = 'System.Boolean')]
        [bool]$FullScreen = (Get-LabConfigurationItem -Name VMConnectFullScreen -Default $false),
        
        [Parameter(HelpMessageResourceId = 'System.Boolean')]
        [bool]$SmartCardsRedirection = $true,
        
        [Parameter(HelpMessageResourceId = 'System.String')]
        [string]$RedirectedPnpDevices,
        
        [Parameter(HelpMessageResourceId = 'System.String')]
        [bool]$ClipboardRedirection = $true,
        
        [Parameter(HelpMessageResourceId = 'System.Drawing.Size')]
        [string]$DesktopSize = (Get-LabConfigurationItem -Name VMConnectDesktopSize -Default '1366, 768'),

        [Parameter(HelpMessageResourceId = 'System.String')]
        [string]$VmServerName = $env:COMPUTERNAME,
        
        [Parameter(HelpMessageResourceId = 'System.String')]
        [string]$RedirectedUsbDevices,
        
        [Parameter(HelpMessageResourceId = 'System.Boolean')]
        [bool]$SavedConfigExists = $true,
        
        [Parameter(HelpMessageResourceId = 'System.Boolean')]
        [bool]$UseAllMonitors = (Get-LabConfigurationItem -Name VMConnectUseAllMonitors -Default $false),
        
        [Parameter(HelpMessageResourceId = 'Microsoft.Virtualization.Client.RdpOptions+AudioPlaybackRedirectionTyp')]
        [string]$AudioPlaybackRedirectionMode = 'AUDIO_MODE_REDIRECT',
        
        [Parameter(HelpMessageResourceId = 'System.Boolean')]
        [bool]$PrinterRedirection,
        
        [Parameter(HelpMessageResourceId = 'System.String')]
        [string]$RedirectedDrives = (Get-LabConfigurationItem -Name VMConnectRedirectedDrives -Default ''),
        
        [Parameter(Mandatory, HelpMessageResourceId = 'System.String')]
        [Alias('ComputerName')]
        [string]$VmName,
        
        [Parameter(HelpMessageResourceId = 'System.Boolean')]
        [bool]$SaveButtonChecked = $true
    )
    
    Write-LogFunctionEntry

    #AutomatedLab does not allow empty strings in the configuration, hence the detour.
    if ($RedirectedDrives -eq 'none')
    {
        $RedirectedDrives = ''
    }
    
    $machineVmConnectConfig = [AutomatedLab.Machines.MachineVmConnectConfig]::new()
    $parameters = $MyInvocation.MyCommand.Parameters

    $vm = Get-VM -Name $VmName

    foreach ($parameter in $parameters.GetEnumerator())
    {
        if (-not $parameter.Value.Attributes.HelpMessageResourceId)
        {
            continue
        }
        
        $value = Get-Variable -Name $parameter.Key -ValueOnly -ErrorAction SilentlyContinue
        $setting = [AutomatedLab.Machines.MachineVmConnectRdpOptionSetting]::new()
        
        $setting.Name = $parameter.Key
        $setting.Type = $parameter.Value.Attributes.HelpMessageResourceId
        $setting.Value = $value
        
        $machineVmConnectConfig.Settings.Add($setting)
        
        #Files will be stored in path 'C:\Users\<Username>\AppData\Roaming\Microsoft\Windows\Hyper-V\Client\1.0'
        $configFilePath = '{0}\Microsoft\Windows\Hyper-V\Client\1.0\vmconnect.rdp.{1}.config' -f $env:APPDATA, $vm.Id
        $configFileParentPath = Split-Path -Path $configFilePath -Parent
        if (-not (Test-Path -Path $configFileParentPath -PathType Container))
        {
            mkdir -Path $configFileParentPath -Force | Out-Null
        }
        $machineVmConnectConfig.Export($configFilePath)
    }
    
    Write-LogFunctionExit

}


function Remove-LWHypervVM
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseCompatibleCmdlets", "", Justification="Not relevant on Linux")]
    Param (
        [Parameter(Mandatory)]
        [string]$Name
    )

    Write-LogFunctionEntry

    $vm = Get-LWHypervVM -Name $Name -ErrorAction SilentlyContinue

    if (-not $vm) { Write-LogFunctionExit}

    $vmPath = Split-Path -Path $vm.HardDrives[0].Path -Parent

    if ($vm.State -eq 'Saved')
    {
        Write-PSFMessage "Deleting saved state of VM '$($Name)'"
        $vm | Remove-VMSavedState
    }
    else
    {
        Write-PSFMessage "Stopping VM '$($Name)'"
        $vm | Hyper-V\Stop-VM -TurnOff -Force -WarningAction SilentlyContinue
    }

    Write-PSFMessage "Removing VM '$($Name)'"
    $doNotAddToCluster = Get-LabConfigurationItem -Name DoNotAddVmsToCluster -Default $false
    if (-not $doNotAddToCluster -and (Get-Command -Name Get-Cluster -Module FailoverClusters -CommandType Cmdlet -ErrorAction SilentlyContinue) -and (Get-Cluster -ErrorAction SilentlyContinue -WarningAction SilentlyContinue))
    {
        Write-PSFMessage "Removing Clustered Resource: $Name"
        $null = Get-ClusterGroup -Name $Name | Remove-ClusterGroup -RemoveResources -Force
    }

    Remove-LWHypervVmConnectSettingsFile -ComputerName $Name

    $vm | Hyper-V\Remove-VM -Force

    Write-PSFMessage "Removing VM files for '$($Name)'"
    Remove-Item -Path $vmPath -Force -Confirm:$false -Recurse
    
    $vmDescription = Join-Path -Path (Get-Lab).LabPath -ChildPath "$Name.xml"
    if (Test-Path -Path $vmDescription) {
        Remove-Item -Path $vmDescription
    }

    Write-LogFunctionExit
}


function Remove-LWHypervVmConnectSettingsFile
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseCompatibleCmdlets", "", Justification="Not relevant on Linux")]
    [Cmdletbinding()]
    param (
        [Parameter(Mandatory)]
        [string]$ComputerName
    )
    
    Write-LogFunctionEntry
    
    $vm = Get-VM -Name $ComputerName
    
    $configFilePath = '{0}\Microsoft\Windows\Hyper-V\Client\1.0\vmconnect.rdp.{1}.config' -f $env:APPDATA, $vm.Id
    if (Test-Path -Path $configFilePath)
    {
        Remove-Item -Path $configFilePath -ErrorAction SilentlyContinue
    }
    
    Write-LogFunctionExit
}


function Remove-LWHypervVMSnapshot
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseCompatibleCmdlets", "", Justification="Not relevant on Linux")]
    [Cmdletbinding()]
    Param (
        [Parameter(Mandatory, ParameterSetName = 'BySnapshotName')]
        [Parameter(Mandatory, ParameterSetName = 'AllSnapshots')]
        [string[]]$ComputerName,

        [Parameter(Mandatory, ParameterSetName = 'BySnapshotName')]
        [string]$SnapshotName,

        [Parameter(ParameterSetName = 'AllSnapshots')]
        [switch]$All
    )

    Write-LogFunctionEntry
    $pool = New-RunspacePool -ThrottleLimit 20 -Variable (Get-Variable -Name SnapshotName,All -ErrorAction SilentlyContinue) -Function (Get-Command Get-LWHypervVM)

    $jobs = foreach ($n in $ComputerName)
    {
        Start-RunspaceJob -RunspacePool $pool -Argument $n,(Get-LabConfigurationItem -Name DoNotAddVmsToCluster -Default $false) -ScriptBlock {
            param ($n, $DisableClusterCheck)
            $vm = Get-LWHypervVM -Name $n -DisableClusterCheck $DisableClusterCheck
            if ($SnapshotName)
            {
                $snapshot = $vm | Get-VMSnapshot | Where-Object -FilterScript {
                    $_.Name -eq $SnapshotName
                }
            }
            else
            {
                $snapshot = $vm | Get-VMSnapshot
            }

            if (-not $snapshot)
            {
                Write-Error -Message "The machine '$n' does not have a snapshot named '$SnapshotName'"
            }
            else
            {
                $snapshot | Remove-VMSnapshot -IncludeAllChildSnapshots -ErrorAction SilentlyContinue
            }
        }
    }

    $jobs | Receive-RunspaceJob

    $pool | Remove-RunspacePool

    Write-LogFunctionExit
}


function Repair-LWHypervNetworkConfig
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseCompatibleCmdlets", "", Justification="Not relevant on Linux")]
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ComputerName
    )

    Write-LogFunctionEntry

    $machine = Get-LabVM -ComputerName $ComputerName
    $vm = Get-LWHypervVM -Name $machine.ResourceName

    if (-not $machine) { return } # No fixing this on a Linux VM

    Wait-LabVM -ComputerName $machine -NoNewLine
    $machineAdapterStream = [System.Management.Automation.PSSerializer]::Serialize($machine.NetworkAdapters,2)

    Invoke-LabCommand -ComputerName $machine -ActivityName "Network config on '$machine' (renaming and ordering)" -ScriptBlock {
        Write-Verbose "Renaming network adapters"
        #rename the adapters as defined in the lab
        $machineAdapter = [System.Management.Automation.PSSerializer]::Deserialize($machineAdapterStream)
        $newNames = @()
        foreach ($adapterInfo in $machineAdapter)
        {
            $newName = if ($adapterInfo.InterfaceName)
            {
                $adapterInfo.InterfaceName
            }
            else
            {
                $tempName = Add-StringIncrement -String $adapterInfo.VirtualSwitch.ResourceName
                while ($tempName -in $newNames)
                {
                    $tempName = Add-StringIncrement -String $tempName
                }
                $tempName
            }
            $newNames += $newName

            if (-not [string]::IsNullOrEmpty($adapterInfo.VirtualSwitch.FriendlyName))
            {
                $adapterInfo.VirtualSwitch.FriendlyName = $newName
            }
            else
            {
                $adapterInfo.VirtualSwitch.Name = $newName
            }

            $machineOs = [Environment]::OSVersion
            if ($machineOs.Version.Major -lt 6 -and $machineOs.Version.Minor -lt 2)
            {
                $mac = (Get-StringSection -String $adapterInfo.MacAddress -SectionSize 2) -join ':'
                $filter = 'MACAddress = "{0}"' -f $mac
                Write-Verbose "Looking for network adapter with using filter '$filter'"
                $adapter = Get-CimInstance -Class Win32_NetworkAdapter -Filter $filter

                Write-Verbose "Renaming adapter '$($adapter.NetConnectionID)' -> '$newName'"
                $adapter.NetConnectionID = $newName
                $adapter.Put()
            }
            else
            {
                $mac = (Get-StringSection -String $adapterInfo.MacAddress -SectionSize 2) -join '-'
                Write-Verbose "Renaming adapter '$($adapter.NetConnectionID)' -> '$newName'"
                Get-NetAdapter | Where-Object MacAddress -eq $mac | Rename-NetAdapter -NewName $newName
            }
        }

        #There is no need to change the network binding order in Windows 10 or 2016
        #Adjusting the Network Protocol Bindings in Windows 10 https://blogs.technet.microsoft.com/networking/2015/08/14/adjusting-the-network-protocol-bindings-in-windows-10/
        if ([System.Environment]::OSVersion.Version.Major -lt 10)
        {
            $retries = $machineAdapter.Count * $machineAdapter.Count * 2
            $i = 0

            $sortedAdapters = New-Object System.Collections.ArrayList
            $sortedAdapters.AddRange(@($machineAdapter | Where-Object { $_.VirtualSwitch.SwitchType.Value -ne 'Internal' }))
            $sortedAdapters.AddRange(@($machineAdapter | Where-Object { $_.VirtualSwitch.SwitchType.Value -eq 'Internal' }))

            Write-Verbose "Setting the network order"
            [array]::Reverse($machineAdapter)
            foreach ($adapterInfo in $sortedAdapters)
            {
                Write-Verbose "Setting the order for adapter '$($adapterInfo.VirtualSwitch.ResourceName)'"
                do {
                    nvspbind.exe /+ $adapterInfo.VirtualSwitch.ResourceName ms_tcpip | Out-File -FilePath c:\nvspbind.log -Append
                    $i++

                    if ($i -gt $retries) { return }
                }  until ($LASTEXITCODE -eq 14)
            }
        }

    } -Function (Get-Command -Name Get-StringSection, Add-StringIncrement) -Variable (Get-Variable -Name machineAdapterStream) -NoDisplay

    foreach ($adapterInfo in $machineAdapter)
    {
        $vmAdapter = $vm | Get-VMNetworkAdapter -Name $adapterInfo.VirtualSwitch.ResourceName

        if ($adapterInfo.VirtualSwitch.ResourceName -ne $vmAdapter.SwitchName)
        {
            $vmAdapter | Connect-VMNetworkAdapter -SwitchName $adapterInfo.VirtualSwitch.ResourceName
        }
    }

    Write-LogFunctionExit
}


function Restore-LWHypervVMSnapshot
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseCompatibleCmdlets", "", Justification="Not relevant on Linux")]
    [Cmdletbinding()]
    Param (
        [Parameter(Mandatory)]
        [string[]]$ComputerName,

        [Parameter(Mandatory)]
        [string]$SnapshotName
    )

    Write-LogFunctionEntry

    $pool = New-RunspacePool -ThrottleLimit 20 -Variable (Get-Variable SnapshotName) -Function (Get-Command Get-LWHypervVM)

    Write-PSFMessage -Message 'Remembering all running machines'
    $jobs = foreach ($n in $ComputerName)
    {
        Start-RunspaceJob -RunspacePool $pool -Argument $n,(Get-LabConfigurationItem -Name DoNotAddVmsToCluster -Default $false) -ScriptBlock {
            param ($n, $DisableClusterCheck)

            if ((Get-LWHypervVM -Name $n -DisableClusterCheck $DisableClusterCheck -ErrorAction SilentlyContinue).State -eq 'Running')
            {
                Write-Verbose -Message "    '$n' was running"
                $n
            }
        }
    }

    $runningMachines = $jobs | Receive-RunspaceJob

    $jobs = foreach ($n in $ComputerName)
    {
        Start-RunspaceJob -RunspacePool $pool -Argument $n -ScriptBlock {
            param ($n)
            $vm = Get-LWHypervVM -Name $n
            $vm | Hyper-V\Suspend-VM -ErrorAction SilentlyContinue
            $vm | Hyper-V\Save-VM -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 5
        }
    }

    $jobs | Wait-RunspaceJob

    $jobs = foreach  ($n in $ComputerName)
    {
        Start-RunspaceJob -RunspacePool $pool -Argument $n -ScriptBlock {
            param (
                [string]$n
            )

            $vm = Get-LWHypervVM -Name $n
            $snapshot = $vm | Get-VMSnapshot | Where-Object Name -eq $SnapshotName

            if (-not $snapshot)
            {
                Write-Error -Message "The machine '$n' does not have a snapshot named '$SnapshotName'"
            }
            else
            {
                $snapshot | Restore-VMSnapshot -Confirm:$false
                $vm | Hyper-V\Set-VM -Notes $snapshot.Notes

                Start-Sleep -Seconds 5
            }
        }
    }

    $result = $jobs | Wait-RunspaceJob -PassThru
    if ($result.Shell.HadErrors)
    {
        foreach ($exception in $result.Shell.Streams.Error.Exception)
        {
            Write-Error -Exception $exception
        }
    }

    Write-PSFMessage -Message "Restore finished, starting the machines that were running previously ($($runningMachines.Count))"

    $jobs = foreach ($n in $ComputerName)
    {
        Start-RunspaceJob -RunspacePool $pool -Argument $n,$runningMachines -ScriptBlock {
            param ($n, [string[]]$runningMachines)
            if ($n -in $runningMachines)
            {
                Write-Verbose -Message "Machine '$n' was running, starting it."
                Hyper-V\Start-VM -Name $n -ErrorAction SilentlyContinue
            }
            else
            {
                Write-Verbose -Message "Machine '$n' was NOT running."
            }
        }
    }

    [void] ($jobs | Wait-RunspaceJob)

    $pool | Remove-RunspacePool
    Write-LogFunctionExit
}


function Save-LWHypervVM
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseCompatibleCmdlets", "", Justification="Not relevant on Linux")]
    param (
        [Parameter(Mandatory)]
        [string[]]$ComputerName
    )

    $runspaceScript = {
        param
        (
            [string]$Name,
            [bool]$DisableClusterCheck
        )
        Write-LogFunctionEntry
        Get-LWHypervVM -Name $Name -DisableClusterCheck $DisableClusterCheck | Hyper-V\Save-VM
        Write-LogFunctionExit
    }

    $pool = New-RunspacePool -ThrottleLimit 50 -Function (Get-Command Get-LWHypervVM)

    $jobs = foreach ($Name in $ComputerName)
    {
        Start-RunspaceJob -RunspacePool $pool -ScriptBlock $runspaceScript -Argument $Name,(Get-LabConfigurationItem -Name DoNotAddVmsToCluster -Default $false)
    }

    [void] ($jobs | Wait-RunspaceJob)

    $pool | Remove-RunspacePool
}


function Set-LWHypervVMDescription
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseCompatibleCmdlets", "", Justification="Not relevant on Linux")]
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [hashtable]$Hashtable,

        [Parameter(Mandatory)]
        [string]$ComputerName
    )

    Write-LogFunctionEntry

    $notePath = Join-Path -Path (Get-Lab).LabPath -ChildPath "$ComputerName.xml"

    $type = Get-Type -GenericType AutomatedLab.DictionaryXmlStore -T string, string
    $dictionary = New-Object $type

    foreach ($kvp in $Hashtable.GetEnumerator())
    {
        $dictionary.Add($kvp.Key, $kvp.Value)
    }

    $dictionary.Export($notePath)

    Write-LogFunctionExit
}


function Start-LWHypervVM
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseCompatibleCmdlets", "", Justification="Not relevant on Linux")]
    param (
        [Parameter(Mandatory)]
        [string[]]$ComputerName,

        [int]$DelayBetweenComputers = 0,

        [int]$PreDelaySeconds = 0,

        [int]$PostDelaySeconds = 0,

        [int]$ProgressIndicator,

        [switch]$NoNewLine
    )

    if ($PreDelaySeconds) {
        $job = Start-Job -Name 'Start-LWHypervVM - Pre Delay' -ScriptBlock { Start-Sleep -Seconds $Using:PreDelaySeconds }
        Wait-LWLabJob -Job $job -NoNewLine -ProgressIndicator $ProgressIndicator -Timeout 15 -NoDisplay
    }

    foreach ($Name in $(Get-LabVM -ComputerName $ComputerName -IncludeLinux | Where-Object SkipDeployment -eq $false))
    {
        $machine = Get-LabVM -ComputerName $Name -IncludeLinux

        try
        {
            Get-LWHypervVM -Name $machine.ResourceName | Hyper-V\Start-VM -ErrorAction Stop
        }
        catch
        {
            $ex = New-Object System.Exception("Could not start Hyper-V machine '$ComputerName': $($_.Exception.Message)", $_.Exception)
            throw $ex
        }

        if ($Name.OperatingSystemType -eq 'Linux')
        {
            Write-PSFMessage -Message "Skipping the wait period for $Name as it is a Linux system"
            continue
        }

        if ($DelayBetweenComputers -and $Name -ne $ComputerName[-1])
        {
            $job = Start-Job -Name 'Start-LWHypervVM - DelayBetweenComputers' -ScriptBlock { Start-Sleep -Seconds $Using:DelayBetweenComputers }
            Wait-LWLabJob -Job $job -NoNewLine:$NoNewLine -ProgressIndicator $ProgressIndicator -Timeout 15 -NoDisplay
        }
    }

    if ($PostDelaySeconds)
    {
        $job = Start-Job -Name 'Start-LWHypervVM - Post Delay' -ScriptBlock { Start-Sleep -Seconds $Using:PostDelaySeconds }
        Wait-LWLabJob -Job $job -NoNewLine:$NoNewLine -ProgressIndicator $ProgressIndicator -Timeout 15 -NoDisplay
    }

    Write-LogFunctionExit
}


function Stop-LWHypervVM
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseCompatibleCmdlets", "", Justification="Not relevant on Linux")]
    param (
        [Parameter(Mandatory)]
        [string[]]$ComputerName,

        [double]$TimeoutInMinutes,

        [int]$ProgressIndicator,

        [switch]$NoNewLine,

        [switch]$ShutdownFromOperatingSystem = $true
    )

    Write-LogFunctionEntry

    $start = Get-Date

    if ($ShutdownFromOperatingSystem)
    {
        $jobs = @()
        $linux, $windows = (Get-LabVM -ComputerName $ComputerName -IncludeLinux).Where({ $_.OperatingSystemType -eq 'Linux' }, 'Split')

        if ($windows)
        {
            $jobs += Invoke-LabCommand -ComputerName $windows -NoDisplay -AsJob -PassThru -ErrorAction SilentlyContinue -ErrorVariable invokeErrors -ScriptBlock {
                Stop-Computer -Force -ErrorAction Stop
            }
        }

        if ($linux)
        {
            $jobs += Invoke-LabCommand -UseLocalCredential -ComputerName $linux -NoDisplay -AsJob -PassThru -ScriptBlock {
                #Sleep as background process so that job does not fail.
                [void] (Start-Job -ScriptBlock {
                        Start-Sleep -Seconds 5
                        shutdown -P now
                })
            }
        }

        Wait-LWLabJob -Job $jobs -NoDisplay -ProgressIndicator $ProgressIndicator -NoNewLine:$NoNewLine
        $failedJobs = $jobs | Where-Object { $_.State -eq 'Failed' }
        if ($failedJobs)
        {
            Write-ScreenInfo -Message "Could not stop Hyper-V VM(s): '$($failedJobs.Location)'" -Type Error
        }

        $stopFailures = [System.Collections.Generic.List[string]]::new()
        
        foreach ($failedJob in $failedJobs)
        {
            if (Get-LabVM -ComputerName $failedJob.Location -IncludeLinux)
            {
                $stopFailures.Add($failedJob.Location)
            }
        }

        foreach ($invokeError in $invokeErrors.TargetObject)
        {
            if ($invokeError -is [System.Management.Automation.Runspaces.Runspace] -and $invokeError.ConnectionInfo.ComputerName -as [ipaddress])
            {
                # Special case - return value is an IP address instead of a host name. We need to look it up.
                $stopFailures.Add((Get-LabVM -ComputerName $ComputerName -IncludeLinux | Where-Object Ipv4Address -eq $invokeError.ConnectionInfo.ComputerName).ResourceName)
            }
            elseif ($invokeError -is [System.Management.Automation.Runspaces.Runspace])
            {
                $stopFailures.Add((Get-LabVM -ComputerName $invokeError.ConnectionInfo.ComputerName -IncludeLinux).ResourceName)
            }
        }

        $stopFailures = $stopFailures | Sort-Object -Unique

        if ($stopFailures)
        {
            Write-ScreenInfo -Message "Force-stopping VMs: $($stopFailures -join ',')"
            Get-LWHypervVM -Name $stopFailures | Hyper-V\Stop-VM -Force
        }
    }
    else
    {
        $jobs = @()
        foreach ($name in (Get-LabVM -ComputerName $ComputerName -IncludeLinux | Where-Object SkipDeployment -eq $false).ResourceName)
        {
            $job = Get-LWHypervVM -Name $name -ErrorAction SilentlyContinue | Hyper-V\Stop-VM -AsJob -Force -ErrorAction Stop
            $job | Add-Member -Name ComputerName -MemberType NoteProperty -Value $name
            $jobs += $job
        }
        Wait-LWLabJob -Job $jobs -ProgressIndicator 5 -NoNewLine:$NoNewLine -NoDisplay

        #receive the result of all finished jobs. The result should be null except if an error occured. The error will be returned to the caller
        $jobs | Where-Object State -eq completed | Receive-Job
    }

    Write-LogFunctionExit
}


function Wait-LWHypervVMRestart
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseCompatibleCmdlets", "", Justification="Not relevant on Linux")]
    param (
        [Parameter(Mandatory)]
        [string[]]$ComputerName,

        [double]$TimeoutInMinutes = 15,

        [ValidateRange(1, 300)]
        [int]$ProgressIndicator,

        [AutomatedLab.Machine[]]$StartMachinesWhileWaiting,

        [System.Management.Automation.Job[]]$MonitorJob,

        [switch]$NoNewLine
    )

    Write-LogFunctionEntry

    $machines = Get-LabVM -ComputerName $ComputerName -IncludeLinux

    $machines | Add-Member -Name Uptime -MemberType NoteProperty -Value 0 -Force
    foreach ($machine in $machines)
    {
        $machine.Uptime = (Get-LWHypervVM -Name $machine.ResourceName).Uptime.TotalSeconds
    }

    $vmDrive = ((Get-Lab).Target.Path)[0]
    $start = (Get-Date)
    $progressIndicatorStart = (Get-Date)
    $diskTime = @()
    $lastMachineStart = (Get-Date).AddSeconds(-5)
    $delayedStart = @()

    #$lastMonitorJob = (Get-Date)

    do
    {
        if (((Get-Date) - $progressIndicatorStart).TotalSeconds -gt $ProgressIndicator)
        {
            Write-ProgressIndicator
            $progressIndicatorStart = (Get-Date)
        }

        $diskTime += 100-([int](((Get-Counter -counter "\\$(hostname.exe)\PhysicalDisk(*)\% Idle Time" -SampleInterval 1).CounterSamples | Where-Object {$_.InstanceName -like "*$vmDrive`:*"}).CookedValue))

        if ($StartMachinesWhileWaiting)
        {
            if ($StartMachinesWhileWaiting[0].NetworkAdapters.Count -gt 1)
            {
                $StartMachinesWhileWaiting = $StartMachinesWhileWaiting | Where-Object { $_ -ne $StartMachinesWhileWaiting[0] }
                $delayedStart += $StartMachinesWhileWaiting[0]
            }
            else
            {
                Write-Debug -Message "Disk Time: $($diskTime[-1]). Average (20): $([int](($diskTime[(($diskTime).Count-15)..(($diskTime).Count)] | Measure-Object -Average).Average)) - Average (5): $([int](($diskTime[(($diskTime).Count-5)..(($diskTime).Count)] | Measure-Object -Average).Average))"
                if (((Get-Date) - $lastMachineStart).TotalSeconds -ge 20)
                {
                    if (($diskTime[(($diskTime).Count - 15)..(($diskTime).Count)] | Measure-Object -Average).Average -lt 50 -and ($diskTime[(($diskTime).Count-5)..(($diskTime).Count)] | Measure-Object -Average).Average -lt 60)
                    {
                        Write-PSFMessage -Message 'Starting next machine'
                        $lastMachineStart = (Get-Date)
                        Start-LabVM -ComputerName $StartMachinesWhileWaiting[0] -NoNewline:$NoNewLine
                        $StartMachinesWhileWaiting = $StartMachinesWhileWaiting | Where-Object { $_ -ne $StartMachinesWhileWaiting[0] }
                        if ($StartMachinesWhileWaiting)
                        {
                            Start-LabVM -ComputerName $StartMachinesWhileWaiting[0] -NoNewline:$NoNewLine
                            $StartMachinesWhileWaiting = $StartMachinesWhileWaiting | Where-Object { $_ -ne $StartMachinesWhileWaiting[0] }
                        }
                    }
                }
            }
        }
        else
        {
            Start-Sleep -Seconds 1
        }

        <#
                Not implemented yet as receive-job displays everything in the console
                if ($lastMonitorJob -and ((Get-Date) - $lastMonitorJob).TotalSeconds -ge 5)
                {
                foreach ($job in $MonitorJob)
                {
                try
                {
                $dummy = Receive-Job -Keep -Id $job.ID -ErrorAction Stop
                }
                catch
                {
                Write-ScreenInfo -Message "Something went wrong with '$($job.Name)'. Please check using 'Receive-Job -Id $($job.Id)'" -Type Error
                throw 'Execution stopped'
                }
                }
                }
        #>

        foreach ($machine in $machines)
        {
            $currentMachineUptime = (Get-LWHypervVM -Name $machine.ResourceName).Uptime.TotalSeconds
            Write-Debug -Message "Uptime machine '$($machine.ResourceName)'=$currentMachineUptime, Saved uptime=$($machine.uptime)"
            if ($machine.Uptime -ne 0 -and $currentMachineUptime -lt $machine.Uptime)
            {
                Write-PSFMessage -Message "Machine '$machine' is now stopped"
                $machine.Uptime = 0
            }
        }

        Start-Sleep -Seconds 2

        if ($MonitorJob)
        {
            foreach ($job in $MonitorJob)
            {
                if ($job.State -eq 'Failed')
                {
                    $result = $job | Receive-Job -ErrorVariable jobError

                    $criticalError = $jobError | Where-Object { $_.Exception.Message -like 'AL_CRITICAL*' }
                    if ($criticalError) { throw $criticalError.Exception }

                    $nonCriticalErrors = $jobError | Where-Object { $_.Exception.Message -like 'AL_ERROR*' }
                    foreach ($nonCriticalError in $nonCriticalErrors)
                    {
                        Write-PSFMessage "There was a non-critical error in job $($job.ID) '$($job.Name)' with the message: '($nonCriticalError.Exception.Message)'"
                    }
                }
            }
        }
    }
    until (($machines.Uptime | Measure-Object -Maximum).Maximum -eq 0 -or (Get-Date).AddMinutes(-$TimeoutInMinutes) -gt $start)

    if (($machines.Uptime | Measure-Object -Maximum).Maximum -eq 0)
    {
        Write-PSFMessage -Message "All machines have stopped: ($($machines.name -join ', '))"
    }

    if ((Get-Date).AddMinutes(-$TimeoutInMinutes) -gt $start)
    {
        foreach ($Computer in $ComputerName)
        {
            if ($machineInfo.($Computer) -gt 0)
            {
                Write-Error -Message "Timeout while waiting for computer '$computer' to restart." -TargetObject $computer
            }
        }
    }

    $remainingMinutes = $TimeoutInMinutes - ((Get-Date) - $start).TotalMinutes
    Wait-LabVM -ComputerName $ComputerName -ProgressIndicator $ProgressIndicator -TimeoutInMinutes $remainingMinutes -NoNewLine:$NoNewLine

    if ($delayedStart)
    {
        Start-LabVM -ComputerName $delayedStart -NoNewline:$NoNewLine
    }

    Write-ProgressIndicatorEnd

    Write-LogFunctionExit
}


function Enable-LWVMWareVMRemoting
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseCompatibleCmdlets", "", Justification="Not relevant on Linux")]
    param(
        [Parameter(Mandatory, Position = 0)]
        $ComputerName
    )

    if ($ComputerName)
    {
        $machines = Get-LabVM -All | Where-Object Name -in $ComputerName
    }
    else
    {
        $machines = Get-LabVM -All
    }

    $script = {
        param ($DomainName, $UserName, $Password)

        $VerbosePreference = 'Continue'

        $RegPath = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon'

        Set-ItemProperty -Path $RegPath -Name AutoAdminLogon -Value 1 -ErrorAction SilentlyContinue
        Set-ItemProperty -Path $RegPath -Name DefaultUserName -Value $UserName -ErrorAction SilentlyContinue
        Set-ItemProperty -Path $RegPath -Name DefaultPassword -Value $Password -ErrorAction SilentlyContinue
        Set-ItemProperty -Path $RegPath -Name DefaultDomainName -Value $DomainName -ErrorAction SilentlyContinue

        Enable-WSManCredSSP -Role Server -Force | Out-Null
    }

    foreach ($machine in $machines)
    {
        $cred = $machine.GetCredential((Get-Lab))
        try
        {
            Invoke-LabCommand -ComputerName $machine -ActivityName SetLabVMRemoting -ScriptBlock $script `
            -ArgumentList $machine.DomainName, $cred.UserName, $cred.GetNetworkCredential().Password -ErrorAction Stop -Verbose
        }
        catch
        {
            Connect-WSMan -ComputerName $machine -Credential $cred
            Set-Item -Path "WSMan:\$machine\Service\Auth\CredSSP" -Value $true
            Disconnect-WSMan -ComputerName $machine
        }
    }
}


function Get-LWVMWareVMStatus
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseCompatibleCmdlets", "", Justification="Not relevant on Linux")]
    param (
        [Parameter(Mandatory)]
        [string[]]$ComputerName
    )

    Write-LogFunctionEntry

    $result = @{ }

    foreach ($name in $ComputerName)
    {
        $vm = VMware.VimAutomation.Core\Get-VM -Name $name
        if ($vm)
        {
            if ($vm.PowerState -eq 'PoweredOn')
            {
                $result.Add($vm.Name, 'Started')
            }
            elseif ($vm.PowerState -eq 'PoweredOff')
            {
                $result.Add($vm.Name, 'Stopped')
            }
            else
            {
                $result.Add($vm.Name, 'Unknown')
            }
        }
    }

    $result

    Write-LogFunctionExit
}


function New-LWVMWareVM
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseCompatibleCmdlets", "", Justification="Not relevant on Linux")]
    [Cmdletbinding()]
    Param (
        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter(Mandatory)]
        [string]$ReferenceVM,

        [Parameter(Mandatory)]
        [string]$AdminUserName,

        [Parameter(Mandatory)]
        [string]$AdminPassword,

        [Parameter(ParameterSetName = 'DomainJoin')]
        [string]$DomainName,

        [Parameter(Mandatory, ParameterSetName = 'DomainJoin')]
        [pscredential]$DomainJoinCredential,

        [switch]$AsJob,

        [switch]$PassThru
    )

    Write-LogFunctionEntry

    $lab = Get-Lab

    #TODO: add logic to determine if machine already exists
    <#
            if (VMware.VimAutomation.Core\Get-VM -Name $Machine.Name -ErrorAction SilentlyContinue)
            {
            Write-ProgressIndicatorEnd
            Write-ScreenInfo -Message "The machine '$Machine' does already exist" -Type Warning
            return $false
            }

            Write-Verbose "Creating machine with the name '$($Machine.Name)' in the path '$VmPath'"

    #>

    $folderName = "AutomatedLab_$($lab.Name)"
    if (-not (Get-Folder -Name $folderName -ErrorAction SilentlyContinue))
    {
        New-Folder -Name $folderName -Location VM | out-null
    }


    $referenceSnapshot = (Get-Snapshot -VM (VMware.VimAutomation.Core\Get-VM $ReferenceVM)).Name | Select-Object -last 1

    $parameters = @{
        Name = $Name
        ReferenceVM = $ReferenceVM
        AdminUserName = $AdminUserName
        AdminPassword = $AdminPassword
        DomainName = $DomainName
        DomainCred = $DomainJoinCredential
        FolderName = $FolderName
    }

    if ($AsJob)
    {
        $job = Start-Job -ScriptBlock {
            throw 'Not implemented yet'  # TODO: implement
        } -ArgumentList $parameters


        if ($PassThru)
        {
            $job
        }
    }
    else
    {
        $osSpecs = Get-OSCustomizationSpec -Name AutomatedLabSpec -Type NonPersistent -ErrorAction SilentlyContinue
        if ($osSpecs)
        {
            Remove-OSCustomizationSpec -OSCustomizationSpec $osSpecs -Confirm:$false
        }

        if (-not $parameters.DomainName)
        {
            $osSpecs = New-OSCustomizationSpec -Name AutomatedLabSpec -FullName $parameters.AdminUserName -AdminPassword $parameters.AdminPassword `
            -OSType Windows -Type NonPersistent -OrgName AutomatedLab -Workgroup AutomatedLab -ChangeSid
            #$osSpecs = Get-OSCustomizationSpec -Name Standard | Get-OSCustomizationNicMapping | Set-OSCustomizationNicMapping -IpMode UseStaticIP -IpAddress $ipaddress -SubnetMask $netmask -DefaultGateway $gateway -Dns $DNS
        }
        else
        {
            $osSpecs = New-OSCustomizationSpec -Name AutomatedLabSpec -FullName $parameters.AdminUserName -AdminPassword $parameters.AdminPassword `
            -OSType Windows -Type NonPersistent -OrgName AutomatedLab -Domain $parameters.DomainName -DomainCredentials $DomainJoinCredential -ChangeSid
        }

        $ReferenceVM_int = VMware.VimAutomation.Core\Get-VM -Name $parameters.ReferenceVM
        if (-not $ReferenceVM_int)
        {
            Write-Error "Reference VM '$($parameters.ReferenceVM)' could not be found, cannot create the machine '$($machine.Name)'"
            return
        }

        # Create Linked Clone
        $result = VMware.VimAutomation.Core\New-VM `
        -Name $parameters.Name `
        -ResourcePool $lab.VMWareSettings.ResourcePool `
        -Datastore $lab.VMWareSettings.DataStore `
        -Location (Get-Folder -Name $parameters.FolderName) `
        -OSCustomizationSpec $osSpecs `
        -VM $ReferenceVM_int `
        -LinkedClone `
        -ReferenceSnapshot $referenceSnapshot `

        #TODO: logic to switch to full clone for AD recovery scenario's etc.
        <# Create full clone
                $result = VMware.VimAutomation.Core\New-VM `
                -Name $parameters.Name `
                -ResourcePool $lab.VMWareSettings.ResourcePool `
                -Datastore $lab.VMWareSettings.DataStore `
                -Location (Get-Folder -Name $parameters.FolderName) `
                -OSCustomizationSpec $osSpecs `
                -VM $ReferenceVM_int
        #>
    }

    if ($PassThru)
    {
        $result
    }

    Write-LogFunctionExit
}


function Remove-LWVMWareVM
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseCompatibleCmdlets", "", Justification="Not relevant on Linux")]
    Param (
        [Parameter(Mandatory)]
        [string]$ComputerName,

        [switch]$AsJob,

        [switch]$PassThru
    )

    Write-LogFunctionEntry

    if ($AsJob)
    {
        $job = Start-Job -ScriptBlock {
            param (
                [Parameter(Mandatory)]
                [hashtable]$ComputerName
            )

            Add-PSSnapin -Name VMware.VimAutomation.Core, VMware.VimAutomation.Vds

            $vm = VMware.VimAutomation.Core\Get-VM -Name $ComputerName
            if ($vm)
            {
                if ($vm.PowerState -eq "PoweredOn")
                {
                    VMware.VimAutomation.Core\Stop-VM -VM $vm -Confirm:$false
                }
                VMware.VimAutomation.Core\Remove-VM -DeletePermanently -VM $ComputerName -Confirm:$false
            }
        } -ArgumentList $ComputerName


        if ($PassThru)
        {
            $job
        }
    }
    else
    {
        $vm = VMware.VimAutomation.Core\Get-VM -Name $ComputerName
        if ($vm)
        {
            if ($vm.PowerState -eq "PoweredOn")
            {
                VMware.VimAutomation.Core\Stop-VM -VM $vm -Confirm:$false
            }
            VMware.VimAutomation.Core\Remove-VM -DeletePermanently -VM $ComputerName -Confirm:$false
        }
    }

    Write-LogFunctionExit
}


function Save-LWVMWareVM
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseCompatibleCmdlets", "", Justification="Not relevant on Linux")]
    param (
        [Parameter(Mandatory)]
        [string[]]$ComputerName
    )

    Write-LogFunctionEntry

    VMware.VimAutomation.Core\Suspend-VM -VM $ComputerName -ErrorAction SilentlyContinue -Confirm:$false

    Write-LogFunctionExit
}


function Start-LWVMWareVM
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseCompatibleCmdlets", "", Justification="Not relevant on Linux")]
    param (
        [Parameter(Mandatory = $true)]
        [string[]]$ComputerName,

        [int]$DelayBetweenComputers = 0
    )

    Write-LogFunctionEntry

    foreach ($name in $ComputerName)
    {
        $vm = $null
        $vm = VMware.VimAutomation.Core\Get-VM -Name $name
        if ($vm)
        {
            VMware.VimAutomation.Core\Start-VM $vm -ErrorAction SilentlyContinue | out-null
            $result = VMware.VimAutomation.Core\Get-VM $vm
            if ($result.PowerState -ne "PoweredOn")
            {
                Write-Error "Could not start machine '$name'"
            }
        }
        Start-Sleep -Seconds $DelayBetweenComputers
    }

    Write-LogFunctionExit
}


function Stop-LWVMWareVM
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseCompatibleCmdlets", "", Justification="Not relevant on Linux")]
    param (
        [Parameter(Mandatory)]
        [string[]]$ComputerName
    )

    Write-LogFunctionEntry

    foreach ($name in $ComputerName)
    {
        if (VMware.VimAutomation.Core\Get-VM -Name $name)
        {
            $result = Shutdown-VMGuest -VM $name -ErrorAction SilentlyContinue -Confirm:$false
            if ($result.PowerState -ne "PoweredOff")
            {
                Write-Error "Could not stop machine '$name'"
            }
        }
        else
        {
            Write-ScreenInfo "The machine '$name' does not exist on the connected ESX Server" -Type Warning
        }
    }

    Write-LogFunctionExit
}


function Wait-LWVMWareRestartVM
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseCompatibleCmdlets", "", Justification="Not relevant on Linux")]
    param (
        [Parameter(Mandatory)]
        [string[]]$ComputerName,

        [double]$TimeoutInMinutes = 15
    )

    Write-LogFunctionEntry

    $prevErrorActionPreference = $Global:ErrorActionPreference
    $Global:ErrorActionPreference = 'SilentlyContinue'
    $preVerboseActionPreference = $Global:VerbosePreference
    $Global:VerbosePreference = 'SilentlyContinue'

    $start = Get-Date

    Write-PSFMessage "Starting monitoring the servers at '$start'"

    $machines = Get-LabVM -ComputerName $ComputerName

    $cmd = {
        param (
            [datetime]$Start
        )

        $events = Get-EventLog -LogName System -InstanceId 2147489653 -After $Start -Before $Start.AddHours(1)

        $events
    }

    do
    {
        $azureVmsToWait = foreach ($machine in $machines)
        {
            $events = Invoke-LabCommand -ComputerName $machine -ActivityName WaitForRestartEvent -ScriptBlock $cmd -ArgumentList $start.Ticks -UseLocalCredential -PassThru

            if ($events)
            {
                Write-PSFMessage "VM '$machine' has been restarted"
            }
            else
            {
                $machine
            }
            Start-Sleep -Seconds 15
        }
    }
    until ($azureVmsToWait.Count -eq 0 -or (Get-Date).AddMinutes(- $TimeoutInMinutes) -gt $start)

    $Global:ErrorActionPreference = $prevErrorActionPreference
    $Global:VerbosePreference = $preVerboseActionPreference

    if ((Get-Date).AddMinutes(- $TimeoutInMinutes) -gt $start)
    {
        Write-Error -Message "Timeout while waiting for computers to restart. Computers not restarted: $($azureVmsToWait.Name -join ', ')"
    }

    Write-LogFunctionExit
}


function Get-LWVMWareNetworkSwitch
{
    param (
        [Parameter(Mandatory)]
        [AutomatedLab.VirtualNetwork[]]$VirtualNetwork
    )

    Write-LogFunctionEntry

    foreach ($network in $VirtualNetwork)
    {
        $network = Get-VDPortgroup -Name $network.Name

        if (-not $network)
        {
            Write-Error "Network '$Name' is not configured"
        }

        $network
    }

    Write-LogFunctionExit
}


function Install-LWLabCAServers
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseCompatibleCmdlets", "", Justification="Not relevant on Linux")]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingWriteHost", "", Justification="Historic cmdlet, will not be updated")]
    param (
        [Parameter(Mandatory = $true)][string]$ComputerName,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$DomainName,
        [Parameter(Mandatory = $true)][string]$UserName,
        [Parameter(Mandatory = $true)][string]$Password,
        [Parameter(Mandatory = $false)][string]$ForestAdminUserName,
        [Parameter(Mandatory = $false)][string]$ForestAdminPassword,
        [Parameter(Mandatory = $false)][string]$ParentCA,
        [Parameter(Mandatory = $false)][string]$ParentCALogicalName,
        [Parameter(Mandatory = $true)][string]$CACommonName,
        [Parameter(Mandatory = $true)][string]$CAType,
        [Parameter(Mandatory = $true)][string]$KeyLength,
        [Parameter(Mandatory = $true)][string]$CryptoProviderName,
        [Parameter(Mandatory = $true)][string]$HashAlgorithmName,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$DatabaseDirectory,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$LogDirectory,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$CpsUrl,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$CpsText,
        [Parameter(Mandatory = $true)][boolean]$UseLDAPAIA,
        [Parameter(Mandatory = $true)][boolean]$UseHTTPAia,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$AIAHTTPURL01,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$AiaHttpUrl02,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$AIAHTTPURL01UploadLocation,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$AiaHttpUrl02UploadLocation,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$OCSPHttpUrl01,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$OCSPHttpUrl02,
        [Parameter(Mandatory = $true)][boolean]$UseLDAPCRL,
        [Parameter(Mandatory = $true)][boolean]$UseHTTPCRL,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$CDPHTTPURL01,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$CDPHTTPURL02,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$CDPHTTPURL01UploadLocation,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$CDPHTTPURL02UploadLocation,
        [Parameter(Mandatory = $true)][boolean]$InstallOCSP,
        [Parameter(Mandatory = $false)][string]$ValidityPeriod,
        [Parameter(Mandatory = $false)][int]$ValidityPeriodUnits,
        [Parameter(Mandatory = $true)][string]$CRLPeriod,
        [Parameter(Mandatory = $true)][int]$CRLPeriodUnits,
        [Parameter(Mandatory = $true)][string]$CRLOverlapPeriod,
        [Parameter(Mandatory = $true)][int]$CRLOverlapUnits,
        [Parameter(Mandatory = $true)][string]$CRLDeltaPeriod,
        [Parameter(Mandatory = $true)][int]$CRLDeltaPeriodUnits,
        [Parameter(Mandatory = $true)][string]$CertsValidityPeriod,
        [Parameter(Mandatory = $true)][int]$CertsValidityPeriodUnits,
        [Parameter(Mandatory = $true)][boolean]$InstallWebEnrollment,
        [Parameter(Mandatory = $true)][boolean]$InstallWebRole,
        [Parameter(Mandatory = $true)][boolean]$DoNotLoadDefaultTemplates,
        [Parameter(Mandatory = $false)][int]$PreDelaySeconds
    )

    Write-LogFunctionEntry

    Install-LabWindowsFeature -ComputerName $ComputerName -FeatureName RSAT-AD-Tools -IncludeAllSubFeature -NoDisplay

    #region - Create parameter table
    $param = @{ }
    $param.Add('ComputerName', $ComputerName)
    $param.add('DomainName', $DomainName)

    $param.Add('UserName', $UserName)
    $param.Add('Password', $Password)
    $param.Add('ForestAdminUserName', $ForestAdminUserName)
    $param.Add('ForestAdminPassword', $ForestAdminPassword)

    $param.Add('CACommonName', $CACommonName)

    $param.Add('CAType', $CAType)

    $param.Add('CryptoProviderName', $CryptoProviderName)
    $param.Add('HashAlgorithmName', $HashAlgorithmName)

    $param.Add('KeyLength', $KeyLength)

    $param.Add('CertEnrollFolderPath', $CertEnrollFolderPath)
    $param.Add('DatabaseDirectory', $DatabaseDirectory)
    $param.Add('LogDirectory', $LogDirectory)

    $param.Add('CpsUrl', $CpsUrl)
    $param.Add('CpsText', """$($CpsText)""")

    $param.Add('UseLDAPAIA', $UseLDAPAIA)
    $param.Add('UseHTTPAia', $UseHTTPAia)
    $param.Add('AIAHTTPURL01', $AIAHTTPURL01)
    $param.Add('AiaHttpUrl02', $AiaHttpUrl02)
    $param.Add('AIAHTTPURL01UploadLocation', $AIAHTTPURL01UploadLocation)
    $param.Add('AiaHttpUrl02UploadLocation', $AiaHttpUrl02UploadLocation)

    $param.Add('OCSPHttpUrl01', $OCSPHttpUrl01)
    $param.Add('OCSPHttpUrl02', $OCSPHttpUrl02)

    $param.Add('UseLDAPCRL', $UseLDAPCRL)
    $param.Add('UseHTTPCRL', $UseHTTPCRL)
    $param.Add('CDPHTTPURL01', $CDPHTTPURL01)
    $param.Add('CDPHTTPURL02', $CDPHTTPURL02)
    $param.Add('CDPHTTPURL01UploadLocation', $CDPHTTPURL01UploadLocation)
    $param.Add('CDPHTTPURL02UploadLocation', $CDPHTTPURL02UploadLocation)

    $param.Add('InstallOCSP', $InstallOCSP)

    $param.Add('ValidityPeriod', $ValidityPeriod)
    $param.Add('ValidityPeriodUnits', $ValidityPeriodUnits)
    $param.Add('CRLPeriod', $CRLPeriod)
    $param.Add('CRLPeriodUnits', $CRLPeriodUnits)
    $param.Add('CRLOverlapPeriod', $CRLOverlapPeriod)
    $param.Add('CRLOverlapUnits', $CRLOverlapUnits)
    $param.Add('CRLDeltaPeriod', $CRLDeltaPeriod)
    $param.Add('CRLDeltaPeriodUnits', $CRLDeltaPeriodUnits)
    $param.Add('CertsValidityPeriod', $CertsValidityPeriod)
    $param.Add('CertsValidityPeriodUnits', $CertsValidityPeriodUnits)

    $param.Add('InstallWebEnrollment', $InstallWebEnrollment)

    $param.Add('InstallWebRole', $InstallWebRole)

    $param.Add('DoNotLoadDefaultTemplates', $DoNotLoadDefaultTemplates)

    #For Subordinate CAs only
    if ($ParentCA) { $param.add('ParentCA', $ParentCA) }
    if ($ParentCALogicalname) { $param.add('ParentCALogicalname', $ParentCALogicalName) }

    $param.Add('PreDelaySeconds', $PreDelaySeconds)
    #endregion - Create parameter table


    #region - Parameters debug
    Write-Debug -Message '---------------------------------------------------------------------------------------'
    Write-Debug -Message 'Parameters - Entered Install-LWLabCAServers'
    Write-Debug -Message '---------------------------------------------------------------------------------------'
    if ($param.GetEnumerator().count)
    {
        foreach ($key in ($param.GetEnumerator() | Sort-Object -Property Name)) { Write-Debug -message "  $($key.key.padright(27)) $($key.value)" }
    }
    else
    {
        Write-Debug -message '  No parameters specified'
    }
    Write-Debug -Message '---------------------------------------------------------------------------------------'
    Write-Debug -Message ''
    #endregion - Parameters debug




    #region ScriptBlock for installation
    $caScriptBlock = {

        param ($param)
        $deployDebug = (New-Item -Force -ItemType Directory -Path $ExecutionContext.InvokeCommand.ExpandString($AL_DeployDebugFolder)).FullName
        $param | Export-Clixml $deployDebug\CaParams.xml

        #Make semi-sure that each install of CA server is not done at the same time
        Start-Sleep -Seconds $param.PreDelaySeconds

        Import-Module -Name ServerManager

        #region - Check if CA is already installed
        if ((Get-WindowsFeature -Name 'ADCS-Cert-Authority').Installed)
        {
            Write-Output "A Certificate Authority is already installed on '$($param.ComputerName)'. Skipping installation."
            return
        }
        #endregion

        #region - Create CAPolicy file
        $caPolicyFileName = "$Env:Windir\CAPolicy.inf"
        if (-not (Test-Path -Path $caPolicyFileName))
        {
            Write-Verbose -Message 'Create CAPolicy.inf file'
            Set-Content $caPolicyFileName -Force -Value ';CAPolicy for CA'
            Add-Content $caPolicyFileName -Value '; Please replace sample CPS OID with your own OID'
            Add-Content $caPolicyFileName -Value ''
            Add-Content $caPolicyFileName -Value '[Version]'
            Add-Content $caPolicyFileName -Value "Signature=`"`$Windows NT`$`" "
            Add-Content $caPolicyFileName -Value ''
            Add-Content $caPolicyFileName -Value '[PolicyStatementExtension]'
            Add-Content $caPolicyFileName -Value 'Policies=LegalPolicy'
            Add-Content $caPolicyFileName -Value 'Critical=0'
            Add-Content $caPolicyFileName -Value ''
            Add-Content $caPolicyFileName -Value '[LegalPolicy]'
            Add-Content $caPolicyFileName -Value 'OID=1.3.6.1.4.1.11.21.43'
            Add-Content $caPolicyFileName -Value "Notice=$($param.CpsText)"
            Add-Content $caPolicyFileName -Value "URL=$($param.CpsUrl)"
            Add-Content $caPolicyFileName -Value ''
            Add-Content $caPolicyFileName -Value '[Certsrv_Server]'
            Add-Content $caPolicyFileName -Value 'ForceUTF8=true'
            Add-Content $caPolicyFileName -Value "RenewalKeyLength=$($param.KeyLength)"
            Add-Content $caPolicyFileName -Value "RenewalValidityPeriod=$($param.ValidityPeriod)"
            Add-Content $caPolicyFileName -Value "RenewalValidityPeriodUnits=$($param.ValidityPeriodUnits)"
            Add-Content $caPolicyFileName -Value "CRLPeriod=$($param.CRLPeriod)"
            Add-Content $caPolicyFileName -Value "CRLPeriodUnits=$($param.CRLPeriodUnits)"
            Add-Content $caPolicyFileName -Value "CRLDeltaPeriod=$($param.CRLDeltaPeriod)"
            Add-Content $caPolicyFileName -Value "CRLDeltaPeriodUnits=$($param.CRLDeltaPeriodUnits)"
            Add-Content $caPolicyFileName -Value 'EnableKeyCounting=0'
            Add-Content $caPolicyFileName -Value 'AlternateSignatureAlgorithm=0'
            if ($param.DoNotLoadDefaultTemplates) { Add-Content $caPolicyFileName -Value 'LoadDefaultTemplates=0' }
            if ($param.CAType -like '*root*')
            {
                Add-Content $caPolicyFileName -Value ''
                Add-Content $caPolicyFileName -Value '[Extensions]'
                Add-Content $caPolicyFileName -Value ';Remove CA Version Index'
                Add-Content $caPolicyFileName -Value '1.3.6.1.4.1.311.21.1='
                Add-Content $caPolicyFileName -Value ';Remove CA Hash of previous CA Certificates'
                Add-Content $caPolicyFileName -Value '1.3.6.1.4.1.311.21.2='
                Add-Content $caPolicyFileName -Value ';Remove V1 Certificate Template Information'
                Add-Content $caPolicyFileName -Value '1.3.6.1.4.1.311.20.2='
                Add-Content $caPolicyFileName -Value ';Remove CA of V2 Certificate Template Information'
                Add-Content $caPolicyFileName -Value '1.3.6.1.4.1.311.21.7='
                Add-Content $caPolicyFileName -Value ';Key Usage Attribute set to critical'
                Add-Content $caPolicyFileName -Value '2.5.29.15=AwIBBg=='
                Add-Content $caPolicyFileName -Value 'Critical=2.5.29.15'
            }

            if ($param.DebugPref -eq 'Continue')
            {
                $file = get-content -Path "$Env:Windir\CAPolicy.inf"
                Write-Debug -Message 'CApolicy.inf contents:'
                foreach ($line in $file)
                {
                    Write-Debug -Message $line
                }
            }
        }
        #endregion - Create CAPolicy file


        #region - Install CA
        $hostOSVersion = [Environment]::OSVersion.Version
        if ($hostOSVersion -ge [system.version]'6.2')
        {
            $InstallFeatures = 'Import-Module -Name ServerManager; Add-WindowsFeature -IncludeManagementTools -Name ADCS-Cert-Authority'
        }
        else
        {
            $InstallFeatures = 'Import-Module -Name ServerManager; Add-WindowsFeature -Name ADCS-Cert-Authority'
        }
        # OCSP not yet supported
        #if ($param.InstallOCSP)          { $InstallFeatures += ", ADCS-Online-Cert" }
        if ($param.InstallWebEnrollment) { $InstallFeatures += ', ADCS-Web-Enrollment' }



        if ($param.ForestAdminUserName)
        {
            Write-Verbose -Message "ForestAdminUserName=$($param.ForestAdminUserName), ForestAdminPassword=$($param.ForestAdminPassword)"

            if ($param.DebugPref -eq 'Continue')
            {
                Write-Verbose -Message "Adding $($param.ForestAdminUserName) to local administrators group"
                Write-Verbose -Message "WinNT:://$($param.ForestAdminUserName.replace('\', '/'))"
            }
            $localGroup = ([ADSI]'WinNT://./Administrators,group')
            $localGroup.psbase.Invoke('Add', ([ADSI]"WinNT://$($param.ForestAdminUserName.replace('\', '/'))").path)
            Write-Verbose -Message "Check 2c -create credential of ""$($param.ForestAdminUserName)"" and ""$($param.ForestAdminPassword)"""
            $forestAdminCred = (New-Object System.Management.Automation.PSCredential($param.ForestAdminUserName, ($param.ForestAdminPassword | ConvertTo-SecureString -AsPlainText -Force)))
        }
        else
        {
            Write-Verbose -Message 'No ForestAdminUserName!'
        }




        Write-Verbose -Message 'Installing roles and features now'
        Write-Verbose -Message "Command: $InstallFeatures"
        Invoke-Expression -Command ($InstallFeatures += " -Confirm:`$false") | Out-Null

        Write-Verbose -Message 'Installing ADCS now'
        $installCommand = 'Install-AdcsCertificationAuthority '
        $installCommand += "-CACommonName                ""$($param.CACommonName)"" "
        $installCommand += "-CAType                      $($param.CAType) "
        $installCommand += "-KeyLength                   $($param.KeyLength) "
        $installCommand += "-CryptoProviderName          ""$($param.CryptoProviderName)"" "
        $installCommand += "-HashAlgorithmName           ""$($param.HashAlgorithmName)"" "
        $installCommand += '-OverwriteExistingKey '
        $installCommand += '-OverwriteExistingDatabase '
        $installCommand += '-Force '
        $installCommand += '-Confirm:$false '
        if ($forestAdminCred) { $installCommand += '-Credential $forestAdminCred ' }

        if ($param.DatabaseDirectory) { $installCommand += "-DatabaseDirectory      $($param.DatabaseDirectory) " }
        if ($param.LogDirectory)      { $installCommand += "-LogDirectory           $($param.LogDirectory) " }

        if ($param.CAType -like '*root*')
        {
            $installCommand += "-ValidityPeriod          $($param.ValidityPeriod) "
            $installCommand += "-ValidityPeriodUnits     $($param.ValidityPeriodUnits) "
        }
        else
        {
            $installCommand += "-ParentCA                $($param.ParentCA)`\$($param.ParentCALogicalName) "
        }
        $installCommand += ' | Out-Null'

        if ($param.DebugPref -eq 'Continue')
        {
            Write-Debug -Message 'Install command:'
            Write-Debug -Message $installCommand
            Set-Content -Path "$deployDebug\debug-CAinst.txt" -value $installCommand
        }


        Invoke-Expression -Command $installCommand


        if ($param.ForestAdminUserName)
        {
            if ($param.DebugPref -eq 'Continue')
            {
                Write-Debug -Message "Removing $($param.ForestAdminUserName) to local administrators group"
            }
            $localGroup = ([ADSI]'WinNT://./Administrators,group')
            $localGroup.psbase.Invoke('Remove', ([ADSI]"WinNT://$($param.ForestAdminUserName.replace('\', '/'))").path)
        }


        if ($param.InstallWebEnrollment)
        {
            Write-Verbose -Message 'Installing Web Enrollment service now'
            Install-ADCSWebEnrollment -Confirm:$False | Out-Null
        }

        if ($param.InstallWebRole)
        {
            if (!(Get-WindowsFeature -Name 'web-server'))
            {
                Add-WindowsFeature -Name 'Web-Server' -IncludeManagementTools

                #Allow "+" characters in URL for supporting delta CRLs
                Set-WebConfiguration -Filter system.webServer/security/requestFiltering -PSPath 'IIS:\sites\Default Web Site' -Value @{allowDoubleEscaping=$true}
            }
        }
        #endregion - Install CA

        #region - Configure IIS virtual directories
        if ($param.UseHTTPAia)
        {
            New-WebVirtualDirectory -Site 'Default Web Site' -Name Aia -PhysicalPath 'C:\Windows\System32\CertSrv\CertEnroll' | Out-Null
            New-WebVirtualDirectory -Site 'Default Web Site' -Name Cdp -PhysicalPath 'C:\Windows\System32\CertSrv\CertEnroll' | Out-Null
        }
        #endregion - Configure IIS virtual directories

        #region - Configure OCSP
        <# OCSP not yet supported
                if ($InstallOCSP)
                {
                Write-Verbose -Message "Installing Online Responder"
                Install-ADCSOnlineResponder -Force | Out-Null
                }
        #>
        #endregion - Configure OCSP







        #region - Configure CA
        function Invoke-CustomExpression
        {
            param ($Command)

            Write-Host $command
            Invoke-Expression -Command $command
        }


        #Declare configuration NC
        if ($param.CAType -like 'Enterprise*')
        {
            $lDAPname = ''
            foreach ($part in ($param.DomainName.split('.')))
            {
                $lDAPname += ",DC=$part"
            }
            Invoke-CustomExpression -Command "certutil -setreg CA\DSConfigDN ""CN=Configuration$lDAPname"""
        }

        #Apply the required CDP Extension URLs
        $command = "certutil -setreg CA\CRLPublicationURLs ""1:$($Env:WinDir)\system32\CertSrv\CertEnroll\%3%8%9.crl"
        if ($param.UseLDAPCRL) { $command += '\n11:ldap:///CN=%7%8,CN=%2,CN=CDP,CN=Public Key Services,CN=Services,%6%10' }
        if ($param.UseHTTPCRL) { $command += "\n2:$($param.CDPHTTPURL01)/%3%8%9.crl" }
        if ($param.CDPHTTPURL01UploadLocation) { $command += "\n1:$($param.CDPHTTPURL01UploadLocation)/%3%8%9.crl" }
        $command += '"'
        Invoke-CustomExpression -Command $command

        #Apply the required AIA Extension URLs
        $command = "certutil -setreg CA\CACertPublicationURLs ""1:$($Env:WinDir)\system3\CertSrv\CertEnroll\%1_%3%4.crt"
        if ($param.UseLDAPAia) { $command += '\n3:ldap:///CN=%7,CN=AIA,CN=Public Key Services,CN=Services,%6%11' }
        if ($param.UseHTTPAia) { $command += "\n2:$($param.AIAHTTPURL01)/%1_%3%4.crt" }
        if ($param.AIAHTTPURL01UploadLocation) { $command += "\n1:$($param.AIAHTTPURL01UploadLocation)/%3%8%9.crl" }
        <# OCSP not yet supported
                if ($param.InstallOCSP -and $param.OCSPHttpUrl01) { $Line += "\n34:$($param.OCSPHttpUrl01)" }
                if ($param.InstallOCSP -and $param.OCSPHttpUrl02) { $Line += "\n34:$($param.OCSPHttpUrl02)" }
        #>
        $command += '"'
        Invoke-CustomExpression -Command $command

        #Define default maximum certificate lifetime for issued certificates
        Invoke-CustomExpression -Command "certutil -setreg ca\ValidityPeriodUnits $($param.CertsValidityPeriodUnits)"
        Invoke-CustomExpression -Command "certutil -setreg ca\ValidityPeriod ""$($param.CertsValidityPeriod)"""

        #Define CRL Publication Intervals
        Invoke-CustomExpression -Command "certutil -setreg CA\CRLPeriodUnits $($param.CRLPeriodUnits)"
        Invoke-CustomExpression -Command "certutil -setreg CA\CRLPeriod ""$($param.CRLPeriod)"""

        #Define CRL Overlap
        Invoke-CustomExpression -Command "certutil -setreg CA\CRLOverlapUnits $($param.CRLOverlapUnits)"
        Invoke-CustomExpression -Command "certutil -setreg CA\CRLOverlapPeriod ""$($param.CRLOverlapPeriod)"""

        #Define Delta CRL
        Invoke-CustomExpression -Command "certutil -setreg CA\CRLDeltaUnits $($param.CRLDeltaPeriodUnits)"
        Invoke-CustomExpression -Command "certutil -setreg CA\CRLDeltaPeriod ""$($param.CRLDeltaPeriod)"""

        #Enable Auditing Logging
        Invoke-CustomExpression -Command 'certutil -setreg CA\Auditfilter 0x7F'

        #Enable UTF-8 Encoding
        Invoke-CustomExpression -Command 'certutil -setreg ca\forceteletex +0x20'

        if ($param.CAType -like '*root*')
        {
            #Disable Discrete Signatures in Subordinate Certificates (WinXP KB968730)
            Invoke-CustomExpression -Command 'certutil -setreg CA\csp\AlternateSignatureAlgorithm 0'

            #Force digital signature removal in KU for cert issuance (see also kb888180)
            Invoke-CustomExpression -Command 'certutil -setreg policy\EditFlags -EDITF_ADDOLDKEYUSAGE'

            #Enable SAN
            Invoke-CustomExpression -Command 'certutil -setreg policy\EditFlags +EDITF_ATTRIBUTESUBJECTALTNAME2'

            #Configure policy module to automatically issue certificates when requested
            Invoke-CustomExpression -Command 'certutil -setreg ca\PolicyModules\CertificateAuthority_MicrosoftDefault.Policy\RequestDisposition 1'
        }
        #If CA is Root CA and Sub CAs are present, disable (do not publish) templates (except SubCA template)
        if ($param.DoNotLoadDefaultTemplates)
        {
            Invoke-CustomExpression -Command 'certutil -SetCATemplates +SubCA'
        }
        #endregion - Configure CA





        #region - Restart of CA
        if ((Get-Service -Name 'CertSvc').Status -eq 'Running')
        {
            Write-Verbose -Message 'Stopping ADCS Service'
            $totalretries = 5
            $retries = 0
            do
            {
                Stop-Service -Name 'CertSvc' -ErrorAction SilentlyContinue
                if ((Get-Service -Name 'CertSvc').Status -ne 'Stopped')
                {
                    $retries++
                    Start-Sleep -Seconds 1
                }
            }
            until (((Get-Service -Name 'CertSvc').Status -eq 'Stopped') -or ($retries -ge $totalretries))

            if ((Get-Service -Name 'CertSvc').Status -eq 'Stopped')
            {
                Write-Verbose -Message 'ADCS service is now stopped'
            }
            else
            {
                Write-Error -Message 'Could not stop ADCS Service after several retries'
                return
            }
        }

        Write-Verbose -Message 'Starting ADCS Service now'
        $totalretries = 5
        $retries = 0
        do
        {
            Start-Service -Name 'CertSvc' -ErrorAction SilentlyContinue
            if ((Get-Service -Name 'CertSvc').Status -ne 'Running')
            {
                $retries++
                Start-Sleep -Seconds 1
            }
        }
        until (((Get-Service -Name 'CertSvc').Status -eq 'Running') -or ($retries -ge $totalretries))

        if ((Get-Service -Name 'CertSvc').Status -eq 'Running')
        {
            Write-Verbose -Message 'ADCS service is now started'
        }
        else
        {
            Write-Error -Message 'Could not start ADCS Service after several retries'
            return
        }
        #endregion - Restart of CA


        Write-Verbose -Message 'Waiting for admin interface to be ready'
        $totalretries = 10
        $retries = 0
        do
        {
            $result = Invoke-Expression -Command "certutil -pingadmin .\$($param.CACommonName)"
            if (!($result | Where-Object { $_ -like '*interface is alive*' }))
            {
                $retries++
                Write-Verbose -Message "Admin interface not ready. Check $retries of $totalretries"
                if ($retries -lt $totalretries) { Start-Sleep -Seconds 10 }
            }
        }
        until (($result | Where-Object { $_ -like '*interface is alive*' }) -or ($retries -ge $totalretries))

        if ($result | Where-Object { $_ -like '*interface is alive*' })
        {
            Write-Verbose -Message 'Admin interface is now ready'
        }
        else
        {
            Write-Error -Message 'Admin interface was not ready after several retries'
            return
        }


        #region - Issue of CRL
        Start-Sleep -Seconds 2
        Invoke-Expression -Command 'certutil -crl' | Out-Null
        $totalretries = 12
        $retries = 0
        do
        {
            Start-Sleep -Seconds 5
            $retries++
        }
        until ((Get-ChildItem "$env:systemroot\system32\CertSrv\CertEnroll\*.crl") -or ($retries -ge $totalretries))

        #endregion - Issue of CRL

        if (($param.CAType -like 'Enterprise*') -and ($param.DoNotLoadDefaultTemplates)) { Invoke-Expression 'certutil -SetCATemplates +SubCA' }
    }

    #endregion

    Write-PSFMessage -Message "Performing installation of $($param.CAType) on '$($param.ComputerName)'"
    $job = Invoke-LabCommand -ActivityName "Install CA on '$($param.Computername)'" -ComputerName $param.ComputerName -Scriptblock $caScriptBlock -ArgumentList $param -NoDisplay -AsJob -PassThru -Variable (Get-Variable -Name AL_DeployDebugFolder -Scope Global)

    $job

    Write-LogFunctionExit
}


function Install-LWLabCAServers2008
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseCompatibleCmdlets", "", Justification="Not relevant on Linux")]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingWriteHost", "", Justification="Historic cmdlet, will not be updated")]
    [Cmdletbinding()]
    param (
        [Parameter(Mandatory)]
        [hashtable]$param
    )

    Write-LogFunctionEntry

    #region - Parameters debug
    Write-Debug -Message '---------------------------------------------------------------------------------------'
    Write-Debug -Message 'Parameters - Entered Install-LWLabCAServers'
    Write-Debug -Message '---------------------------------------------------------------------------------------'
    if ($param.GetEnumerator().count)
    {
        foreach ($key in ($param.GetEnumerator() | Sort-Object -Property Name)) { Write-Debug -message "  $($key.key.padright(27)) $($key.value)" }
    }
    else
    {
        Write-Debug -message '  No parameters specified'
    }
    Write-Debug -Message '---------------------------------------------------------------------------------------'
    Write-Debug -Message ''
    #endregion - Parameters debug




    #region ScriptBlock for installation
    $caScriptBlock = {

        param ($param)

        function Install-WebEnrollment
        {
            [CmdletBinding()]

            param
            (
                [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
                [string]$CAConfig
            )

            # check if web enrollment binaries are installed
            Import-Module ServerManager

            # instanciate COM object
            try
            {
                $EWPSetup = New-Object -ComObject CertOCM.CertSrvSetup.1
            }
            catch
            {
                Write-ScreenInfo "Unable to load necessary interfaces. Your Windows Server operating system is not supported!" -Type Warning
                return
            }

            # initialize the object to install only web enrollment
            $EWPSetup.InitializeDefaults($false,$true)
            try
            {
                # set required information and install the role
                $EWPSetup.SetWebCAInformation($CAConfig)
                $EWPSetup.Install()
            }
            catch
            {
                $_
                return
            }
            Write-Host "Successfully installed Enrollment Web Pages on local computer!" -ForegroundColor Green
        }

        Import-Module -Name ServerManager

        #region - Check if CA is already installed
        Write-Verbose -Message 'Check if ADCS-Cert-Authority is already installed'
        if ((Get-WindowsFeature -Name 'ADCS-Cert-Authority').Installed)
        {
            Write-Verbose -Message 'ADCS-Cert-Authority is already installed. Returning'
            Write-Output "A Certificate Authority is already installed on '$($param.ComputerName)'. Skipping installation."
            return
        }
        #endregion

        #region - Create CAPolicy file
        $caPolicyFileName = "$Env:Windir\CAPolicy.inf"
        if (-not (Test-Path -Path $caPolicyFileName))
        {
            Write-Verbose -Message 'Create CAPolicy.inf file'
            Set-Content $caPolicyFileName -Force -Value ';CAPolicy for CA'
            Add-Content $caPolicyFileName -Value '; Please replace sample CPS OID with your own OID'
            Add-Content $caPolicyFileName -Value ''
            Add-Content $caPolicyFileName -Value '[Version]'
            Add-Content $caPolicyFileName -Value "Signature=`"`$Windows NT`$`" "
            Add-Content $caPolicyFileName -Value ''
            Add-Content $caPolicyFileName -Value '[PolicyStatementExtension]'
            Add-Content $caPolicyFileName -Value 'Policies=LegalPolicy'
            Add-Content $caPolicyFileName -Value 'Critical=0'
            Add-Content $caPolicyFileName -Value ''
            Add-Content $caPolicyFileName -Value '[LegalPolicy]'
            Add-Content $caPolicyFileName -Value 'OID=1.3.6.1.4.1.11.21.43'
            Add-Content $caPolicyFileName -Value "Notice=$($param.CpsText)"
            Add-Content $caPolicyFileName -Value "URL=$($param.CpsUrl)"
            Add-Content $caPolicyFileName -Value ''
            Add-Content $caPolicyFileName -Value '[Certsrv_Server]'
            Add-Content $caPolicyFileName -Value 'ForceUTF8=true'
            Add-Content $caPolicyFileName -Value "RenewalKeyLength=$($param.KeyLength)"
            Add-Content $caPolicyFileName -Value "RenewalValidityPeriod=$($param.ValidityPeriod)"
            Add-Content $caPolicyFileName -Value "RenewalValidityPeriodUnits=$($param.ValidityPeriodUnits)"
            Add-Content $caPolicyFileName -Value "CRLPeriod=$($param.CRLPeriod)"
            Add-Content $caPolicyFileName -Value "CRLPeriodUnits=$($param.CRLPeriodUnits)"
            Add-Content $caPolicyFileName -Value "CRLDeltaPeriod=$($param.CRLDeltaPeriod)"
            Add-Content $caPolicyFileName -Value "CRLDeltaPeriodUnits=$($param.CRLDeltaPeriodUnits)"
            Add-Content $caPolicyFileName -Value 'EnableKeyCounting=0'
            Add-Content $caPolicyFileName -Value 'AlternateSignatureAlgorithm=0'
            if ($param.DoNotLoadDefaultTemplates -eq 'True') { Add-Content $caPolicyFileName -Value 'LoadDefaultTemplates=0' }
            if ($param.CAType -like '*root*')
            {
                Add-Content $caPolicyFileName -Value ''
                Add-Content $caPolicyFileName -Value '[Extensions]'
                Add-Content $caPolicyFileName -Value ';Remove CA Version Index'
                Add-Content $caPolicyFileName -Value '1.3.6.1.4.1.311.21.1='
                Add-Content $caPolicyFileName -Value ';Remove CA Hash of previous CA Certificates'
                Add-Content $caPolicyFileName -Value '1.3.6.1.4.1.311.21.2='
                Add-Content $caPolicyFileName -Value ';Remove V1 Certificate Template Information'
                Add-Content $caPolicyFileName -Value '1.3.6.1.4.1.311.20.2='
                Add-Content $caPolicyFileName -Value ';Remove CA of V2 Certificate Template Information'
                Add-Content $caPolicyFileName -Value '1.3.6.1.4.1.311.21.7='
                Add-Content $caPolicyFileName -Value ';Key Usage Attribute set to critical'
                Add-Content $caPolicyFileName -Value '2.5.29.15=AwIBBg=='
                Add-Content $caPolicyFileName -Value 'Critical=2.5.29.15'
            }

            if ($param.DebugPref -eq 'Continue')
            {
                $file = get-content -Path "$Env:Windir\CAPolicy.inf"
                Write-Debug -Message 'CApolicy.inf contents:'
                foreach ($line in $file)
                {
                    Write-Debug -Message $line
                }
            }
        }
        #endregion - Create CAPolicy file


        #region - Install CA
        $hostOSVersion = [Environment]::OSVersion.Version
        if ($hostOSVersion -ge [system.version]'6.2')
        {
            $InstallFeatures = 'Import-Module -Name ServerManager; Add-WindowsFeature -IncludeManagementTools -Name ADCS-Cert-Authority'
        }
        else
        {
            $InstallFeatures = 'Import-Module -Name ServerManager; Add-WindowsFeature -Name ADCS-Cert-Authority'
        }
        # OCSP not yet supported
        #if ($param.InstallOCSP)          { $InstallFeatures += ", ADCS-Online-Cert" }
        if ($param.InstallWebEnrollment) { $InstallFeatures += ', ADCS-Web-Enrollment' }

        Write-Verbose -Message "Install roles and feature using command '$InstallFeatures'"
        Invoke-Expression -Command ($InstallFeatures += " -Confirm:`$false") | Out-Null

        if ($param.ForestAdminUserName)
        {
            Write-Verbose -Message "ForestAdminUserName=$($param.ForestAdminUserName), ForestAdminPassword=$($param.ForestAdminPassword)"

            Write-Verbose -Message "Adding $($param.ForestAdminUserName) to local administrators group"
            Write-Verbose -Message "WinNT:://$($param.ForestAdminUserName.replace('\', '/'))"
            $localGroup = ([ADSI]'WinNT://./Administrators,group')
            $localGroup.psbase.Invoke('Add', ([ADSI]"WinNT://$($param.ForestAdminUserName.replace('\', '/'))").path)
            $forestAdminCred = (New-Object System.Management.Automation.PSCredential($param.ForestAdminUserName, ($param.ForestAdminPassword | ConvertTo-SecureString -AsPlainText -Force)))
        }
        else
        {
            Write-Verbose -Message 'No ForestAdminUserName!'
        }





        try
        {
            $CASetup = New-Object -ComObject CertOCM.CertSrvSetup.1
        }
        catch
        {
            Write-Verbose -Message "Unable to load necessary interfaces. Operating system is not supported for PKI."
            return
        }

        try
        {
            $CASetup.InitializeDefaults($true, $false)
        }
        catch
        {
            Write-Verbose -Message "Cannot initialize setup binaries!"
        }


        $CATypesByVal = @{}
        $CATypesByName.keys | ForEach-Object {$CATypesByVal.Add($CATypesByName[$_],$_)}
        $CAPRopertyByName = @{"CAType"=0
            "CAKeyInfo"=1
            "Interactive"=2
            "ValidityPeriodUnits"=5
            "ValidityPeriod"=6
            "ExpirationDate"=7
            "PreserveDataBase"=8
            "DBDirectory"=9
            "Logdirectory"=10
            "ParentCAMachine"=12
            "ParentCAName"=13
            "RequestFile"=14
            "WebCAMachine"=15
        "WebCAName"=16}
        $CAPRopertyByVal = @{}
        $CAPRopertyByName.keys | ForEach-Object `
        {
            $CAPRopertyByVal.Add($CAPRopertyByName[$_],$_)
        }
        $ValidityUnitsByName = @{"years" = 6}
        $ValidityUnitsByVal = @{6 = "years"}

        $ofs = ", "



        #key length and hashing algorithm verification
        $CAKey = $CASetup.GetCASetupProperty(1)
        if ($param.CryptoProviderName -ne "")
        {
            if ($CASetup.GetProviderNameList() -notcontains $param.CryptoProviderName)
            {
                # TODO add available CryptoProviderName list
                Write-Host "Specified CSP '$param.CryptoProviderName' is not valid!"
            }
            else
            {
                $CAKey.ProviderName = $param.CryptoProviderName
            }
        }
        else
        {
            $CAKey.ProviderName = "RSA#Microsoft Software Key Storage Provider"
        }
        Write-Verbose -Message "ProviderName = '$($CAKey.ProviderName)'"


        if ($param.KeyLength -ne 0)
        {
            if ($CASetup.GetKeyLengthList($param.CryptoProviderName).Length -eq 1)
            {
                $CAKey.Length = $CASetup.GetKeyLengthList($param.CryptoProviderName)[0]
            }
            else
            {
                if ($CASetup.GetKeyLengthList($param.CryptoProviderName) -notcontains $param.KeyLength)
                {
                    Write-Host "The specified key length '$KeyLength' is not supported by the selected CryptoProviderName '$param.CryptoProviderName'"
                    Write-Host "The following key lengths are supported by this CryptoProviderName:"
                    foreach ($provider in ($CASetup.GetKeyLengthList($param.CryptoProviderName)))
                    {
                        Write-Host " $provider"
                    }
                }
                $CAKey.Length = $param.KeyLength
            }
        }
        Write-Verbose -Message "KeyLength = '$($CAKey.KeyLength)'"


        if ($param.HashAlgorithmName -ne "")
        {
            if ($CASetup.GetHashAlgorithmList($param.CryptoProviderName) -notcontains $param.HashAlgorithmName)
            {
                Write-ScreenInfo -Message "The specified hash algorithm is not supported by the selected CryptoProviderName '$param.CryptoProviderName'"
                Write-ScreenInfo -Message "The following hash algorithms are supported by this CryptoProviderName:" -Type Error
                foreach ($algorithm in ($CASetup.GetHashAlgorithmList($param.CryptoProviderName)))
                {
                    Write-ScreenInfo -Message " $algorithm" -Type Error
                }
            }
            $CAKey.HashAlgorithm = $param.HashAlgorithmName
        }
        $CASetup.SetCASetupProperty(1,$CAKey)
        Write-Verbose -Message "Hash Algorithm = '$($CAKey.HashAlgorithm)'"



        if ($param.CAType)
        {
            $SupportedTypes = $CASetup.GetSupportedCATypes()

            $CATypesByName = @{'EnterpriseRootCA'=0;'EnterpriseSubordinateCA'=1;'StandaloneRootCA'=3;'StandaloneSubordinateCA'=4}
            $SelectedType = $CATypesByName[$param.CAType]

            if ($SupportedTypes -notcontains $SelectedType)
            {
                Write-Host "Selected CA type: '$CAType' is not supported by current Windows Server installation."
                Write-Host "The following CA types are supported by this installation:"
                #foreach ($caType in (
                {
                    #Write-ScreenInfo -Message "$([int[]]$CASetup.GetSupportedCATypes() | %{$CATypesByVal[$_]})
                }
            }
        }
        else
        {
            $CASetup.SetCASetupProperty($CAPRopertyByName.CAType,$SelectedType)
        }
        Write-Verbose -Message "CAType = '$($param.CAType)'"



        if ($SelectedType -eq 0 -or $SelectedType -eq 3 -and $param.ValidityPeriodUnits -ne 0)
        {
            try
            {
                $CASetup.SetCASetupProperty(6,([int]$param.ValidityPeriodUnits))
            }
            catch
            {
                Write-Host "The specified CA certificate validity period '$($param.ValidityPeriodUnits)' is invalid."
            }
        }
        Write-Verbose -Message "ValidityPeriod = '$($param.ValidityPeriodUnits)'"



        $DN = New-Object -ComObject X509Enrollment.CX500DistinguishedName
        # validate X500 name format
        try
        {
            $DN.Encode("CN=$($param.CACommonName)",0x0)
        }
        catch
        {
            Write-Host "Specified CA name or CA name suffix is not correct X.500 Distinguished Name."
        }
        $CASetup.SetCADistinguishedName("CN=$($param.CACommonName)", $true, $true, $true)
        Write-Verbose -Message "CADistinguishedName = 'CN=$($param.CACommonName)'"



        if ($CASetup.GetCASetupProperty(0) -eq 1 -and $param.ParentCA)
        {
            [void]($param.ParentCA -match "^(.+)\\(.+)$")
            try
            {
                $CASetup.SetParentCAInformation($param.ParentCA)
            }
            catch
            {
                Write-Host "The specified parent CA information '$param.ParentCA' is incorrect. Make sure if parent CA information is correct (you must specify existing CA) and is supplied in a 'CAComputerName\CASanitizedName' form."
            }
        }
        Write-Verbose -Message "PArentCA = 'CN=$($param.CACommonName)'"









        if ($param.DatabaseDirectory -eq '')
        {
            $param.DatabaseDirectory = 'C:\Windows\system32\CertLog'
        }
        Write-Verbose -Message "DatabaseDirectory = '$($param.DatabaseDirectory)'"

        if ($param.LogDirectory -eq '')
        {
            $param.LogDirectory = 'C:\Windows\system32\CertLog'
        }
        Write-Verbose -Message "LogDirectory = '$($param.LogDirectory)'"



        if ($param.DatabaseDirectory -ne "" -and $param.LogDirectory -ne "")
        {
            try
            {
                $CASetup.SetDatabaseInformation($param.DatabaseDirectory,$param.LogDirectory,$null,$OverwriteExisting)
            }
            catch
            {
                Write-Verbose -Message 'Specified path to either database directory or log directory is invalid.'
            }
        }


        try
        {
            Write-Verbose -Message 'Installing Certification Authority'
            $CASetup.Install()
            if ($CASetup.GetCASetupProperty(0) -eq 1)
            {
                $CASName = (Get-ItemProperty HKLM:\System\CurrentControlSet\Services\CertSvc\Configuration).Active
                $SetupStatus = (Get-ItemProperty HKLM:\System\CurrentControlSet\Services\CertSvc\Configuration\$CASName).SetupStatus
                $RequestID = (Get-ItemProperty HKLM:\System\CurrentControlSet\Services\CertSvc\Configuration\$CASName).RequestID
            }
            Write-Verbose -Message 'Certification Authority role is successfully installed'
        }
        catch
        {
            Write-Error $_ -ErrorAction Stop
        }








        if ($param.ForestAdminUserName)
        {
            Write-Verbose -Message "Removing $($param.ForestAdminUserName) to local administrators group"
            $localGroup = ([ADSI]'WinNT://./Administrators,group')
            $localGroup.psbase.Invoke('Remove', ([ADSI]"WinNT://$($param.ForestAdminUserName.replace('\', '/'))").path)
        }


        if ($param.InstallWebEnrollment)
        {
            Write-Verbose -Message 'InstallWebRole is True, hence setting InstallWebRole to True'
            $param.InstallWebRole = $true
        }

        if ($param.InstallWebRole)
        {
            Write-Verbose -Message 'Check if web role is already installed'
            if (!((Get-WindowsFeature -Name 'web-server').Installed))
            {
                Write-Verbose -Message 'Web role is NOT already installed. Installing it now.'
                Add-WindowsFeature -Name 'Web-Server' -IncludeManagementTools

                #Allow "+" characters in URL for supporting delta CRLs
                #Set-WebConfiguration -Filter system.webServer/security/requestFiltering -PSPath 'IIS:\sites\Default Web Site' -Value @{allowDoubleEscaping=$true}
            }
        }

        if ($param.InstallWebEnrollment)
        {
            Write-Verbose -Message 'Installing Web Enrollment service'
            Install-WebEnrollment "$($param.ComputerName)\$($param.CACommonName)"
        }


        #endregion - Install CA

        #region - Configure IIS virtual directories
        if ($param.UseHTTPAia)
        {
            #New-WebVirtualDirectory -Site 'Default Web Site' -Name Aia -PhysicalPath 'C:\Windows\System32\CertSrv\CertEnroll' | Out-Null
            #New-WebVirtualDirectory -Site 'Default Web Site' -Name Cdp -PhysicalPath 'C:\Windows\System32\CertSrv\CertEnroll' | Out-Null
        }
        #endregion - Configure IIS virtual directories








        #region - Configure CA
        function Invoke-CustomExpression
        {
            param ($Command)

            Invoke-Expression -Command $command
            Write-Verbose -Message $command
        }




        #Declare configuration NC
        if ($param.CAType -like 'Enterprise*')
        {
            $lDAPname = ''
            foreach ($part in ($param.DomainName.split('.')))
            {
                $lDAPname += ",DC=$part"
            }
            Invoke-CustomExpression -Command "certutil.exe -setreg ""CA\DSConfigDN"" ""CN=Configuration$lDAPname"""
        }

        #Apply the required CDP Extension URLs
        $command = "certutil.exe -setreg CA\CRLPublicationURLs ""1:$($Env:WinDir)\system32\CertSrv\CertEnroll\%3%8%9.crl"
        if ($param.UseLDAPCRL) { $command += '\n11:ldap:///CN=%7%8,CN=%2,CN=CDP,CN=Public Key Services,CN=Services,%6%10' }
        if ($param.UseHTTPCRL) { $command += "\n2:$($param.CDPHTTPURL01)/%3%8%9.crl" }
        if ($param.AIAHTTPURL01UploadLocation) { $command += "\n1:$($param.AIAHTTPURL01UploadLocation)/%3%8%9.crl" }
        $command += '"'
        Invoke-CustomExpression -Command $command

        #Apply the required AIA Extension URLs
        $command = "certutil.exe -setreg CA\CACertPublicationURLs ""1:$($Env:WinDir)\system3\CertSrv\CertEnroll\%1_%3%4.crt"
        if ($param.UseLDAPAia) { $command += '\n3:ldap:///CN=%7,CN=AIA,CN=Public Key Services,CN=Services,%6%11' }
        if ($param.UseHTTPAia) { $command += "\n2:$($param.AIAHTTPURL01)/%1_%3%4.crt" }
        if ($param.AIAHTTPURL01UploadLocation) { $command += "\n1:$($param.AIAHTTPURL01UploadLocation)/%3%8%9.crl" }
        $command += '"'
        Invoke-CustomExpression -Command $command

        #Define default maximum certificate lifetime for issued certificates
        Invoke-CustomExpression -Command "certutil.exe -setreg ca\ValidityPeriodUnits $($param.CertsValidityPeriodUnits)"
        Invoke-CustomExpression -Command "certutil.exe -setreg ca\ValidityPeriod ""$($param.CertsValidityPeriod)"""

        #Define CRL Publication Intervals
        Invoke-CustomExpression -Command "certutil.exe -setreg CA\CRLPeriodUnits $($param.CRLPeriodUnits)"
        Invoke-CustomExpression -Command "certutil.exe -setreg CA\CRLPeriod ""$($param.CRLPeriod)"""

        #Define CRL Overlap
        Invoke-CustomExpression -Command "certutil.exe -setreg CA\CRLOverlapUnits $($param.CRLOverlapUnits)"
        Invoke-CustomExpression -Command "certutil.exe -setreg CA\CRLOverlapPeriod ""$($param.CRLOverlapPeriod)"""

        #Define Delta CRL
        Invoke-CustomExpression -Command "certutil.exe -setreg CA\CRLDeltaUnits $($param.CRLDeltaPeriodUnits)"
        Invoke-CustomExpression -Command "certutil.exe -setreg CA\CRLDeltaPeriod ""$($param.CRLDeltaPeriod)"""

        #Enable Auditing Logging
        Invoke-CustomExpression -Command 'certutil.exe -setreg CA\Auditfilter 0x7F'

        #Enable UTF-8 Encoding
        Invoke-CustomExpression -Command 'certutil.exe -setreg ca\forceteletex +0x20'

        if ($param.CAType -like '*root*')
        {
            #Disable Discrete Signatures in Subordinate Certificates (WinXP KB968730)
            Invoke-CustomExpression -Command 'certutil.exe -setreg CA\csp\AlternateSignatureAlgorithm 0'

            #Force digital signature removal in KU for cert issuance (see also kb888180)
            Invoke-CustomExpression -Command 'certutil.exe -setreg policy\EditFlags -EDITF_ADDOLDKEYUSAGE'

            #Enable SAN
            Invoke-CustomExpression -Command 'certutil.exe -setreg policy\EditFlags +EDITF_ATTRIBUTESUBJECTALTNAME2'

            #Configure policy module to automatically issue certificates when requested
            Invoke-CustomExpression -Command 'certutil.exe -setreg ca\PolicyModules\CertificateAuthority_MicrosoftDefault.Policy\RequestDisposition 1'
        }
        #If CA is Root CA and Sub CAs are present, disable (do not publish) templates (except SubCA template)
        if ($param.DoNotLoadDefaultTemplates)
        {
            Invoke-CustomExpression -Command 'certutil.exe -SetCATemplates +SubCA'
        }
        #endregion - Configure CA





        #region - Restart of CA
        if ((Get-Service -Name 'CertSvc').Status -eq 'Running')
        {
            Write-Verbose -Message 'Stopping ADCS Service'
            $totalretries = 5
            $retries = 0
            do
            {
                Stop-Service -Name 'CertSvc' -ErrorAction SilentlyContinue
                if ((Get-Service -Name 'CertSvc').Status -ne 'Stopped')
                {
                    $retries++
                    Start-Sleep -Seconds 1
                }
            }
            until (((Get-Service -Name 'CertSvc').Status -eq 'Stopped') -or ($retries -ge $totalretries))

            if ((Get-Service -Name 'CertSvc').Status -eq 'Stopped')
            {
                Write-Verbose -Message 'ADCS service is now stopped'
            }
            else
            {
                Write-Error -Message 'Could not stop ADCS Service after several retries'
                return
            }
        }

        Write-Verbose -Message 'Starting ADCS Service now'
        $totalretries = 5
        $retries = 0
        do
        {
            Start-Service -Name 'CertSvc' -ErrorAction SilentlyContinue
            if ((Get-Service -Name 'CertSvc').Status -ne 'Running')
            {
                $retries++
                Start-Sleep -Seconds 1
            }
        }
        until (((Get-Service -Name 'CertSvc').Status -eq 'Running') -or ($retries -ge $totalretries))

        if ((Get-Service -Name 'CertSvc').Status -eq 'Running')
        {
            Write-Verbose -Message 'ADCS service is now started'
        }
        else
        {
            Write-Error -Message 'Could not start ADCS Service after several retries'
            return
        }
        #endregion - Restart of CA


        Write-Verbose -Message 'Waiting for admin interface to be ready'
        $totalretries = 10
        $retries = 0
        do
        {
            $result = Invoke-Expression -Command "certutil.exe -pingadmin .\$($param.CACommonName)"
            if (!($result | Where-Object { $_ -like '*interface is alive*' }))
            {
                $retries++
                Write-Verbose -Message "Admin interface not ready. Check $retries of $totalretries"
                if ($retries -lt $totalretries) { Start-Sleep -Seconds 10 }
            }
        }
        until (($result | Where-Object { $_ -like '*interface is alive*' }) -or ($retries -ge $totalretries))

        if ($result | Where-Object { $_ -like '*interface is alive*' })
        {
            Write-Verbose -Message 'Admin interface is now ready'
        }
        else
        {
            Write-Error -Message 'Admin interface was not ready after several retries'
            return
        }


        #region - Issue of CRL
        Start-Sleep -Seconds 2
        Invoke-Expression -Command 'certutil.exe -crl' | Out-Null
        $totalretries = 12
        $retries = 0
        do
        {
            Start-Sleep -Seconds 5
            $retries++
        }
        until ((Get-ChildItem "$env:systemroot\system32\CertSrv\CertEnroll\*.crl") -or ($retries -ge $totalretries))

        #endregion - Issue of CRL

        if (($param.CAType -like 'Enterprise*') -and ($param.DoNotLoadDefaultTemplates)) { Invoke-Expression 'certutil.exe -SetCATemplates +SubCA' }
    }

    #endregion

    Write-PSFMessage -Message "Performing installation of $($param.CAType) on '$($param.ComputerName)'"
    $cred = (New-Object System.Management.Automation.PSCredential($param.UserName, ($param.Password | ConvertTo-SecureString -AsPlainText -Force)))
    $caSession = New-LabPSSession -ComputerName $param.ComputerName
    $Job = Invoke-Command -Session $caSession -Scriptblock $caScriptBlock -ArgumentList $param -AsJob -JobName "Install CA on '$($param.Computername)'" -Verbose

    $Job

    Write-LogFunctionExit
}
