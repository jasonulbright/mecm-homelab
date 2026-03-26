
# Get public and private function definition files.
$modulebase =  $PSScriptRoot

$script:hostFilePath = if ($PSEdition -eq 'Desktop' -or $IsWindows)
{
    "$($env:SystemRoot)\System32\drivers\etc\hosts"
}
elseif ($PSEdition -eq 'Core' -and $IsLinux)
{
    '/etc/hosts'
}

# Types first
$typeExists = try { [AutomatedLab.Common.Win32Exception] }catch { }
if (-not $typeExists)
{
    try
    {
        if ($PSEdition -eq 'Core')
        {
            Add-Type -Path $modulebase/lib/core/AutomatedLab.Common.dll -ErrorAction Stop
        }
        else
        {
            Add-Type -Path $modulebase/lib/full/AutomatedLab.Common.dll -ErrorAction Stop
        }
    }
    catch
    {
        Write-Warning -Message "Unable to add AutomatedLab.Common.dll - GPO and PKI functionality might be impaired.`r`nException was: $($_.Exception.Message), $($_.Exception.LoaderExceptions)"
    }
}

try
{
    [ServerCertificateValidationCallback]::Ignore()
}
catch { }

function Add-AccountPrivilege
{
    param
    (
        [Parameter(Mandatory = $true)]
        [string[]]
        $UserName,

        [validateSet('SeNetworkLogonRight', 
            'SeRemoteInteractiveLogonRight', 
            'SeBatchLogonRight', 
            'SeInteractiveLogonRight', 
            'SeServiceLogonRight', 
            'SeDenyNetworkLogonRight', 
            'SeDenyInteractiveLogonRight', 
            'SeDenyBatchLogonRight', 
            'SeDenyServiceLogonRight', 
            'SeDenyRemoteInteractiveLogonRight', 
            'SeTcbPrivilege', 
            'SeMachineAccountPrivilege', 
            'SeIncreaseQuotaPrivilege', 
            'SeBackupPrivilege', 
            'SeChangeNotifyPrivilege', 
            'SeSystemTimePrivilege', 
            'SeCreateTokenPrivilege', 
            'SeCreatePagefilePrivilege', 
            'SeCreateGlobalPrivilege', 
            'SeDebugPrivilege', 
            'SeEnableDelegationPrivilege', 
            'SeRemoteShutdownPrivilege', 
            'SeAuditPrivilege', 
            'SeImpersonatePrivilege', 
            'SeIncreaseBasePriorityPrivilege', 
            'SeLoadDriverPrivilege', 
            'SeLockMemoryPrivilege', 
            'SeSecurityPrivilege', 
            'SeSystemEnvironmentPrivilege', 
            'SeManageVolumePrivilege', 
            'SeProfileSingleProcessPrivilege', 
            'SeSystemProfilePrivilege', 
            'SeUndockPrivilege', 
            'SeAssignPrimaryTokenPrivilege', 
            'SeRestorePrivilege', 
            'SeShutdownPrivilege', 
            'SeSynchAgentPrivilege', 
            'SeTakeOwnershipPrivilege' 
        )]
        [string[]]
        $Privilege
    )    
    
    $lsaWrapper = New-Object -TypeName MyLsaWrapper.LsaWrapper -ErrorAction Stop

    foreach ($User in $UserName)
    {
        foreach ($Priv in $Privilege)
        {
            $lsaWrapper.AddPrivileges($User, $Priv)
            Start-Sleep -Milliseconds 250
            $lsaWrapper.AddPrivileges($User, $Priv)
        }
    }    
}

function Add-FunctionToPSSession
{
    [CmdletBinding(
        SupportsShouldProcess = $false,
        ConfirmImpact = 'None'
    )]

    param
    ( 
        [Parameter(
            HelpMessage	= 'Provide the session(s) to load the functions into', 
            Mandatory	= $true,
            Position	= 0
        )]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.Runspaces.PSSession[]] 
        $Session,

        [Parameter( 
            HelpMessage = 'Provide the function info to load into the session(s)', 
            Mandatory = $true, 
            Position = 1, 
            ValueFromPipeline	= $true 
        )]
        [ValidateNotNull()]
        [System.Management.Automation.FunctionInfo]
        $FunctionInfo
    )

    begin 
    {
        $cmdName = (Get-PSCallStack)[0].Command
        Write-Debug "[$cmdName] Entering function"

        $scriptBlock = 
        {
            param([string]$Path, [string]$Definition)
            $null = Set-Item -Path Function:\$Path -Value $Definition
        }
    }

    process
    {
        Invoke-Command -Session $Session -ScriptBlock $scriptBlock -ArgumentList $FunctionInfo.Name, $FunctionInfo.Definition
    }

    end
    {
        Write-Debug "[$cmdName] Exiting function"
    }
}

function Add-StringIncrement
{
    param(
        [Parameter(Mandatory = $true)]
        [string]$String
    )
    
    $testNumberPattern = '^(?<text>.*?) (?<number>\d+)$'
    
    $result = $String -match $testNumberPattern
    
    if ($Matches.Number)
    {
        $String = $String.Substring(0, $String.Length - $Matches.Number.Length) + ([int]$Matches.Number + 1)
    }
    else
    {
        $String = $String + ' 0'
    }
    
    $String
}

function Add-VariableToPSSession
{
    [CmdletBinding(
        SupportsShouldProcess = $false,
        ConfirmImpact = 'None'
    )]

    param
    ( 
        [Parameter(
            HelpMessage	= 'Provide the session(s) to load the functions into', 
            Mandatory	= $true,
            Position	= 0
        )]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.Runspaces.PSSession[]] 
        $Session,

        [Parameter( 
            HelpMessage = 'Provide the variable info to load into the session(s)', 
            Mandatory = $true, 
            Position = 1, 
            ValueFromPipeline	= $true 
        )]
        [ValidateNotNull()]
        [System.Management.Automation.PSVariable]
        $PSVariable
    )

    begin 
    {
        $cmdName = (Get-PSCallStack)[0].Command
        Write-Debug "[$cmdName] Entering function"

        $scriptBlock = 
        {
            param([string]$_AL_Path, [object]$Value)
            $null = Set-Item -Path Variable:\$_AL_Path -Value $Value
        }
    }

    process
    {
        if ($PSVariable.Name -eq 'PSBoundParameters')
        {
            Invoke-Command -Session $Session -ScriptBlock $scriptBlock -ArgumentList 'ALBoundParameters', $PSVariable.Value
        }
        else
        {
            Invoke-Command -Session $Session -ScriptBlock $scriptBlock -ArgumentList $PSVariable.Name, $PSVariable.Value
        }
    }

    end
    {
        Write-Debug "[$cmdName] Exiting function"
    }
}

function Get-ConsoleText
{
    [CmdletBinding()]
    param()
    
    # Check the host name and exit if the host is not the Windows PowerShell console host. 
    if ($host.Name -eq 'Windows PowerShell ISE Host')
    { 
        $psISE.CurrentPowerShellTab.ConsolePane.Text
    }
    elseif ($host.Name -eq 'ConsoleHost')
    {
        $textBuilderConsole = New-Object System.Text.StringBuilder
        $textBuilderLine = New-Object System.Text.StringBuilder

        # Grab the console screen buffer contents using the Host console API.
        $bufferWidth = $host.UI.RawUI.BufferSize.Width
        $bufferHeight = $host.UI.RawUI.CursorPosition.Y 
        $rec = New-Object System.Management.Automation.Host.Rectangle(0, 0, ($bufferWidth), $bufferHeight)
        $buffer = $host.UI.RawUI.GetBufferContents($rec) 

        # Iterate through the lines in the console buffer. 
        for ($i = 0; $i -lt $bufferHeight; $i++) 
        { 
            for ($j = 0; $j -lt $bufferWidth; $j++) 
            { 
                $cell = $buffer[$i, $j] 
                $null = $textBuilderLine.Append($cell.Character)
            }
            $null = $textBuilderConsole.AppendLine($textBuilderLine.ToString().TrimEnd())
            $textBuilderLine = New-Object System.Text.StringBuilder
        }

        $textBuilderConsole.ToString()
        Write-Verbose "$bufferHeight lines have been copied to the clipboard"
    }
}

<#
        Script Name	: Get-NetFrameworkVersion.ps1
        Description	: This script reports the various .NET Framework versions installed on the local or a remote computer.
        Author		: Martin Schvartzman
        Reference   : https://msdn.microsoft.com/en-us/library/hh925568
#>
function Get-DotNetFrameworkVersion
{
    [CmdletBinding()]
    param
    (
        [string]$ComputerName = $env:COMPUTERNAME
    )

    $dotNetRegistry = 'SOFTWARE\Microsoft\NET Framework Setup\NDP'
    $dotNet4Registry = 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full'
    $dotNet4Builds = @{
        '30319'  = @{ Version = [System.Version]'4.0' }
        '378389' = @{ Version = [System.Version]'4.5' }
        '378675' = @{ Version = [System.Version]'4.5.1'   ; Comment = '(8.1/2012R2)' }
        '378758' = @{ Version = [System.Version]'4.5.1'   ; Comment = '(8/7 SP1/Vista SP2)' }
        '379893' = @{ Version = [System.Version]'4.5.2' }
        '380042' = @{ Version = [System.Version]'4.5'     ; Comment = 'and later with KB3168275 rollup' }
        '393295' = @{ Version = [System.Version]'4.6'     ; Comment = '(Windows 10)' }
        '393297' = @{ Version = [System.Version]'4.6'     ; Comment = '(NON Windows 10)' }
        '394254' = @{ Version = [System.Version]'4.6.1'   ; Comment = '(Windows 10)' }
        '394271' = @{ Version = [System.Version]'4.6.1'   ; Comment = '(NON Windows 10)' }
        '394802' = @{ Version = [System.Version]'4.6.2'   ; Comment = '(Windows 10 1607)' }
        '394806' = @{ Version = [System.Version]'4.6.2'   ; Comment = '(NON Windows 10)' }
        '460798' = @{ Version = [System.Version]'4.7'     ; Comment = '(Windows 10 1703)' }
        '460805' = @{ Version = [System.Version]'4.7'     ; Comment = '(NON Windows 10)' }
        '461308' = @{ Version = [System.Version]'4.7.1'   ; Comment = '(Windows 10 1709)' }
        '461310' = @{ Version = [System.Version]'4.7.1'   ; Comment = '(NON Windows 10)' }
        '461808' = @{ Version = [System.Version]'4.7.2'   ; Comment = '(Windows 10 1803)' }
        '461814' = @{ Version = [System.Version]'4.7.2'   ; Comment = '(NON Windows 10)' }
        '528040' = @{ Version = [System.Version]'4.8'     ; Comment = '(Windows 10 1903)' }
        '528049' = @{ Version = [System.Version]'4.8'     ; Comment = '(NON Windows 10)' }
        '528372' = @{ Version = [System.Version]'4.8'     ; Comment = '(Windows 10 2004)' }
        '528449' = @{ Version = [System.Version]'4.8'     ; Comment = '(Windows 11 / Server 2022)' }
        '533320' = @{ Version = [System.Version]'4.8.1'   ; Comment = '(Windows 11 / Server 2022)' }
        '533325' = @{ Version = [System.Version]'4.8.1'   ; Comment = '(NON Windows 11)' }	
    }

    foreach ($computer in $ComputerName)
    {
        if ($regKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $computer))
        {
            if ($netRegKey = $regKey.OpenSubKey("$dotNetRegistry"))
            {
                foreach ($versionKeyName in $netRegKey.GetSubKeyNames())
                {
                    if ($versionKeyName -match '^v[123]')
                    {
                        $versionKey = $netRegKey.OpenSubKey($versionKeyName)
                        $version = [System.Version]($versionKey.GetValue('Version', ''))
                        New-Object -TypeName PSObject -Property ([ordered]@{
                                ComputerName = $computer
                                Build        = $version.Build
                                Version      = $version
                                Comment      = ''
                            })
                    }
                }
            }

            if ($net4RegKey = $regKey.OpenSubKey("$dotNet4Registry"))
            {
                if (-not ($net4Release = $net4RegKey.GetValue('Release')))
                {
                    $net4Release = 30319
                }
                New-Object -TypeName PSObject -Property ([ordered]@{
                        ComputerName = $Computer
                        Build        = $net4Release
                        Version      = $dotNet4Builds["$net4Release"].Version
                        Comment      = $dotNet4Builds["$net4Release"].Comment
                    })
            }
        }
    }
}

function Get-FullMesh
{
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [array]$List,

        [switch]$OneWay
    )

    $mesh = New-Object System.Collections.ArrayList

    foreach ($item1 in $List)
    {
        foreach ($item2 in $list)
        {
            if ($item1 -eq $item2)
            { continue }

            if ($mesh.Contains(($item1, $item2)))
            { continue }

            if ($OneWay)
            {
                if ($mesh.Contains(($item2, $item1)))
                { continue }
            }

            $mesh.Add((New-Object (Get-Type -GenericType Mesh.Item -T string) -Property @{ Source = $item1; Destination = $item2 } )) | Out-Null
        }
    }

    $mesh
}

function Get-ModuleDependency
{
	[CmdletBinding()]
	param
	(
        [Parameter(Mandatory = $true)]
		[System.Management.Automation.PSModuleInfo]
		$Module,

        [switch]
        $AsModuleInfo
	)

	if ($Module.RequiredModules)
	{
		Write-Verbose "$($Module.Name) has required modules"
		foreach ($moduleName in $Module.RequiredModules)
		{
			$moduleInfo = Get-Module -ListAvailable -Name $moduleName.Name
			if ($moduleName.Version) {$moduleInfo = $moduleInfo | Where-Object Version -eq $moduleName.Version}
			$moduleInfo = $moduleInfo | Sort-Object Version -Descending | Select-Object -First 1
			Write-Verbose "Detecting dependencies for $($moduleInfo.Name)"
			Get-ModuleDependency -Module $moduleInfo -AsModuleInfo:$AsModuleInfo.IsPresent
		}
	}
	
	if ($AsModuleInfo.IsPresent)
    {
        return $Module
    }
    
    $Module.ModuleBase
}

function Get-RunspacePool
{
    [OutputType([System.Management.Automation.Runspaces.RunspacePool[]])]
    [CmdletBinding()]
    param
    (
        [Parameter()]
        [int]
        $ThrottleLimit,

        [Parameter()]
        [System.Threading.ApartmentState]
        $ApartmentState
    )

    $pools = $(Get-Variable -Name ALCommonRunspacePool_* -Scope Script -ErrorAction SilentlyContinue).Value

    if ($ThrottleLimit)
    {
        $pools = $pools.Where({$_.GetMaxRunspaces() -eq $ThrottleLimit})
    }

    if ($ApartmentState)
    {
        $pools = $pools.Where({$_.ApartmentState -eq $ApartmentState})
    }

    $pools
}

function Get-StringSection
{
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [string]$String,

        [Parameter(Mandatory = $true)]
        [int]$SectionSize
    )

    process
    {
        0..($String.Length - 1) | 
            Group-Object -Property { [System.Math]::Truncate($_ / $SectionSize) } |
            ForEach-Object { -join $String[$_.Group] }
    }
}

function Get-Type
{
    param (
        [Parameter(Position = 0, Mandatory = $true)]
        [string] $GenericType,
        
        [Parameter(Position = 1, Mandatory = $true)]
        [string[]] $T
    )
    
    $T = $T -as [type[]]
    
    try
    {
        $generic = [type]($GenericType + '`' + $T.Count)
        $generic.MakeGenericType($T)
    }
    catch [Exception]
    {
        throw New-Object -TypeName System.Exception -ArgumentList ('Cannot create generic type', $_.Exception)
    }
}

function Install-SoftwarePackage
{
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Path,
        
        [string]$CommandLine,
        
        [bool]$AsScheduledJob,
        
        [bool]$UseShellExecute,

        [string]$WorkingDirectory,

        [int[]]$ExpectedReturnCodes,

        [system.management.automation.pscredential]$Credential
    )    
    
    #region New-InstallProcess
    function New-InstallProcess
    {
        param(
            [Parameter(Mandatory = $true)]
            [string]$Path,

            [string]$CommandLine,
            
            [bool]$UseShellExecute,

            [string]$WorkingDirectory
        )

        $pInfo = New-Object -TypeName System.Diagnostics.ProcessStartInfo
        $pInfo.FileName = $Path
        if (-not [string]::IsNullOrWhiteSpace($WorkingDirectory)) { $pInfo.WorkingDirectory = $WorkingDirectory }

        $pInfo.UseShellExecute = $UseShellExecute
        if (-not $UseShellExecute)
        {
            $pInfo.RedirectStandardError = $true
            $pInfo.RedirectStandardOutput = $true
        }
        $pInfo.Arguments = $CommandLine

        $p = New-Object -TypeName System.Diagnostics.Process
        $p.StartInfo = $pInfo
        Write-Verbose -Message "Starting process: $($pInfo.FileName) $($pInfo.Arguments)"
        $p.Start() | Out-Null
        Write-Verbose "The installation process ID is $($p.Id)"
        $p.WaitForExit()
        Write-Verbose -Message 'Process exited. Reading output'

        $params = @{
            Process = $p
            LabSourcesConnectOutput = $labSourcesConnectOutput
        }
        if (-not $UseShellExecute)
        {
            $params.Output = $p.StandardOutput.ReadToEnd()
            $params.Error = $p.StandardError.ReadToEnd()
        }
        New-Object -TypeName PSObject -Property $params
    }
    #endregion New-InstallProcess

    #if the path cannot be found and starts with \\automatedlabsources...
    if ((-not (Test-Path -Path $Path) -and $Path -match '\\automatedlabsources[a-z]{5}\.file\.core\.windows\.net'))
    {
        #we assume, the LabSources share was not mapped correctly and try again by calling 'C:\AL\AzureLabSources.ps1'
        $labSourcesConnectOutput = C:\AL\AzureLabSources.ps1 2> $null
        if ($labSourcesConnectOutput.AlternativeLabSourcesPath)
        {
            $Path = $Path.Replace($labSourcesConnectOutput.LabSourcesPath, $labSourcesConnectOutput.AlternativeLabSourcesPath)
        }
    }

    if (-not (Test-Path -Path $Path -PathType Leaf))
    {
        Write-Error "The file '$Path' could not found"
        return        
    }
        
    $start = Get-Date
    Write-Verbose -Message "Starting setup of '$Path' with the following command"
    Write-Verbose -Message "`t$CommandLine"
    Write-Verbose -Message "The timeout is $Timeout minutes, starting at '$start'"
    
    $installationMethod = [System.IO.Path]::GetExtension($Path)
    $installationFile = [System.IO.Path]::GetFileName($Path)
    
    if ($installationMethod -eq '.msi')
    {        
        [string]$CommandLine = if (-not $CommandLine)
        {
            @(
                "/I `"$Path`"", # Install this MSI
                '/QN', # Quietly, without a UI
                "/L*V `"$([System.IO.Path]::GetTempPath())$([System.IO.Path]::GetFileNameWithoutExtension($Path)).log`""     # Verbose output to this log
            )
        }
        else
        {
            '/I "{0}" {1}' -f $Path, $CommandLine # Install this MSI
        }
        
        Write-Verbose -Message 'Installation arguments for MSI are:'
        Write-Verbose -Message "`tPath: $Path"
        Write-Verbose -Message "`tLog File: '`t$([System.IO.Path]::GetTempPath())$([System.IO.Path]::GetFileNameWithoutExtension($Path)).log'"
        
        $Path = 'msiexec.exe'
    }
    elseif ($installationMethod -eq '.msp')
    {
        [string]$CommandLine = if (-not $CommandLine)
        {
            @(
                "/P `"$Path`"", # Install this MSI
                '/QN', # Quietly, without a UI
                "/L*V `"$([System.IO.Path]::GetTempPath())$([System.IO.Path]::GetFileNameWithoutExtension($Path)).log`""     # Verbose output to this log
            )
        }
        else
        {
            '/P {0} {1}' -f $Path, $CommandLine # Install this MSI
        }
        
        Write-Verbose -Message 'Installation arguments for MSI are:'
        Write-Verbose -Message "`tPath: $Path"
        Write-Verbose -Message "`tLog File: '`t$([System.IO.Path]::GetTempPath())$([System.IO.Path]::GetFileNameWithoutExtension($Path)).log'"
        
        $Path = 'msiexec.exe'
    }
    elseif ($installationMethod -eq '.msu')
    {        
        $tempRemoteFolder = [System.IO.Path]::GetTempFileName()
        Remove-Item -Path $tempRemoteFolder
        New-Item -ItemType Directory -Path $tempRemoteFolder
        expand.exe -F:* $Path $tempRemoteFolder
        $Path = 'dism.exe'
        $CommandLine = "/Online /Add-Package /PackagePath:""$tempRemoteFolder"" /NoRestart /Quiet"
    }
    elseif ($installationMethod -eq '.exe')
    { }
    else
    {
        Write-Error -Message 'The extension of the file to install is unknown'
        return
    }

    Write-Verbose -Message "Starting installation of $installationMethod file"

    if ($AsScheduledJob)
    {
        $jobName = "AL_$([guid]::NewGuid())"
        Write-Verbose "In the AsScheduledJob mode, creating scheduled job named '$jobName'"
            
        if ($PSVersionTable.PSVersion -lt '3.0')
        {
            Write-Verbose "Running SCHTASKS.EXE as PowerShell Version is <2.0"
            $processName = [System.IO.Path]::GetFileNameWithoutExtension($Path)
            $d = "{0:HH:mm}" -f (Get-Date).AddMinutes(1)

            "$Path $CommandLine" | Out-File -FilePath "C:\$jobName.cmd" -Encoding default

            if ($Credential)
            {
                SCHTASKS /Create /SC ONCE /ST $d /TN $jobName /RU "$($Credential.UserName)" /RP "$($Credential.GetNetworkCredential().Password)" /TR "C:\$jobName.cmd" | Out-Null
            }
            else
            {
                SCHTASKS /Create /SC ONCE /ST $d /TN $jobName /TR "C:\$jobName.cmd" /RU "SYSTEM" | Out-Null
            }

            Start-Sleep -Seconds 5 #allow some time to let the scheduled task run
            while (-not ($p))
            {
                Start-Sleep -Milliseconds 500
                $p = Get-Process -Name $processName -ErrorAction SilentlyContinue
            }

            $p.WaitForExit()
            Write-Verbose -Message 'Process exited. Reading output'

            $params = @{ Process = $p }
            $params.Add('Output', "Output cannot be retrieved using AsScheduledJob on PowerShell 2.0")
            $params.Add('Error', "Errors cannot be retrieved using AsScheduledJob on PowerShell 2.0")
            New-Object -TypeName PSObject -Property $params
        }
        else
        {
            Write-Verbose "Running Register-ScheduledJob as PowerShell Version is >=3.0"

            $scheduledJobParams = @{
                Name         = $jobName
                ScriptBlock  = (Get-Command -Name New-InstallProcess).ScriptBlock
                ArgumentList = $Path, $CommandLine, $UseShellExecute
                RunNow       = $true
            }
            if ($WorkingDirectory) { $scheduledJobParams.ArgumentList += $WorkingDirectory }
            if ($Credential) { $scheduledJobParams.Add('Credential', $Credential) }
            $scheduledJob = Register-ScheduledJob @scheduledJobParams
            Write-Verbose "ScheduledJob object registered with the ID $($scheduledJob.Id)"
            Start-Sleep -Seconds 5 #allow some time to let the scheduled task run
            
            while (-not $job)
            {
                Start-Sleep -Milliseconds 500
                $job = Get-Job -Name $jobName -ErrorAction SilentlyContinue
            }        
            $job | Wait-Job | Out-Null
            $result = $job | Receive-Job
        }
    }
    else
    {
        $result = New-InstallProcess -Path $Path -CommandLine $CommandLine -UseShellExecute $UseShellExecute -WorkingDirectory $WorkingDirectory
    }
    
    Start-Sleep -Seconds 5
    
    if ($AsScheduledJob)
    {
        if ($PSVersionTable.PSVersion -lt '3.0')
        {
            schtasks.exe /DELETE /TN $jobName /F | Out-Null
            Remove-Item -Path "C:\$jobName.cmd"
        }
        else
        {
            Write-Verbose "Unregistering scheduled job with ID $($scheduledJob.Id)"
            $scheduledJob | Unregister-ScheduledJob
        }
    }

    if ($installationMethod -eq '.msu')
    {
        Remove-Item -Path $tempRemoteFolder -Recurse -Confirm:$false
    }
        
    Write-Verbose "Exit code of installation process is '$($result.Process.ExitCode)'"
    if ($null -ne $result.Process.ExitCode -and (0, 3010 + $ExpectedReturnCodes) -notcontains $result.Process.ExitCode)
    {
        $onLegacyOs = try
        {
            $type = [AutomatedLab.Common.Win32Exception]
            $false
        }
        catch
        { $true }

        if ($onLegacyOs)
        {
            throw (New-Object System.ComponentModel.Win32Exception($result.Process.ExitCode))
        }

        throw (New-Object AutomatedLab.Common.Win32Exception($result.Process.ExitCode))
    }
    else
    {
        Write-Verbose -Message "Installation of '$installationFile' finished successfully"
        $result.Output
    }
}

function Invoke-Ternary 
{
    param
    (
        [scriptblock]
        $decider,

        [scriptblock]
        $ifTrue,

        [scriptblock]
        $ifFalse
    )

    if (&$decider)
    {
        &$ifTrue
    }
    else
    {
        &$ifFalse
    }
}
Set-Alias -Name ?? -Value Invoke-Ternary -Option AllScope -Description "Ternary Operator like '?' in C#" -Scope Global

function New-RunspacePool
{
    [OutputType([System.Management.Automation.Runspaces.RunspacePool])]
    [CmdletBinding()]
    param
    (
        [Parameter()]
        [int]
        $ThrottleLimit = 10,

        [Parameter()]
        [System.Threading.ApartmentState]
        $ApartmentState = 'Unknown',

        [Parameter()]
        [System.Management.Automation.PSVariable[]]
        $Variable,

        [Parameter()]
        [System.Management.Automation.FunctionInfo[]]
        $Function
    )

    $pool = Get-Variable -Name "ALCommonRunspacePool_$($ThrottleLimit)_$($ApartmentState)" -Scope Script -ErrorAction SilentlyContinue
    $InitialSessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()

    foreach ($func in $Function)
    {
        $ssFunc = [System.Management.Automation.Runspaces.SessionStateFunctionEntry]::new($func.Name, $func.ScriptBlock)
        $InitialSessionState.Commands.Add($ssFunc)
    }

    foreach ($var in $Variable)
    {
        $sessionVariable = [System.Management.Automation.Runspaces.SessionStateVariableEntry]::new($var.Name, $var.Value, $null)
        $InitialSessionState.Variables.Add($sessionVariable)
    }

    if (-not ($pool))
    {
        Write-Verbose -Message "Creating new runspace pool. Maximum Runspaces: $ThrottleLimit, ApartmentState: $ApartmentState, Variables: $($Variable.Count)"
        $pool = New-Variable -Name "ALCommonRunspacePool_$($ThrottleLimit)_$($ApartmentState)" -Scope Script -Value $([runspacefactory]::CreateRunspacePool($InitialSessionState)) -PassThru
        [void] $($pool.Value.SetMaxRunspaces($ThrottleLimit))

        if ($PSEdition -eq 'Desktop')
        {
            $pool.Value.ApartmentState = $ApartmentState
        }
    }
        
    $pool.Value
}

function Read-Choice
{ 
    param(
        [Parameter(Mandatory = $true)]
        [String[]]$ChoiceList, 

        [Parameter(Mandatory = $true)]
        [String]$Caption,
        
        [String]$Message,

        [int]$Default = 0
    )
    
    if (-not $Message) { $Message = $Caption }

    $choices = New-Object System.Collections.ObjectModel.Collection[System.Management.Automation.Host.ChoiceDescription]

    $choiceList | ForEach-Object { $choices.Add((New-Object "System.Management.Automation.Host.ChoiceDescription" -ArgumentList $_)) }

    $Host.UI.PromptForChoice($Caption, $Message, $choices, $Default) 
}

function Read-HashTable
{ 
    param(
        [Parameter(Mandatory = $true)]
        [String[]]$ChoiceList, 

        [Parameter(Mandatory = $true)]
        [String]$Caption,
        
        [String]$Message,

        [int]$Default = 0
    )
    
    #if (-not $Message) { $Message = $Caption }

    $fields = New-Object System.Collections.ObjectModel.Collection[System.Management.Automation.Host.FieldDescription]

    $choiceList | ForEach-Object { $fields.Add((New-Object System.Management.Automation.Host.FieldDescription -ArgumentList $_)) }

    $Host.UI.Prompt($Caption, $Message, $fields)    
}

function Receive-RunspaceJob
{
    [CmdletBinding()]
    param
    (
        [Parameter(ValueFromPipeline = $true)]
        [object[]]
        $RunspaceJob
    )

    process
    {
        while ($RunspaceJob.Handle.IsCompleted -contains $false)
        {
            Start-Sleep -Milliseconds 100
        }

        foreach ($job in $RunspaceJob)
        {
            $job.Shell.EndInvoke($job.handle)    
            $job.Shell.Dispose()    
        }
    }
}

function Remove-RunspacePool
{
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Low')]
    param
    (
        [Parameter(ValueFromPipeline = $true)]
        [System.Management.Automation.Runspaces.RunspacePool[]]
        $RunspacePool
    )

    process
    {
        foreach ($pool in $RunspacePool)
        {
            if ($PSCmdlet.ShouldProcess($pool.InstanceId, 'Closing runspace pool'))
            {
                $max = $pool.GetMaxRunspaces()
                $state = if ($null -ne $pool.ApartmentState) { $pool.ApartmentState } else {'Unknown'}

                $pool.Close()
                $pool.Dispose()

                Write-Verbose -Message "Attempting to remove ALCommonRunspacePool_$($max)_$($state)"
                Remove-Variable -Name "ALCommonRunspacePool_$($max)_$($state)" -Scope Script -ErrorAction SilentlyContinue
            }
        }
    }
}

function Send-ModuleToPSSession
{
    [CmdletBinding(  
        RemotingCapability = 'PowerShell', #V3 and above, values documented here: http://msdn.microsoft.com/en-us/library/system.management.automation.remotingcapability(v=vs.85).aspx
        SupportsShouldProcess = $false,
        ConfirmImpact = 'None',
        DefaultParameterSetName = ''
    )]
    
    [OutputType([System.IO.FileInfo])] #OutputType is supported in 3.0 and above
     
    param
    (
        [Parameter(
            HelpMessage = 'Provide the source module info object',
            Position = 0,
            Mandatory = $true, 
            ValueFromPipeline = $true
        )]
        [ValidateNotNullOrEmpty()]
        [PSModuleInfo]
        $Module,

        [Parameter(
            HelpMessage = 'Enter the destination path on the remote computer',
            Position = 1,
            Mandatory = $true, 
            ValueFromPipelineByPropertyName = $true
        )]
        [System.Management.Automation.Runspaces.PSSession[]] 
        $Session,
        
        [ValidateSet('AllUsers', 'CurrentUser')]
        [string]
        $Scope = 'AllUsers',

        [switch]
        $IncludeDependencies,

        [switch]
        $Move,

        [switch]
        $Encrypt,

        [switch]
        $NoWriteBuffer,

        [switch]
        $Verify,

        [switch]
        $Force,

        [switch]
        $NoClobber,

        [ValidateRange(1KB, 7.4MB)] #might be good to have much higher top end as the underlying max is controlled by New-PSSessionOption
        [uint32]
        $MaxBufferSize = 1MB
    )

    begin
    {
        $isCalledRecursivly = (Get-PSCallStack | Where-Object Command -eq $MyInvocation.InvocationName | Measure-Object | Select-Object -ExpandProperty Count) -gt 1
    }
    
    process
    {
        $fileParams = ([hashtable]$PSBoundParameters).Clone()
        [void]$fileParams.Remove('Module')
        [void]$fileParams.Remove('Scope')
        [void]$fileParams.Remove('IncludeDependencies')
        
        if ($Local:Module.ModuleType -eq 'Script' -and ($Local:Module.Path -notmatch '\.psd1$'))
        {
            Write-Error "Cannot send the module '$($Module.Name)' that is not described by a .psd1 file"
            return
        }

        #Remove any sessions where the same or newer module version already exists
        if (-not $Force.IsPresent)
        {
            Write-Verbose 'Filtering out target sessions that do not need the module'
            $Session = foreach ($item in $PSBoundParameters.Session)
            {
                #recursive calls will need to refresh the cached module list because we may have just placed new modules there
                if ($isCalledRecursivly)
                {
                    $modules = Get-Module -PSSession $item -ListAvailable -Name $Local:Module.Name -Refresh
                }
                else
                {
                    $modules = Get-Module -PSSession $item -ListAvailable -Name $Local:Module.Name
                }
                    
                #no version of the module installed, select for sending
                if (-not $modules)
                {
                    $item
                }
                else
                {
                    #determine what versions we have
                    $versions = $modules | ForEach-Object { [System.Version]$_.Version } | Sort-Object -Unique -Descending
                    $highestVersion = $versions | Select-Object -First 1

                    #if the version we are sending is newer than the highest installed version, select for sending
                    if ([System.Version]$Local:Module.Version -gt $highestVersion)
                    {
                        $item
                    }
                    elseif ($highestVersion -gt [System.Version]$Local:Module.Version)
                    {
                        write-Warning "Skipping $($item.ComputerName) which has a higher version $highestVersion of the module installed"
                    }
                    else
                    {
                        write-Verbose  "Skipping $($item.ComputerName) because the same version of the module is installed already"
                    }
                }
            }
        }

        foreach ($s in $Session)
        {
            [version]$sessionVersion = Invoke-Command -Session $s -ScriptBlock {
                if ($PSEdition -eq 'core') {return ('{0}.{1}.{2}' -f $PSVersionTable.PSVersion.Major,$PSVersionTable.PSVersion.Minor,$PSVersionTable.PSVersion.Patch)}
                $PSVersionTable.PSVersion   
            }

            if ($Local:Module.PowerShellVersion -gt $sessionVersion)
            {
                Write-Warning -Message "Module $($Local:Module.Name) requires PS Version $($Local:Module.PowerShellVersion). We only found $($sessionVersion) on $($s.ComputerName). Skipping."
                continue
            }

            $destination = if ($Scope -eq 'AllUsers')
            {
                Invoke-Command -Session $s -ScriptBlock {
                    $destination = if (-not $IsLinux -and -not $IsMacOs)
                    {
                        if ($PSVersionTable.PSVersion.Major -ge 4)
                        {
                            Join-Path -Path ([System.Environment]::GetFolderPath('ProgramFiles')) -ChildPath WindowsPowerShell\Modules
                        }
                        else
                        {
                            Join-Path -Path ([System.Environment]::GetFolderPath('System')) -ChildPath WindowsPowerShell\v1.0\Modules
                        }
                    }
                    else
                    {
                        '/usr/local/share/powershell/Modules'
                    }

                    if (-not (Test-Path -Path $destination))
                    {
                        New-Item -ItemType Directory -Path $destination -Force | Out-Null
                    }

                    $destination
                }
            }
            else
            {
                Invoke-Command -Session $s -ScriptBlock { 
                    $destination = if (-not $IsLinux -and -not $IsMacOs)
                    {
                        Join-Path -Path ([System.Environment]::GetFolderPath('MyDocuments')) -ChildPath WindowsPowerShell\Modules
                    }
                    else
                    {
                        '~/.local/share/powershell/Modules'
                    }

                    if (-not (Test-Path -Path $destination))
                    {
                        New-Item -ItemType Directory -Path $destination -Force | Out-Null
                    }
                    $destination
                }
            }

            Write-Verbose "Sending psd1 manifest module in directory $($Local:Module.ModuleBase)"

            if (($Local:Module.ModuleBase -match '\d{1,4}\.\d{1,4}\.\d{1,4}\.\d{1,4}$' -or $Local:Module.ModuleBase -match '\d{1,4}\.\d{1,4}\.\d{1,4}$') -and $sessionVersion -ge ([version]::new(5,0)))
            {
                #parent folder contains a specific version. In order to copy the module right, the parent of this parent is required
                $Local:moduleParentFolder = Split-Path -Path $Local:Module.ModuleBase -Parent
            }
            else
            {
                $Local:moduleParentFolder = $Local:Module.ModuleBase
            }
            
            Send-Directory -SourceFolderPath $Local:moduleParentFolder -DestinationFolderPath $destination -Session $s

            if ($PSBoundParameters.IncludeDependencies -and ($Local:Module.RequiredAssemblies -or $Local:Module.RequiredModules))
            {
                foreach ($requiredModule in $Module.RequiredModules)
                {
                    $requiredModule = Get-Module -ListAvailable $requiredModule | Sort-Object Version -Descending | Select-Object -First 1
                    $params = ([hashtable]$PSBoundParameters).Clone()
                    [void]$params.Remove('Module')
                    Send-ModuleToPSSession -Module $requiredModule @params
                }

                foreach ($requiredAssembly in $Local:Module.RequiredAssemblies)
                {
                    if (Test-Path -Path $requiredAssembly)
                    {
                        Send-FileToPSSession -Source (Get-Item -Path $requiredAssembly -Force).FullName @fileParams
                    }
                    else
                    {
                        write-Warning "Sending required assemblies that do not have the full path information is not currently supported, $requiredAssembly not sent"
                    }
                }
            }
        }
    }
    
    end
    {
    }
}

function Split-Array
{
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.IEnumerable]$List,

        [Parameter(Mandatory = $true, ParameterSetName = 'MaxChunkSize')]
        [Alias('ChunkSize')]
        [int]$MaxChunkSize,
        
        [ValidateRange(2, [long]::MaxValue)]
        [Parameter(Mandatory = $true, ParameterSetName = 'ChunkCount')]
        [int]$ChunkCount,
        
        [switch]$AllowEmptyChunks
    )
    
    if (-not $AllowEmptyChunks -and ($list.Count -lt $ChunkCount))
    {
        Write-Error "List count ($($List.Count)) is smaller than ChunkCount ($ChunkCount).)"
        return
    }
    
    if ($PSCmdlet.ParameterSetName -eq 'MaxChunkSize')
    {        
        $ChunkCount = [Math]::Ceiling($List.Count / $MaxChunkSize)
    }
    $containers = foreach ($i in 1..$ChunkCount)
    {
        New-Object System.Collections.Generic.List[object]
    }
        
    $iContainer = 0
    foreach ($item in $List)
    {
        $containers[$iContainer].Add($item)
        $iContainer++
        if ($iContainer -ge $ChunkCount) {
            $iContainer = 0
        }
    }
        
    $containers
}

function Start-RunspaceJob
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.ScriptBlock]
        $ScriptBlock,

        [Parameter(Mandatory = $true)]
        [System.Management.Automation.Runspaces.RunspacePool]
        $RunspacePool,

        [Parameter()]
        [Object[]]
        $Argument
    )

    if ($RunspacePool.RunspacePoolStateInfo.State -eq 'Closed')
    {
        Write-Error -Message "Runspace pool $($RunspacePool.InstanceId) is already closed. Cannot queue job."
        return
    }

    if ($RunspacePool.RunspacePoolStateInfo.State -ne 'Opened')
    {
        $RunspacePool.Open()
    }

    $shell = [powershell]::Create()
    $shell.RunspacePool = $RunspacePool
    [void] $($shell.AddScript($ScriptBlock, $true))

    foreach ($arg in $Argument)
    {
        [void] $($shell.AddArgument($arg))
    }

    [PSCustomObject]@{
        Shell  = $shell
        Handle = $shell.BeginInvoke()
    }
}

function Sync-Parameter
{
    [Cmdletbinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateScript( {
                $_ -is [System.Management.Automation.FunctionInfo] -or $_ -is [System.Management.Automation.CmdletInfo] -or $_ -is [System.Management.Automation.ExternalScriptInfo]
            })]
        [object]$Command,
        
        [hashtable]$Parameters,

        [switch]$ConvertValue
    )
    
    if (-not $PSBoundParameters.ContainsKey('Parameters'))
    {
        $Parameters = ([hashtable]$ALBoundParameters).Clone()
    }
    else
    {
        $Parameters = ([hashtable]$Parameters).Clone()
    }
    
    $commonParameters = [System.Management.Automation.Internal.CommonParameters].GetProperties().Name
    $commandParameterKeys = $Command.Parameters.Keys.GetEnumerator() | ForEach-Object { $_ }
    $parameterKeys = $Parameters.Keys.GetEnumerator() | ForEach-Object { $_ }
    
    $keysToRemove = Compare-Object -ReferenceObject $commandParameterKeys -DifferenceObject $parameterKeys |
        Select-Object -ExpandProperty InputObject

    $keysToRemove = $keysToRemove + $commonParameters | Select-Object -Unique #remove the common parameters
    
    foreach ($key in $keysToRemove)
    {
        $Parameters.Remove($key)
    }

    if ($ConvertValue.IsPresent)
    {
        $keysToUpdate = @{}
        foreach ($kvp in $Parameters.GetEnumerator())
        {
            if (-not $kvp.Value) # $null or empty string will not trip up conversion
            {
                continue
            }

            $targetType = $Command.Parameters[$kvp.Key].ParameterType
            $sourceType = $kvp.Value.GetType()
            $targetValue = $kvp.Value -as $targetType

            if (-not $targetValue -and $targetType.ImplementedInterfaces -contains [Collections.IList])
            {
                $targetValue = $targetType::new()
                foreach ($v in $kvp.Value)
                {
                    $targetValue.Add($v)
                }
            }

            if (-not $targetValue -and $targetType.ImplementedInterfaces -contains [Collections.IDictionary] )
            {
                $targetValue = $targetType::new()
                foreach ($k in $kvp.Value.GetEnumerator())
                {
                    $targetValue.Add($k.Key, $k.Value)
                }
            }

            if (-not $targetValue -and ($sourceType.ImplementedInterfaces -contains [Collections.IList] -and $targetType.ImplementedInterfaces -notcontains [Collections.IList]))
            {
                Write-Verbose -Message "Value of source parameter $($kvp.Key) is a collection, but target parameter is not. Selecting first object"
                $targetValue = $kvp.Value | Select-Object -First 1
            }

            if (-not $targetValue)
            {
                Write-Error -Message "Conversion of source parameter $($kvp.Key) (Type: $($sourceType.FullName)) to type $($targetType.FullName) was impossible"
                return
            }

            $keysToUpdate[$kvp.Key] = $targetValue
        }
    }

    if ($keysToUpdate)
    {
        foreach ($kvp in $keysToUpdate.GetEnumerator())
        {
            $Parameters[$kvp.Key] = $kvp.Value
        }
    }
    
    if ($PSBoundParameters.ContainsKey('Parameters'))
    {
        $Parameters
    }
    else
    {
        $global:ALBoundParameters = $Parameters
    }
}

function Test-HashtableKeys
{
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$Hashtable,

        [string[]]$MandatoryKeys,

        [string[]]$ValidKeys,

        [switch]$Quiet
    )

    $result = $true
    
    if ($ValidKeys)
    {
        $compareResult = Compare-Object -ReferenceObject $ValidKeys -DifferenceObject ([array]$Hashtable.Keys) | Where-Object SideIndicator -eq '=>'
        if ($compareResult -and -not $Quiet)
        {
            Write-Error "The keys '$($compareResult.InputObject -join ', ')' are not valid"
        }

        $result = -not $compareResult
    }

    if ($MandatoryKeys)
    {
        $compareResult = Compare-Object -ReferenceObject $MandatoryKeys -DifferenceObject ([array]$Hashtable.Keys) | Where-Object SideIndicator -eq '<='
        if ($compareResult -and -not $Quiet)
        {
            Write-Error "The keys '$($compareResult.InputObject -join ', ')' are mandatory and not defined"
        }

        $result = -not $compareResult
    }

    $result
}

function Test-IsAdministrator
{
    
    [CmdletBinding()]
    param ()
    
    if ($IsLinux -or $IsMacOS)
    {
        # If sudo-ing or logged on as root, returns user ID 0
        $idCmd = (Get-Command -Name id).Source
        [int64] $idResult = & $idCmd -u
        $idResult -eq 0
    }
    else
    {
        $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
        (New-Object -TypeName Security.Principal.WindowsPrincipal -ArgumentList $currentUser).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
    }
}

function Wait-RunspaceJob
{
    [CmdletBinding()]
    param
    (
        [Parameter(ValueFromPipeline = $true)]
        [object[]]
        $RunspaceJob,

        [Parameter()]
        [switch]
        $PassThru
    )

    begin
    {
        $jobs = @()
    }

    process
    {
        $jobs += $RunspaceJob
    }

    end
    {
        while ($jobs.Handle.IsCompleted -contains $false)
        {
            Start-Sleep -Milliseconds 100
        }

        if ($PassThru) { $jobs }
    }
}

function Get-DscConfigurationImportedResource
{
    param(
        [Parameter(Mandatory = $true, ParameterSetName = 'ByFile')]
        [string]$FilePath,
        
        [Parameter(Mandatory = $true, ParameterSetName = 'ByConfiguration')]
        [System.Management.Automation.ConfigurationInfo]$Configuration
    )
    
    $modules = New-Object System.Collections.ArrayList

    if ($Configuration)
    {
        $ast = $Configuration.ScriptBlock.Ast
        $FilePath = $ast.FindAll( { $args[0] -is [System.Management.Automation.Language.ScriptBlockAst] }, $true)[0].Extent.File
        if (-not $FilePath)
        {
            Write-Error "The configuration '$Name' could not be found in a file. Please put the configuration into a file and try again."
            return
        }
    }
    
    $ast = [scriptblock]::Create((Get-Content -Path $FilePath -Raw)).Ast
    
    $configurations = $ast.FindAll( { $args[0] -is [System.Management.Automation.Language.ConfigurationDefinitionAst] }, $true)
    Write-Verbose "Script knwos about $($configurations.Count) configurations"
    foreach ($c in $configurations)
    {
        $importCmds = $c.Body.ScriptBlock.FindAll( { $args[0].Value -eq 'Import-DscResource' -and $args[0] -is [System.Management.Automation.Language.StringConstantExpressionAst] }, $true)
        Write-Verbose "Configuration $($c.InstanceName) knows about $($importCmds.Count) Import-DscResource commands"
    
        foreach ($importCmd in $importCmds)
        {
            $commandElements = $importCmd.Parent.CommandElements | Select-Object -Skip 1 | Where-Object {$_ -is [System.Management.Automation.Language.ArrayLiteralAst] -or $_ -is [System.Management.Automation.Language.StringConstantExpressionAst] }     
            
            $moduleNames = $commandElements.SafeGetValue()
            if ($moduleNames.GetType().IsArray)
            {
                $modules.AddRange($moduleNames)
            }
            else
            {
                [void]$modules.Add($moduleNames)
            }
        }
    }
    
    $compositeResources = $modules | Where-Object { $_ -ne 'PSDesiredStateConfiguration' } | ForEach-Object { Get-DscResource -Module $_ } | Where-Object { $_.ImplementedAs -eq 'Composite' }
    foreach ($compositeResource in $compositeResources)
    {
        $modulesInResource = Get-DscConfigurationImportedResource -FilePath $compositeResource.Path
        if ($modulesInResource)
        {
            if ($modulesInResource.GetType().IsArray)
            {
                $modules.AddRange($modulesInResource)
            }
            else
            {
                [void]$modules.Add($modulesInResource)
            }
        }
    }
    
    $modules | Select-Object -Unique
}

#author Iain Brighton, from here: https://gist.github.com/iainbrighton/9d3dd03630225ee44126769c5d9c50a9
function Get-RequiredModulesFromMOF
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [System.String] $Path
    )
    process
    {

        $modules = @{ }
        $moduleName = $null
        $moduleVersion = $null

        Get-Content -Path $Path -Encoding Unicode | ForEach-Object {
    
            $line = $_;
            if ($line -match '^\s?Instance of')
            {
                ## We have a new instance so write the existing one
                if (($null -ne $moduleName) -and ($null -ne $moduleVersion))
                {
            
                    $modules[$moduleName] = $moduleVersion;
                    $moduleName = $null
                    $moduleVersion = $null
                    Write-Verbose "Module Instance found: $moduleName $moduleVersion"
                }
            }
            elseif ($line -match '(?<=^\s?ModuleName\s?=\s?")\S+(?=";)')
            {

                ## Ignore the default PSDesiredStateConfiguration module
                if ($Matches[0] -notmatch 'PSDesiredStateConfiguration')
                {
                    $moduleName = $Matches[0]
                    Write-Verbose "Found Module Name $modulename"
                }
                else
                {
                    Write-Verbose 'Excluding PSDesiredStateConfiguration module'
                }
            }
            elseif ($line -match '(?<=^\s?ModuleVersion\s?=\s?")\S+(?=";)')
            {
                $moduleVersion = $Matches[0] -as [System.Version]
                Write-Verbose "Module version = $moduleVersion"
            }
        }

        Write-Output -InputObject $modules
    } #end process
}

function Add-HostEntry
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ParameterSetName = 'ByString')]
        [System.Net.IPAddress]$IpAddress,

        [Parameter(Mandatory, ParameterSetName = 'ByString')]
        $HostName,

        [Parameter(Mandatory, ParameterSetName = 'ByHostEntry')]
        $InputObject,

        [Parameter(Mandatory)]
        [string]$Section
    )

    if (-not $InputObject)
    {
        $InputObject = New-Object System.Net.HostRecord $IpAddress, $HostName.ToLower()
    }

    $hostContent, $hostEntries = Get-HostFile


    if ($hostEntries.Contains($InputObject))
    {
        return $false
    }

    if (($hostEntries | Where-Object HostName -eq $HostName) -and ($hostEntries | Where-Object HostName -eq $HostName).IpAddress.IPAddressToString -ne $IpAddress)
    {
        throw "Trying to add entry to hosts file with name '$HostName'. There is already another entry with this name pointing to another IP address."
    }

    $startMark = ("#$Section - start").ToLower()
    $endMark = ("#$Section - end").ToLower()

    if (-not ($hostContent | Where-Object { $_ -eq $startMark }))
    {
        $hostContent.Add($startMark) | Out-Null
        $hostContent.Add($endMark) | Out-Null
    }

    $hostContent.Insert($hostContent.IndexOf($endMark), $InputObject.ToString().ToLower())
    $hostEntries.Add($InputObject.ToString().ToLower()) | Out-Null

    $hostContent | Out-File -FilePath $script:hostFilePath

    return $true
}

function Clear-HostFile
{
    [CmdletBinding()]

    param
    (
        [Parameter(Mandatory)]
        [string]$Section
    )

    $hostContent, $hostEntries = Get-HostFile

    $startMark = ("#$Section - start").ToLower()
    $endMark = ("#$Section - end").ToLower()

    $startPosition = $hostContent.IndexOf($startMark)
    $endPosition = $hostContent.IndexOf($endMark)
    if ($startPosition -eq -1 -and $endPosition - 1)
    {
        Write-Error "Trying to remove all entries for lab from host file. However, there is no section named '$Section' defined in the hosts file which is a requirement for removing entries from this."
        return
    }

    $hostContent.RemoveRange($startPosition, $endPosition - $startPosition + 1)
    $hostContent | Out-File -FilePath $script:hostFilePath
}

Function ConvertTo-BinaryIP
{
    
	
    [CmdLetBinding()]
    Param (
        [Parameter(Mandatory = $True, Position = 0, ValueFromPipeline = $True)]
        [Net.IPAddress]$IPAddress
    )
	
    Process
    {
        Return [String]::Join('.', $($IPAddress.GetAddressBytes() |
                    ForEach-Object -Process {
                    [Convert]::ToString($_, 2).PadLeft(8, '0')
                }
            ))
    }
}

Function ConvertTo-DecimalIP
{
    
	
    [CmdLetBinding()]
    Param (
        [Parameter(Mandatory = $True, Position = 0, ValueFromPipeline = $True)]
        [Net.IPAddress]$IPAddress
    )
	
    Process
    {
        $i = 3
        $decimalIP = 0
        $IPAddress.GetAddressBytes() | ForEach-Object -Process {
            $decimalIP += $_ * [Math]::Pow(256, $i)
            $i--
        }
		
        Return [UInt32]$decimalIP
    }
}

Function ConvertTo-DottedDecimalIP
{
    
	
    [CmdLetBinding()]
    Param (
        [Parameter(Mandatory = $True, Position = 0, ValueFromPipeline = $True)]
        [String]$IPAddress
    )
	
    process
    {
        switch -RegEx ($IPAddress)
        {
            '([01]{8}\.){3}[01]{8}'
            {
                return [String]::Join('.', $($IPAddress.Split('.') | ForEach-Object -Process {
                            [Convert]::ToUInt32($_, 2)
                        }
                    ))
            }
            {$_ -as [Uint32]}
            {
                $IPAddress = [UInt32]$IPAddress
                $dottedIP = $(For ($i = 3; $i -gt -1; $i--)
                    {
                        $remainder = $IPAddress % [Math]::Pow(256, $i)
                        ($IPAddress - $remainder) / [Math]::Pow(256, $i)
                        $IPAddress = $remainder
                    }
                )
				
                return [String]::Join('.', $dottedIP)
            }
            default
            {
                throw 'Cannot convert this format'
            }
        }
    }
}

Function ConvertTo-Mask
{
    
	
    [CmdLetBinding()]
    Param (
        [Parameter(Mandatory = $True, Position = 0, ValueFromPipeline = $True)]
        [Alias('Length')]
        [ValidateRange(0, 32)]
        $MaskLength
    )
	
    Process
    {
        Return ConvertTo-DottedDecimalIP ([Convert]::ToUInt32($(('1' * $MaskLength).PadRight(32, '0')), 2))
    }
}

Function ConvertTo-MaskLength
{
    
	
    [CmdLetBinding()]
    Param (
        [Parameter(Mandatory = $True, Position = 0, ValueFromPipeline = $True)]
        [Alias('Mask')]
        [Net.IPAddress]$SubnetMask
    )
	
    Process
    {
        $Bits = "$( $SubnetMask.GetAddressBytes() | ForEach-Object  -Process { [Convert]::ToString($_, 2) 
    } )"
        $Bitsx = $Bits -Replace '[\s0]'
		
        Return $Bitsx.Length
    }
}

Function Get-BroadcastAddress
{
    
	
    [CmdLetBinding()]
    Param (
        [Parameter(Mandatory = $True, Position = 0, ValueFromPipeline = $True)]
        [Net.IPAddress]$IPAddress,
		
        [Parameter(Mandatory = $True, Position = 1)]
        [Alias('Mask')]
        [Net.IPAddress]$SubnetMask
    )
	
    process
    {
        return ConvertTo-DottedDecimalIP $((ConvertTo-DecimalIP $IPAddress) -BOr `
            ((-bnot (ConvertTo-DecimalIP $SubnetMask)) -band [UInt32]::MaxValue))
    }
}

function Get-HostEntry
{
    [CmdletBinding()]
    param (
        [Parameter(ParameterSetName = 'ByHostName')]
        [ValidateNotNullOrEmpty()][string]$HostName,

        [Parameter(ParameterSetName = 'ByIpAddress')]
        [ValidateNotNullOrEmpty()]
        [System.Net.IPAddress]$IpAddress,

        [Parameter()]
        [string]$Section
    )

    if ($Section)
    {
        $hostContent, $hostEntries = Get-HostFile -Section $Section
    }
    else
    {
        $hostContent, $hostEntries = Get-HostFile
    }

    if ($HostName)
    {
        $results = $hostEntries | Where-Object HostName -eq $HostName

        $hostEntries | Where-Object HostName -eq $HostName
    }
    elseif ($IpAddress)
    {
        $results = $hostEntries | Where-Object IpAddress -contains $IpAddress
        if (($results).count -gt 1)
        {
            Write-ScreenInfo -Message "More than one entry found in hosts file with IP address '$IpAddress' (host names: $($results.Hostname -join ','). Returning the last entry" -Type Warning
        }

        @($hostEntries | Where-Object IpAddress -contains $IpAddress)[-1]
    }
    else
    {
        $hostEntries
    }
}

function Get-HostFile
{
    [CmdletBinding()]
    param
    (
        [switch]$SuppressOutput,

        [string]$Section
    )

    $hostContent = New-Object -TypeName System.Collections.ArrayList
    $hostEntries = New-Object -TypeName System.Collections.ArrayList

    Write-PSFMessage "Opening file '$script:hostFilePath'"

    $currentHostContent = (Get-Content -Path $script:hostFilePath)
    if ($currentHostContent)
    {
        $currentHostContent = $currentHostContent.ToLower()
    }

    if ($Section)
    {
        $startMark = ("#$Section - start").ToLower()
        $endMark = ("#$Section - end").ToLower()

        if (($currentHostContent | Where-Object { $_ -eq $startMark }) -and ($currentHostContent | Where-Object { $_ -eq $endMark }))
        {
            $startPosition = $currentHostContent.IndexOf($startMark) + 1
            $endPosition = $currentHostContent.IndexOf($endMark) - 1
            $currentHostContent = $currentHostContent[$startPosition..$endPosition]
        }
        else
        {
            $currentHostContent = ''
        }
    }

    if ($currentHostContent)
    {
        $hostContent.AddRange($currentHostContent)

        foreach ($entry in $currentHostContent)
        {
            $hostfileIpAddress = [System.Text.RegularExpressions.Regex]::Matches($entry, '^(([2]([0-4][0-9]|[5][0-5])|[0-1]?[0-9]?[0-9])[.]){3}(([2]([0-4][0-9]|[5][0-5])|[0-1]?[0-9]?[0-9]))')[0].Value
            $hostfileHostName = [System.Text.RegularExpressions.Regex]::Matches($entry, '[\w\.-]+$')[0].Value

            if ($entry -notmatch '^(([2]([0-4][0-9]|[5][0-5])|[0-1]?[0-9]?[0-9])[.]){3}(([2]([0-4][0-9]|[5][0-5])|[0-1]?[0-9]?[0-9]))[\t| ]+[\w\.-]+')
            {
                continue
            }

            if (-not $hostfileIpAddress -or -not $hostfileHostName)
            {
                #could not get the IP address or hostname from current line
                continue
            }

            $newEntry = New-Object System.Net.HostRecord($hostfileIpAddress, $hostfileHostName.ToLower())
            $null = $hostEntries.Add($newEntry)
        }
    }

    Write-PSFMessage "File loaded with $($hostContent.Count) lines"

    $hostContent, $hostEntries
}

Function Get-NetworkAddress
{
    
	
    [CmdLetBinding()]
    Param (
        [Parameter(Mandatory = $True, Position = 0, ValueFromPipeline = $True)]
        [Net.IPAddress]$IPAddress,
		
        [Parameter(Mandatory = $True, Position = 1)]
        [Alias('Mask')]
        [Net.IPAddress]$SubnetMask
    )
	
    Process
    {
        Return ConvertTo-DottedDecimalIP ((ConvertTo-DecimalIP $IPAddress) -BAnd (ConvertTo-DecimalIP $SubnetMask))
    }
}

function Get-NetworkRange
{
    [CmdletBinding()]
    param 
    (
        [string]$IPAddress,
        [string]$SubnetMask
    )
	
    if ($IPAddress.Contains('/'))
    {
        $temp = $IPAddress.Split('/')
        $IPAddress = $temp[0]
        $SubnetMask = $temp[1]
    }
	
    If (-not $SubnetMask.Contains('.'))
    {
        $SubnetMask = ConvertTo-Mask -MaskLength $SubnetMask
    }
	
    $decimalIP = ConvertTo-DecimalIP -IPAddress $IPAddress
    $decimalMask = ConvertTo-DecimalIP -IPAddress $SubnetMask
	
    $network = $decimalIP -band $decimalMask
    $broadcast = $decimalIP -bor ((-bnot $decimalMask) -band [UInt32]::MaxValue)
	
    for ($i = $($network + 1); $i -lt $broadcast; $i++)
    {
        ConvertTo-DottedDecimalIP -IPAddress $i
    }
}

function Get-NetworkSummary
{
    param (
        [Parameter(Mandatory = $true)]
        [String]$IPAddress,
        [Parameter(Mandatory = $true)]
        [String]$SubnetMask
    )
    If ($IPAddress.Contains('/'))
    {
        $temp = $IP.Split('/')
        $IPAddress = $temp[0]
        $SubnetMask = $temp[1]
    }
	
    If (!$SubnetMask.Contains('.'))
    {
        $SubnetMask = ConvertTo-Mask $SubnetMask
    }
	
    $decimalIP = ConvertTo-DecimalIP $IPAddress
    $decimalMask = ConvertTo-DecimalIP $SubnetMask
	
    $network = $decimalIP -BAnd $decimalMask
    $broadcast = $decimalIP -BOr
    ((-BNot $decimalMask) -BAnd [UInt32]::MaxValue)
    $networkAddress = ConvertTo-DottedDecimalIP $network
    $rangeStart = ConvertTo-DottedDecimalIP ($network + 1)
    $rangeEnd = ConvertTo-DottedDecimalIP ($broadcast - 1)
    $broadcastAddress = ConvertTo-DottedDecimalIP $broadcast
    $MaskLength = ConvertTo-MaskLength $SubnetMask
	
    $binaryIP = ConvertTo-BinaryIP $IPAddress
    $private = $false
	
    switch -RegEx ($binaryIP)
    {
        '^1111'
        {
            $class = 'E'
            $subnetBitMap = '1111'
        }
        '^1110'
        {
            $class = 'D'
            $subnetBitMap = '1110'
        }
        '^110'
        {
            $class = 'C'
            If ($binaryIP -Match '^11000000.10101000')
            {
                $private = $True
            }
        }
        '^10'
        {
            $class = 'B'
            If ($binaryIP -Match '^10101100.0001')
            {
                $private = $True
            }
        }
        '^0'
        {
            $class = 'A'
            If ($binaryIP -Match '^00001010')
            {
                $private = $True
            }
        }
    }
	
    $netInfo = New-Object -TypeName Object
    Add-Member -MemberType NoteProperty -Name 'Network' -InputObject $netInfo -Value $networkAddress
    Add-Member -MemberType NoteProperty -Name 'Broadcast' -InputObject $netInfo -Value $broadcastAddress
    Add-Member -MemberType NoteProperty -Name 'Range' -InputObject $netInfo `
        -Value "$rangeStart - $rangeEnd"
    Add-Member -MemberType NoteProperty -Name 'Mask' -InputObject $netInfo -Value $SubnetMask
    Add-Member -MemberType NoteProperty -Name 'MaskLength' -InputObject $netInfo -Value $MaskLength
    Add-Member -MemberType NoteProperty -Name 'Hosts' -InputObject $netInfo `
        -Value $($broadcast - $network - 1)
    Add-Member -MemberType NoteProperty -Name 'Class' -InputObject $netInfo -Value $class
    Add-Member -MemberType NoteProperty -Name 'IsPrivate' -InputObject $netInfo -Value $private
	
    return $netInfo
}

function Get-PublicIpAddress
{
    [CmdletBinding()]
    param
    ()

    $ipProviderUris = @(
        'https://api.ipify.org?format=json'
        'https://ip.seeip.org/jsonip?'
        'https://api.myip.com'
    )

    foreach ($uri in $ipProviderUris)
    {
        $ip = (Invoke-RestMethod -Method Get -UseBasicParsing -Uri $uri -ErrorAction SilentlyContinue).Ip

        if ($ip)
        {
            return $ip
        }
    }
}

function Remove-HostEntry
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ParameterSetName = 'ByIpAddress')]
        [System.Net.IPAddress]$IpAddress,

        [Parameter(Mandatory, ParameterSetName = 'ByHostName')]
        $HostName,

        [Parameter(Mandatory, ParameterSetName = 'ByHostEntry')]
        $InputObject,

        [Parameter(Mandatory)]
        [string]$Section
    )

    if (-not $InputObject -and -not $IpAddress -and -not $HostName)
    {
        return
    }

    if ($InputObject)
    {
        $entriesToRemove = $InputObject
    }
    else
    {
        if (-not $InputObject -and ($IpAddress -or $HostName))
        {
            $entriesToRemove = Get-HostEntry @PSBoundParameters
        }
    }

    if (-not $entriesToRemove)
    {
        Write-Error "Trying to remove entry '$HostName' from hosts file. However, there is no entry by that name in this file"
    }

    $hostContent, $hostEntries = Get-HostFile -SuppressOutput

    $startMark = ("#$Section - start").ToLower()
    if (-not ($hostContent | Where-Object { $_ -eq $startMark }))
    {
        Write-Error "Trying to remove entry '$HostName' from hosts file. However, there is no section named '$Section' defined in the hosts file which is a requirement for removing entries from this."
        return
    }
    elseif ($entriesToRemove.Count -gt 1)
    {
        Write-Error "Trying to remove entry '$HostName' from hosts file. However, there are more than one entry with this name in the hosts file. Please remove this entry manually."
        return
    }

    if ($entriesToRemove)
    {
        $entryToRemove = ($hostContent -match "^($($entriesToRemove.IpAddress))[\t| ]+$($entriesToRemove.HostName)")[0]
        $entryToRemoveIndex = $hostContent.IndexOf($entryToRemove)

        $hostContent.RemoveAt($entryToRemoveIndex)
        $hostEntries.Remove($entriesToRemove)

        $hostContent | Out-File -FilePath $script:hostFilePath
    }
}

function Test-Port
{  
    [Cmdletbinding()]
    Param(  
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [string[]]$ComputerName,

        [Parameter(Mandatory = $true, Position = 1, ValueFromPipelineByPropertyName = $true)]
        [int]$Port,

        [int]$Count = 1,

        [int]$Delay = 500,
        
        [int]$TcpTimeout = 1000,
        [int]$UdpTimeout = 1000,
        [switch]$Tcp,
        [switch]$Udp
    )

    begin
    {  
        if (-not $Tcp -and -not $Udp)
        {
            $Tcp = $true
        }
        #Typically you never do this, but in this case I felt it was for the benefit of the function  
        #as any errors will be noted in the output of the report          
        $ErrorActionPreference = 'SilentlyContinue'
        $report = @()

        $sw = New-Object System.Diagnostics.Stopwatch
    }

    process
    {
        foreach ($c in $ComputerName)
        {
            for ($i = 0; $i -lt $Count; $i++) 
            {
                $result = New-Object PSObject | Select-Object Server, Port, TypePort, Open, Notes, ResponseTime
                $result.Server = $c
                $result.Port = $Port
                $result.TypePort = 'TCP'

                if ($Tcp)
                {
                    $tcpClient = New-Object System.Net.Sockets.TcpClient
                    $sw.Start()
                    $connect = $tcpClient.BeginConnect($c, $Port, $null, $null)
                    $wait = $connect.AsyncWaitHandle.WaitOne($TcpTimeout, $false)
                    
                    if (-not $wait)
                    {
                        $tcpClient.Close()
                        $sw.Stop()

                        $result.Open = $false
                        $result.Notes = 'Connection to Port Timed Out'
                        $result.ResponseTime = $sw.ElapsedMilliseconds
                    }
                    else
                    {
                        try
                        {
                            [void]$tcpClient.EndConnect($connect)
                            $tcpClient.Close()
                            $sw.Stop()

                            $result.Open = $true
                        }
                        catch
                        {
                            $result.Open = $false
                        }
                    }

                    $result.ResponseTime = $sw.ElapsedMilliseconds
                }
                if ($Udp)
                {
                    $udpClient = New-Object System.Net.Sockets.UdpClient
                    $udpClient.Client.ReceiveTimeout = $UdpTimeout

                    $a = New-Object System.Text.ASCIIEncoding
                    $byte = $a.GetBytes("$(Get-Date)")

                    $result.Server = $c
                    $result.Port = $Port
                    $result.TypePort = 'UDP'

                    Write-Verbose 'Making UDP connection to remote server'
                    $sw.Start()
                    $udpClient.Connect($c, $Port)
                    Write-Verbose 'Sending message to remote host'
                    [void]$udpClient.Send($byte, $byte.Length)
                    Write-Verbose 'Creating remote endpoint'
                    $remoteEndpoint = New-Object System.Net.IPEndPoint([System.Net.IPAddress]::Any, 0)

                    try
                    {
                        Write-Verbose 'Waiting for message return'
                        $receiveBytes = $udpClient.Receive([ref]$remoteEndpoint)
                        $sw.Stop()
                        [string]$returnedData = $a.GetString($receiveBytes)
                        
                        Write-Verbose 'Connection Successful'
                            
                        $result.Open = $true
                        $result.Notes = $returnedData
                    }
                    catch
                    {
                        Write-Verbose 'Host maybe unavailable'
                        $result.Open = $false
                        $result.Notes = 'Unable to verify if port is open or if host is unavailable.'
                    }
                   finally
                    {
                        $udpClient.Close()
                        $result.ResponseTime = $sw.ElapsedMilliseconds
                    }
                }

                $sw.Reset()
                $report += $result

                Start-Sleep -Milliseconds $Delay
            }
        }
    }

    end
    {
        $report 
    }
} 

function Get-PerformanceCounterID
{
    param
    (
        [Parameter(Mandatory = $true)]
        $Name
    )
 
    if ($script:perfHash -eq $null)
    {
        Write-Progress -Activity 'Retrieving PerfIDs' -Status 'Working'
 
        $key = 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Perflib\CurrentLanguage'
        $counters = (Get-ItemProperty -Path $key -Name Counter).Counter
        $script:perfHash = @{}
        $all = $counters.Count
 
        for($i = 0; $i -lt $all; $i += 2)
        {
            Write-Progress -Activity 'Retrieving PerfIDs' -Status 'Working' -PercentComplete ($i * 100 / $all)
            $script:perfHash.$($counters[$i + 1]) = $counters[$i]
        }
    }
 
    $script:perfHash.$Name
}
function Get-PerformanceCounterLocalName
{
    param
    (
        [Parameter(Mandatory = $true)]
        [UInt32]$ID,
        
        [string]$ComputerName = $env:COMPUTERNAME
    )
 
    $code = @'
    [DllImport("pdh.dll", SetLastError=true, CharSet=CharSet.Unicode)]
    public static extern UInt32 PdhLookupPerfNameByIndex(string szMachineName, uint dwNameIndex, System.Text.StringBuilder szNameBuffer, ref uint pcchNameBufferSize);
'@
 
    $buffer = New-Object System.Text.StringBuilder(1024)
    [UInt32]$bufferSize = $buffer.Capacity
 
    $t = Add-Type -MemberDefinition $code -PassThru -Name PerfCounter -Namespace Utility
    $rv = $t::PdhLookupPerfNameByIndex($ComputerName, $id, $buffer, [Ref]$bufferSize)
 
    if ($rv -eq 0)
    {
        $buffer.ToString().Substring(0, $bufferSize - 1)
    }
    else
    {
        throw 'Get-PerformanceCounterLocalName : Unable to retrieve localized name. Check computer name and performance counter ID.'
    }
}
function Get-PerformanceDataCollectorSet
{
    Param(
        [Parameter(Mandatory = $true)]
        [string]$CollectorSetName,
        
        [string]$ComputerName = 'localhost'
    )

    $collectorSet = New-Object -ComObject Pla.DataCollectorSet
    
    try
    {
        $collectorSet.Query($CollectorSetName, $ComputerName)
        return $collectorSet
    }
    catch
    {
        Write-Error -Message "Could not query data collector set. The error was: $($_.Exception.Message)" -Exception $_.Exception
    }
}
function New-PerformanceDataCollectorSet
{
    [CmdletBinding(DefaultParameterSetName = 'Counter')]
    Param(
        [Parameter(Mandatory = $true)]
        [string]$CollectorSetName,

        [datetime]$StartDate,

        [Parameter(ParameterSetName = 'Counters')]
        [string[]]$Counters,

        [Parameter(ParameterSetName = 'Xml')]
        [string[]]$XmlTemplatePath,
        
        [string]$ComputerName = 'localhost'
    )
    
    if ((Get-PerformanceDataCollectorSet -CollectorSetName $CollectorSetName -ComputerName $ComputerName -ErrorAction SilentlyContinue))
    {
        Write-Error "There is already a data collector set named '$CollectorSetName' on machine '$ComputerName'"
        return
    }

    $collectorSet = New-Object -COM Pla.DataCollectorSet
    if ($XmlTemplatePath)
    {
        if (-not (Test-Path -Path $XmlTemplatePath -PathType Leaf))
        {
            Write-Error "The file '$XmlTemplatePath' could not be found."
            return
        }
        $xml = Get-Content -Path $XmlTemplatePath
        $collectorSet.SetXml($xml)
    }
    else
    {
        $collectorSet.DisplayName = $CollectorSetName
        $collectorSet.Duration = 50400 
        $collectorSet.SubdirectoryFormat = 1 
        $collectorSet.SubdirectoryFormatPattern = 'yyyy\-MM'
        $collectorSet.RootPath = "%systemdrive%\PerfLogs\Admin\$CollectorSetName"

        $collector = $collectorSet.DataCollectors.CreateDataCollector(0) 
        $collector.FileName = $CollectorSetName + '_'
        $collector.FileNameFormat = 0x1
        $collector.FileNameFormatPattern = 'yyyy\-MM\-dd'
        $collector.SampleInterval = 15
        $collector.LogAppend = $true

        if (-not $Counters)
        {
            $Counters = @(
                '\PhysicalDisk\Avg. Disk Sec/Read',
                '\PhysicalDisk\Avg. Disk Sec/Write',
                '\PhysicalDisk\Avg. Disk Queue Length',
                '\Memory\Available MBytes', 
                '\Processor(_Total)\% Processor Time', 
                '\System\Processor Queue Length'
            )
        }

        $collector.PerformanceCounters = $Counters
        $collectorSet.DataCollectors.Add($collector)

        if ($StartDate)
        {
            $newSchedule = $collectorSet.Schedules.CreateSchedule()
            $newSchedule.Days = 127
            $newSchedule.StartDate = $StartDate
            $newSchedule.StartTime = $StartDate
    
            $collectorSet.Schedules.Add($newSchedule)
        }        
    }

    try
    {
        $collectorSet.Commit($CollectorSetName, $ComputerName, 3) | Out-Null #3 = CreateOrModify
    }
    catch
    { 
        Write-Host 'Exception Caught: ' $_.Exception -ForegroundColor Red 
        return 
    }
}
function Remove-PerformanceDataCollectorSet
{
    Param(
        [Parameter(Mandatory = $true)]
        [string]$CollectorSetName,
        
        [string]$ComputerName = 'localhost'
    )
    
    $collectorSet = Get-PerformanceDataCollectorSet -CollectorSetName $CollectorSetName -ComputerName $ComputerName -ErrorAction SilentlyContinue
    if (-not $collectorSet)
    {
        Write-Error "The data collector set '$CollectorSetName' could not be found on '$ComputerName'"
        return
    }
    
    try
    {
        $collectorSet.Delete()
    }
    catch
    {
        Write-Error -Message "Could not remove data collector set. The error was: $($_.Exception.Message)" -Exception $_.Exception
    }
}
function Start-PerformanceDataCollectorSet
{
    Param(
        [Parameter(Mandatory = $true)]
        [string]$CollectorSetName,
        
        [string]$ComputerName = 'localhost'
    )
    
    $collectorSet = Get-PerformanceDataCollectorSet -CollectorSetName $CollectorSetName -ComputerName $ComputerName -ErrorAction SilentlyContinue
    if (-not $collectorSet)
    {
        Write-Error "The data collector set '$CollectorSetName' could not be found on '$ComputerName'"
        return
    }
    
    try
    {
        $collectorSet.Start($false)
    }
    catch
    {
        Write-Error -Message "Could not start data collector set. The error was: $($_.Exception.Message)" -Exception $_.Exception
    }
}
function Stop-PerformanceDataCollectorSet
{
    Param(
        [Parameter(Mandatory = $true)]
        [string]$CollectorSetName,
        
        [string]$ComputerName = 'localhost'
    )
    
    $collectorSet = Get-PerformanceDataCollectorSet -CollectorSetName $CollectorSetName -ComputerName $ComputerName -ErrorAction SilentlyContinue
    if (-not $collectorSet)
    {
        Write-Error "The data collector set '$CollectorSetName' could not be found on '$ComputerName'"
        return
    }
    
    try
    {
        $collectorSet.Stop($false)
    }
    catch
    {
        Write-Error -Message "Could not start data collector set. The error was: $($_.Exception.Message)" -Exception $_.Exception
    }
}
function Add-CATemplateStandardPermission
{
    [cmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$TemplateName,
        
        [Parameter(Mandatory = $true)]
        [string[]]$SamAccountName
    )
    
    $configNc = ([adsi]'LDAP://RootDSE').configurationNamingContext
    $templateContainer = [adsi]"LDAP://CN=Certificate Templates,CN=Public Key Services,CN=Services,$configNc"
    Write-Verbose "Template container is '$templateContainer'"

    $template = $templateContainer.Children | Where-Object Name -eq $TemplateName
    if (-not $template)
    {
        Write-Error "The template '$TemplateName' could not be found"
        return
    }
   
    foreach ($name in $SamAccountName)
    {
        try
        {
            $sid = ([System.Security.Principal.NTAccount]$name).Translate([System.Security.Principal.SecurityIdentifier])
            $name = $sid.Translate([System.Security.Principal.NTAccount])

            dsacls $template.DistinguishedName /G "$($name):GR"
            dsacls $template.DistinguishedName /G "$($name):CA;Enroll"
            dsacls $template.DistinguishedName /G "$($name):CA;AutoEnrollment"
        }
        catch
        {
            Write-Error "The principal '$name' could not be found"
        }
    }
}

function Add-Certificate2
{
    [cmdletBinding(DefaultParameterSetName = 'ByteArray')]
    param(
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'File')]
        [string]$Path,
        
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'ByteArray')]
        [byte[]]$RawContentBytes,
        
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$Store,
        
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [System.Security.Cryptography.X509Certificates.CertStoreLocation]$Location,
        
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [string]$ServiceName,
        
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [ValidateSet('CER', 'PFX')]
        [string]$CertificateType = 'CER',
        
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [string]$Password
    )
    
    process
    {
        if ($Location -eq 'CERT_SYSTEM_STORE_SERVICES' -and (-not $ServiceName))
        {
            Write-Error "Please specify a ServiceName if the Location is set to 'CERT_SYSTEM_STORE_SERVICES'"
            return
        }
    
        $storePath = $Store
        
        if ($Path -and -not (Test-Path -Path $Path))
        {
            Write-Error "The path '$Path' does not exist."
            continue
        }
        
        if ($ServiceName)
        {
            if (-not (Get-Service -Name $ServiceName))
            {
                Write-Error "The service '$ServiceName' could not be found."
                return
            }
            else
            {
                $storePath = "$ServiceName\$Store"
            }
        }
    
        $storeProvider = [System.Security.Cryptography.X509Certificates.CertStoreProvider]::CERT_STORE_PROV_SYSTEM_REGISTRY

        $Location = $Location -bor [System.Security.Cryptography.X509Certificates.CertOpenStoreFlags]::CERT_STORE_MAXIMUM_ALLOWED_FLAG
    
        $storePtr = [System.Security.Cryptography.X509Certificates.Win32]::CertOpenStore($storeProvider, 0, 0, $Location, $storePath)
        if ($storePtr -eq [System.IntPtr]::Zero)
        {
            Write-Error "Store '$Store' in location '$Location' could not be opened."
            return
        }
        $s = New-Object System.Security.Cryptography.X509Certificates.X509Store($storePtr)
        
        if ($Path)
        {
            $RawContentBytes = [System.IO.File]::ReadAllBytes($Path)
        }
        
        try
        {
            if ($Password)
            {
                $securePassword = $Password | ConvertTo-SecureString -AsPlainText -Force
            }
            $certInfo = if ([System.Security.Cryptography.X509Certificates.X509Certificate2]::GetCertContentType($RawContentBytes) -eq 'Pfx')
            {
                New-Object Pki.Certificates.CertificateInfo($RawContentBytes, $securePassword)
            }
            else
            {
                New-Object Pki.Certificates.CertificateInfo(,$RawContentBytes)
            }
        }
        catch
        {
            Write-Error -ErrorRecord $_
            return
        }

        Write-Verbose "Store '$Store' in location '$Location' knowns about $($s.Certificates.Count) certificates before import."
        
        $s.Add($certInfo.Certificate)
        
        Write-Verbose "Store '$Store' in location '$Location' knowns about $($s.Certificates.Count) certificates after import."

        [void][System.Security.Cryptography.X509Certificates.Win32]::CertCloseStore($storePtr, 0)
    }
}

function Enable-AutoEnrollment
{
    param
    (
        [switch]$Computer,
        [switch]$UserOrCodeSigning
    )
    
    Write-Verbose -Message "Computer: '$Computer'"
    Write-Verbose -Message "Computer: '$UserOrCodeSigning'"

    if ($PSEdition -eq 'Core') 
    { 
        Write-Warning -Message 'Cannot execute Enable-AutoEnrollment on PowerShell Core!'
        return 
    }
    
    if ($Computer)
    {
        Write-Verbose -Message 'Configuring for computer auto enrollment'
        [GPO.Helper]::SetGroupPolicy($true, 'Software\Policies\Microsoft\Cryptography\AutoEnrollment', 'AEPolicy', 7)
        [GPO.Helper]::SetGroupPolicy($true, 'Software\Policies\Microsoft\Cryptography\AutoEnrollment', 'OfflineExpirationPercent', 10)
        [GPO.Helper]::SetGroupPolicy($true, 'Software\Policies\Microsoft\Cryptography\AutoEnrollment', 'OfflineExpirationStoreNames', 'MY')
    }
    if ($UserOrCodeSigning)
    {
        Write-Verbose -Message 'Configuring for user auto enrollment'
        [GPO.Helper]::SetGroupPolicy($false, 'Software\Policies\Microsoft\Cryptography\AutoEnrollment', 'AEPolicy', 7)
        [GPO.Helper]::SetGroupPolicy($false, 'Software\Policies\Microsoft\Cryptography\AutoEnrollment', 'OfflineExpirationPercent', 10)
        [GPO.Helper]::SetGroupPolicy($false, 'Software\Policies\Microsoft\Cryptography\AutoEnrollment', 'OfflineExpirationStoreNames', 'MY')
    }
    
    1..3 | ForEach-Object { gpupdate.exe /force; certutil.exe -pulse; Start-Sleep -Seconds 1 }
}

function Find-CertificateAuthority
{
    [cmdletBinding()]
    param(
        [string]$DomainName
    )

    Add-Type -AssemblyName System.DirectoryServices.AccountManagement

    try
    {
        $ctx = New-Object System.DirectoryServices.AccountManagement.PrincipalContext('Domain', $DomainName)
    }
    catch
    {
        Write-Error "The domain '$DomainName' could not be contacted"
        return
    }
    
    try
    {
        $configDn = ([ADSI]'LDAP://RootDSE').configurationNamingContext
        $cdpContainer = [ADSI]"LDAP://CN=CDP,CN=Public Key Services,CN=Services,$configDn"

        if (-not $cdpContainer)
        {
            Write-Error 'Could not connect to CDP container' -ErrorAction Stop
        }
    }
    catch
    {
        Write-Error "The domain '$DomainName' could not be contacted" -TargetObject $DomainName
        return
    }
                
    $caFound = $false
    foreach ($item in $cdpContainer.Children)
    {
        if (-not $caFound)
        {
            $machine = ($item.distinguishedName -split '=|,')[1]
            $caName = ($item.Children.distinguishedName -split '=|,')[1]

            if ($DomainName)
            {
                $group = [System.DirectoryServices.AccountManagement.GroupPrincipal]::FindByIdentity($ctx, 'Cert Publishers')
                $machine = $group.Members | Where-Object Name -eq $machine
                if ($machine.Context.Name -ne $DomainName)
                {
                    continue
                }
            }
                        
            $certificateAuthority = "$machine\$caName"
                        
            $result = certutil.exe -ping $certificateAuthority
            if ($result -match 'interface is alive*' )
            {
                $caFound = $true
            }
        }
    }
    
    if ($caFound)
    {
        $certificateAuthority
    }
    else
    {
        Write-Error "No Certificate Authority could be found in domain '$DomainName'"
    }
}
function Get-CATemplate
{
    [cmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$TemplateName
    )
    
    $configNc = ([adsi]'LDAP://RootDSE').configurationNamingContext
    $templateContainer = [adsi]"LDAP://CN=Certificate Templates,CN=Public Key Services,CN=Services,$configNc"
    Write-Verbose "Template container is '$($templateContainer.distinguishedName)'"

    $templateContainer.Children | Where-Object Name -eq $TemplateName
}

function Get-Certificate2
{
    [cmdletBinding(DefaultParameterSetName = 'FindCer')]
    param (
        [Parameter(Mandatory = $true, ParameterSetName = 'FindCer')]
        [Parameter(Mandatory = $true, ParameterSetName = 'FindPfx')]
        [string]$SearchString,

        [Parameter(Mandatory = $true, ParameterSetName = 'FindCer')]
        [Parameter(Mandatory = $true, ParameterSetName = 'FindPfx')]        
        [System.Security.Cryptography.X509Certificates.X509FindType]$FindType,
        
        [Parameter(ParameterSetName = 'AllCer')]
        [Parameter(ParameterSetName = 'AllPfx')]
        [Parameter(ParameterSetName = 'FindCer')]
        [Parameter(ParameterSetName = 'FindPfx')]
        [System.Security.Cryptography.X509Certificates.CertStoreLocation]$Location,
        
        [Parameter(ParameterSetName = 'AllCer')]
        [Parameter(ParameterSetName = 'AllPfx')]
        [Parameter(ParameterSetName = 'FindCer')]
        [Parameter(ParameterSetName = 'FindPfx')]
        [string]$Store,
        
        [Parameter(ParameterSetName = 'AllCer')]
        [Parameter(ParameterSetName = 'AllPfx')]
        [Parameter(ParameterSetName = 'FindCer')]
        [Parameter(ParameterSetName = 'FindPfx')]
        [string]$ServiceName,

        [Parameter(Mandatory = $true, ParameterSetName = 'AllCer')]
        [Parameter(Mandatory = $true, ParameterSetName = 'AllPfx')]
        [switch]$All,

        [Parameter(ParameterSetName = 'AllCer')]
        [Parameter(ParameterSetName = 'AllPfx')]
        [switch]$IncludeServices,
        
        [Parameter(Mandatory = $true, ParameterSetName = 'FindPfx')]
        [Parameter(Mandatory = $true, ParameterSetName = 'AllPfx')]
        [securestring]$Password,
        
        [Parameter(ParameterSetName = 'FindPfx')]
        [Parameter(ParameterSetName = 'AllPfx')]
        [switch]$ExportPrivateKey
    )
    
    $services = Get-Service
    
    if ($ServiceName -and $Location -ne 'CERT_SYSTEM_STORE_SERVICES')
    {
        $Location = 'CERT_SYSTEM_STORE_SERVICES'
    }
    
    if ($ServiceName -and $ServiceName -notin $services.Name)
    {
        Write-Error "The service '$ServiceName' could not be found."
        return
    }
    
    $storeProvider = [System.Security.Cryptography.X509Certificates.CertStoreProvider]::CERT_STORE_PROV_SYSTEM
    
    $certs = foreach ($currentLocation in [Enum]::GetNames([System.Security.Cryptography.X509Certificates.CertStoreLocation]))
    {
        if ($Location -and $Location -ne $currentLocation)
        {
            Write-Verbose "Skipping location '$currentLocation'"
            continue
        }
        Write-Verbose "Enumerating stores location '$currentLocation'"

        $internalLocation = [System.Security.Cryptography.X509Certificates.CertStoreLocation]$currentLocation -bor [System.Security.Cryptography.X509Certificates.CertOpenStoreFlags]::CERT_STORE_READONLY_FLAG
    
        $availableStores = if ($ServiceName)
        {
            [System.Security.Cryptography.X509Certificates.Win32]::GetServiceCertificateStores($ServiceName)
        }
        elseif ($Location -eq [System.Security.Cryptography.X509Certificates.CertStoreLocation]::CERT_SYSTEM_STORE_SERVICES)
        {
            $services = Get-Service
            foreach ($Service in $services)
            {
                [System.Security.Cryptography.X509Certificates.Win32]::GetServiceCertificateStores($service.Name)
            }
        }
        else
        {
            [System.Security.Cryptography.X509Certificates.Win32]::GetCertificateStores()
        }
        
        $availableStores = if ($Location -eq [System.Security.Cryptography.X509Certificates.CertStoreLocation]::CERT_SYSTEM_STORE_CURRENT_USER)
        {
            $availableStores | Where-Object Location -eq CurrentUser
        }
        elseif ($Location -eq [System.Security.Cryptography.X509Certificates.CertStoreLocation]::CERT_SYSTEM_STORE_LOCAL_MACHINE)
        {
            $availableStores | Where-Object Location -eq LocalMachine
        }
        elseif ($Location -eq [System.Security.Cryptography.X509Certificates.CertStoreLocation]::CERT_SYSTEM_STORE_LOCAL_MACHINE)
        {
            $availableStores | Where-Object Location -eq LocalMachine
        }
        elseif ($Location -eq [System.Security.Cryptography.X509Certificates.CertStoreLocation]::CERT_SYSTEM_STORE_SERVICES)
        {
            $availableStores | Where-Object Location -eq Services
        }
        elseif ($Location -eq [System.Security.Cryptography.X509Certificates.CertStoreLocation]::CERT_SYSTEM_STORE_USERS)
        {
            $availableStores | Where-Object Location -eq Users
        }
        elseif ($Location -eq [System.Security.Cryptography.X509Certificates.CertStoreLocation]::CERT_SYSTEM_STORE_CURRENT_USER_GROUP_POLICY)
        {
            $availableStores | Where-Object Location -eq CurrentUserGroupPolicy
        }
        elseif ($Location -eq [System.Security.Cryptography.X509Certificates.CertStoreLocation]::CERT_SYSTEM_STORE_LOCAL_MACHINE_GROUP_POLICY)
        {
            $availableStores | Where-Object Location -eq LocalMachineGroupPolicy
        }
        elseif ($Location -eq [System.Security.Cryptography.X509Certificates.CertStoreLocation]::CERT_SYSTEM_STORE_LOCAL_MACHINE_ENTERPRISE)
        {
            $availableStores | Where-Object Location -eq LocalMachineEnterprise
        }
        else
        {
            $availableStores
        }
        
        if ($Store)
        {
            if ($ServiceName)
            {
                if ("$ServiceName\$Store" -notin $availableStores.Name)
                {
                    Write-Error "The store '$Store' does not exist for location '$currentLocation' for service '$ServiceName'"
                    continue
                }
                else
                {
                    $availableStores = $availableStores | Where-Object Name -eq "$ServiceName\$Store"
                }
                
            }
            else
            {
                if ($Store -notin $availableStores.Name)
                {
                    Write-Error "The store '$Store' does not exist for location '$currentLocation'"
                    continue
                }
                else
                {
                    $availableStores = $availableStores | Where-Object Name -eq $Store
                }
            }
        }
            
        foreach ($storePath in $availableStores)
        {
            Write-Verbose "Enumerating certificates in store '$storePath' in location '$currentLocation'"
                
            $storePtr = [System.Security.Cryptography.X509Certificates.Win32]::CertOpenStore($storeProvider, 0, 0, $internalLocation, $storePath.Name)
            if ($storePtr -eq [System.IntPtr]::Zero)
            {
                Write-Verbose "Store '$storePath' in location '$currentLocation' could not be opened."
                continue
            }
            
            $s = New-Object System.Security.Cryptography.X509Certificates.X509Store($storePtr)
            $result = if ($All)
            {
                $s.Certificates
            }
            else
            {
                $s.Certificates.Find($FindType, $SearchString, $false)
            }
                
            foreach ($item in $result)
            {
                $item | Add-Member -MemberType NoteProperty -Name Location -Value $currentLocation
                $item | Add-Member -MemberType NoteProperty -Name Store -Value $storePath
                $item | Add-Member -MemberType NoteProperty -Name Password -Value $plainPassword
                    
                if ($Location -eq 'CERT_SYSTEM_STORE_SERVICES')
                {
                    $item | Add-Member -MemberType NoteProperty -Name ServiceName -Value ($storePath -split '\\')[0]
                    $item | Add-Member -MemberType NoteProperty -Name Store -Value ($storePath -split '\\')[1] -Force
                }
                    
                $item
            }

            [void][System.Security.Cryptography.X509Certificates.Win32]::CertCloseStore($storePtr, 0)
        }
    }

    Write-Verbose "Found $($certs.Count) certificates"
    
    if ($SearchString -and $certs.Count -eq 0)
    {
        Write-Error "No certificate found applying search string '$SearchString' and looking for '$FindType'"
        return
    }
    
    foreach ($cert in $certs)
    {
        $tempFile = [System.IO.Path]::GetTempFileName()
        Remove-Item -Path $tempFile

        Write-Verbose "Current certificate is $($cert.Thumbprint)"

        try
        {
            if ($cert.HasPrivateKey -and $ExportPrivateKey)
            {
                Write-Verbose 'Calling Export-PfxCertificate'
                Export-PfxCertificate -Cert $cert -FilePath $tempFile -Password $Password -ErrorAction Stop | Out-Null
            }
            else
            {
                Write-Verbose 'Calling Export-Certificate'
                Export-Certificate -Cert $cert -FilePath $tempFile -ErrorAction Stop | Out-Null
            }
            Write-Verbose 'Export finished'
        }
        catch
        {
            if ($SearchString) #A specific cert is desired so an error is written as not in list mode
            {
                Write-Error $_
            }
            continue
        }

        $certInfo = if ($ExportPrivateKey)
        {
            New-Object Pki.Certificates.CertificateInfo($tempFile, $Password)
        }
        else
        {
            New-Object Pki.Certificates.CertificateInfo($tempFile)
        }
        Remove-Item -Path $tempFile
        
        $certInfo.ComputerName = $env:COMPUTERNAME
        $certInfo.Location = $cert.Location
        $certInfo.Store = $cert.Store.Name
        $certInfo.ServiceName = $cert.ServiceName
        
        $certInfo
    }
}

function Get-NextOid
{
    param(
        [Parameter(Mandatory = $true)]
        [string]$Oid
    )
    
    $oidRange = $Oid.Substring(0, $Oid.LastIndexOf('.'))
    $lastNumber = $Oid.Substring($Oid.LastIndexOf('.') + 1)
    '{0}.{1}' -f $oidRange, ([int]$lastNumber + 1)
}

$ApplicationPolicies = @{
    # Remote Desktop
    'Remote Desktop'                            = '1.3.6.1.4.1.311.54.1.2'
    # Windows Update
    'Windows Update'                            = '1.3.6.1.4.1.311.76.6.1'
    # Windows Third Party Applicaiton Component
    'Windows Third Party Application Component' = '1.3.6.1.4.1.311.10.3.25'
    # Windows TCB Component
    'Windows TCB Component'                     = '1.3.6.1.4.1.311.10.3.23'
    # Windows Store
    'Windows Store'                             = '1.3.6.1.4.1.311.76.3.1'
    # Windows Software Extension verification
    ' Windows Software Extension Verification'  = '1.3.6.1.4.1.311.10.3.26'
    # Windows RT Verification
    'Windows RT Verification'                   = '1.3.6.1.4.1.311.10.3.21'
    # Windows Kits Component
    'Windows Kits Component'                    = '1.3.6.1.4.1.311.10.3.20'
    # ROOT_PROGRAM_NO_OCSP_FAILOVER_TO_CRL
    'No OCSP Failover to CRL'                   = '1.3.6.1.4.1.311.60.3.3'
    # ROOT_PROGRAM_AUTO_UPDATE_END_REVOCATION
    'Auto Update End Revocation'                = '1.3.6.1.4.1.311.60.3.2'
    # ROOT_PROGRAM_AUTO_UPDATE_CA_REVOCATION
    'Auto Update CA Revocation'                 = '1.3.6.1.4.1.311.60.3.1'
    # Revoked List Signer
    'Revoked List Signer'                       = '1.3.6.1.4.1.311.10.3.19'
    # Protected Process Verification
    'Protected Process Verification'            = '1.3.6.1.4.1.311.10.3.24'
    # Protected Process Light Verification
    'Protected Process Light Verification'      = '1.3.6.1.4.1.311.10.3.22'
    # Platform Certificate
    'Platform Certificate'                      = '2.23.133.8.2'
    # Microsoft Publisher
    'Microsoft Publisher'                       = '1.3.6.1.4.1.311.76.8.1'
    # Kernel Mode Code Signing
    'Kernel Mode Code Signing'                  = '1.3.6.1.4.1.311.6.1.1'
    # HAL Extension
    'HAL Extension'                             = '1.3.6.1.4.1.311.61.5.1'
    # Endorsement Key Certificate
    'Endorsement Key Certificate'               = '2.23.133.8.1'
    # Early Launch Antimalware Driver
    'Early Launch Antimalware Driver'           = '1.3.6.1.4.1.311.61.4.1'
    # Dynamic Code Generator
    'Dynamic Code Generator'                    = '1.3.6.1.4.1.311.76.5.1'
    # Domain Name System (DNS) Server Trust
    'DNS Server Trust'                          = '1.3.6.1.4.1.311.64.1.1'
    # Document Encryption
    'Document Encryption'                       = '1.3.6.1.4.1.311.80.1'
    # Disallowed List
    'Disallowed List'                           = '1.3.6.1.4.1.10.3.30'
    # Attestation Identity Key Certificate
    # System Health Authentication
    'System Health Authentication'              = '1.3.6.1.4.1.311.47.1.1'
    # Smartcard Logon
    'IdMsKpScLogon'                             = '1.3.6.1.4.1.311.20.2.2'
    # Certificate Request Agent
    'ENROLLMENT_AGENT'                          = '1.3.6.1.4.1.311.20.2.1'
    # CTL Usage
    'AUTO_ENROLL_CTL_USAGE'                     = '1.3.6.1.4.1.311.20.1'
    # Private Key Archival
    'KP_CA_EXCHANGE'                            = '1.3.6.1.4.1.311.21.5'
    # Key Recovery Agent
    'KP_KEY_RECOVERY_AGENT'                     = '1.3.6.1.4.1.311.21.6'
    # Secure Email
    'PKIX_KP_EMAIL_PROTECTION'                  = '1.3.6.1.5.5.7.3.4'
    # IP Security End System
    'PKIX_KP_IPSEC_END_SYSTEM'                  = '1.3.6.1.5.5.7.3.5'
    # IP Security Tunnel Termination
    'PKIX_KP_IPSEC_TUNNEL'                      = '1.3.6.1.5.5.7.3.6'
    # IP Security User
    'PKIX_KP_IPSEC_USER'                        = '1.3.6.1.5.5.7.3.7'
    # Time Stamping
    'PKIX_KP_TIMESTAMP_SIGNING'                 = '1.3.6.1.5.5.7.3.8'
    # OCSP Signing
    'KP_OCSP_SIGNING'                           = '1.3.6.1.5.5.7.3.9'
    # IP security IKE intermediate
    'IPSEC_KP_IKE_INTERMEDIATE'                 = '1.3.6.1.5.5.8.2.2'
    # Microsoft Trust List Signing
    'KP_CTL_USAGE_SIGNING'                      = '1.3.6.1.4.1.311.10.3.1'
    # Microsoft Time Stamping
    'KP_TIME_STAMP_SIGNING'                     = '1.3.6.1.4.1.311.10.3.2'
    # Windows Hardware Driver Verification
    'WHQL_CRYPTO'                               = '1.3.6.1.4.1.311.10.3.5'
    # Windows System Component Verification
    'NT5_CRYPTO'                                = '1.3.6.1.4.1.311.10.3.6'
    # OEM Windows System Component Verification
    'OEM_WHQL_CRYPTO'                           = '1.3.6.1.4.1.311.10.3.7'
    # Embedded Windows System Component Verification
    'EMBEDDED_NT_CRYPTO'                        = '1.3.6.1.4.1.311.10.3.8'
    # Root List Signer
    'ROOT_LIST_SIGNER'                          = '1.3.6.1.4.1.311.10.3.9'
    # Qualified Subordination
    'KP_QUALIFIED_SUBORDINATION'                = '1.3.6.1.4.1.311.10.3.10'
    # Key Recovery
    'KP_KEY_RECOVERY'                           = '1.3.6.1.4.1.311.10.3.11'
    # Document Signing
    'KP_DOCUMENT_SIGNING'                       = '1.3.6.1.4.1.311.10.3.12'
    # Lifetime Signing
    'KP_LIFETIME_SIGNING'                       = '1.3.6.1.4.1.311.10.3.13'
    'DRM'                                       = '1.3.6.1.4.1.311.10.5.1'
    'DRM_INDIVIDUALIZATION'                     = '1.3.6.1.4.1.311.10.5.2'
    # Key Pack Licenses
    'LICENSES'                                  = '1.3.6.1.4.1.311.10.6.1'
    # License Server Verification
    'LICENSE_SERVER'                            = '1.3.6.1.4.1.311.10.6.2'
    'Server Authentication'                     = '1.3.6.1.5.5.7.3.1' #The certificate can be used for OCSP authentication.            
    KP_IPSEC_USER                               = '1.3.6.1.5.5.7.3.7' #The certificate can be used for an IPSEC user.            
    'Code Signing'                              = '1.3.6.1.5.5.7.3.3' #The certificate can be used for signing code.
    'Client Authentication'                     = '1.3.6.1.5.5.7.3.2' #The certificate can be used for authenticating a client.
    KP_EFS                                      = '1.3.6.1.4.1.311.10.3.4' #The certificate can be used to encrypt files by using the Encrypting File System.
    EFS_RECOVERY                                = '1.3.6.1.4.1.311.10.3.4.1' #The certificate can be used for recovery of documents protected by using Encrypting File System (EFS).
    DS_EMAIL_REPLICATION                        = '1.3.6.1.4.1.311.21.19' #The certificate can be used for Directory Service email replication.         
    ANY_APPLICATION_POLICY                      = '1.3.6.1.4.1.311.10.12.1' #The applications that can use the certificate are not restricted.
}
function New-CATemplate
{
    [cmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$TemplateName,
        
        [string]$DisplayName,
        
        [Parameter(Mandatory = $true)]
        [string]$SourceTemplateName,
        
        [ValidateSet('EFS_RECOVERY', 'Auto Update CA Revocation', 'No OCSP Failover to CRL', 'OEM_WHQL_CRYPTO', 'Windows TCB Component', 'DNS Server Trust', 'Windows Third Party Application Component', 'ANY_APPLICATION_POLICY', 'KP_LIFETIME_SIGNING', 'Disallowed List', 'DS_EMAIL_REPLICATION', 'LICENSE_SERVER', 'KP_KEY_RECOVERY', 'Windows Kits Component', 'AUTO_ENROLL_CTL_USAGE', 'PKIX_KP_TIMESTAMP_SIGNING', 'Windows Update', 'Document Encryption', 'KP_CTL_USAGE_SIGNING', 'IPSEC_KP_IKE_INTERMEDIATE', 'PKIX_KP_IPSEC_TUNNEL', 'Code Signing', 'KP_KEY_RECOVERY_AGENT', 'KP_QUALIFIED_SUBORDINATION', 'Early Launch Antimalware Driver', 'Remote Desktop', 'WHQL_CRYPTO', 'EMBEDDED_NT_CRYPTO', 'System Health Authentication', 'DRM', 'PKIX_KP_EMAIL_PROTECTION', 'KP_TIME_STAMP_SIGNING', 'Protected Process Light Verification', 'Endorsement Key Certificate', 'KP_IPSEC_USER', 'PKIX_KP_IPSEC_END_SYSTEM', 'LICENSES', 'Protected Process Verification', 'IdMsKpScLogon', 'HAL Extension', 'KP_OCSP_SIGNING', 'Server Authentication', 'Auto Update End Revocation', 'KP_EFS', 'KP_DOCUMENT_SIGNING', 'Windows Store', 'Kernel Mode Code Signing', 'ENROLLMENT_AGENT', 'ROOT_LIST_SIGNER', 'Windows RT Verification', 'NT5_CRYPTO', 'Revoked List Signer', 'Microsoft Publisher', 'Platform Certificate', ' Windows Software Extension Verification', 'KP_CA_EXCHANGE', 'PKIX_KP_IPSEC_USER', 'Dynamic Code Generator', 'Client Authentication', 'DRM_INDIVIDUALIZATION')]
        [string[]]$ApplicationPolicy,

        [Pki.CATemplate.EnrollmentFlags]$EnrollmentFlags = 'None',

        [Pki.CATemplate.PrivateKeyFlags]$PrivateKeyFlags = 0,

        [Pki.CATemplate.KeyUsage]$KeyUsage = 0,
        
        [int]$Version,

        [timespan]$ValidityPeriod,
        
        [timespan]$RenewalPeriod
    )

    $configNc = ([adsi]'LDAP://RootDSE').ConfigurationNamingContext
    $templateContainer = [adsi]"LDAP://CN=Certificate Templates,CN=Public Key Services,CN=Services,$configNc"
    Write-Verbose "Template container is '$templateContainer'"

    $sourceTemplate = $templateContainer.Children | Where-Object Name -eq $SourceTemplateName
    if (-not $sourceTemplate)
    {
        Write-Error "The source template '$SourceTemplateName' could not be found"
        return
    }

    if (($templateContainer.Children | Where-Object Name -eq $TemplateName))
    {
        Write-Error "The template '$TemplateName' does aleady exist"
        return
    }
    
    if (-not $DisplayName) { $DisplayName = $TemplateName }
    
    $newCertTemplate = $templateContainer.Create('pKICertificateTemplate', "CN=$TemplateName") 
    $newCertTemplate.put('distinguishedName', "CN=$TemplateName,CN=Certificate Templates,CN=Public Key Services,CN=Services,$configNc")

    $lastOid = $templateContainer.Children | 
        Sort-Object -Property { [int]($_.'msPKI-Cert-Template-OID' -split '\.')[-1] } | 
        Select-Object -Last 1 -ExpandProperty msPKI-Cert-Template-OID
    $oid = Get-NextOid -Oid $lastOid
    
    $flags = $sourceTemplate.flags.Value
    $flags = $flags -bor [Pki.CATemplate.Flags]::IsModified -bxor [Pki.CATemplate.Flags]::IsDefault
    
    $newCertTemplate.put('flags', $flags)
    $newCertTemplate.put('displayName', $DisplayName)
    $newCertTemplate.put('revision', '100')
    $newCertTemplate.put('pKIDefaultKeySpec', $sourceTemplate.pKIDefaultKeySpec.Value)

    $newCertTemplate.put('pKIMaxIssuingDepth', $sourceTemplate.pKIMaxIssuingDepth.Value)
    $newCertTemplate.put('pKICriticalExtensions', $sourceTemplate.pKICriticalExtensions.Value)
    
    $eku = @($sourceTemplate.pKIExtendedKeyUsage.Value)
    $newCertTemplate.put('pKIExtendedKeyUsage', $eku)
    
    #$newCertTemplate.put('pKIDefaultCSPs','2,Microsoft Base Cryptographic Provider v1.0, 1,Microsoft Enhanced Cryptographic Provider v1.0')
    $newCertTemplate.put('msPKI-RA-Signature', '0')
    $newCertTemplate.put('msPKI-Enrollment-Flag', $EnrollmentFlags)
    $newCertTemplate.put('msPKI-Private-Key-Flag', $PrivateKeyFlags)
    $newCertTemplate.put('msPKI-Certificate-Name-Flag', $sourceTemplate.'msPKI-Certificate-Name-Flag'.Value)
    $newCertTemplate.put('msPKI-Minimal-Key-Size', $sourceTemplate.'msPKI-Minimal-Key-Size'.Value)
    
    if (-not $Version)
    {
        $Version = $sourceTemplate.'msPKI-Template-Schema-Version'.Value
    }
    $newCertTemplate.put('msPKI-Template-Schema-Version', $Version)
    $newCertTemplate.put('msPKI-Template-Minor-Revision', '1')
                   
    $newCertTemplate.put('msPKI-Cert-Template-OID', $oid)
    
    if (-not $ApplicationPolicy)
    {
        #V2 template
        $ap = $sourceTemplate.'msPKI-Certificate-Application-Policy'.Value
        if (-not $ap)
        {
            #V1 template
            $ap = $sourceTemplate.pKIExtendedKeyUsage.Value
        }
    }
    else
    {
        $ap = $ApplicationPolicy | ForEach-Object { $ApplicationPolicies[$_] }
    }
    
    if ($ap)
    {
        $newCertTemplate.put('msPKI-Certificate-Application-Policy', $ap)
    }
    
    $newCertTemplate.SetInfo()

    if ($KeyUsage)
    {
        $newCertTemplate.pKIKeyUsage = $KeyUsage
    }
    else
    {
        $newCertTemplate.pKIKeyUsage = $sourceTemplate.pKIKeyUsage
    }
    
    if ($ValidityPeriod)
    {
        $newCertTemplate.pKIExpirationPeriod.Value = [Pki.Period]::ToByteArray($ValidityPeriod)
    }
    else
    {
        $newCertTemplate.pKIExpirationPeriod = $sourceTemplate.pKIExpirationPeriod
    }
    
    if ($RenewalPeriod)
    {
        $newCertTemplate.pKIOverlapPeriod.Value = [Pki.Period]::ToByteArray($RenewalPeriod)
    }
    else
    {
        $newCertTemplate.pKIOverlapPeriod = $sourceTemplate.pKIOverlapPeriod
    }    
    $newCertTemplate.SetInfo()
}

function Publish-CaTemplate
{
    [cmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$TemplateName
    )
    
    $ca = Find-CertificateAuthority
    $caInfo = certutil.exe -CAInfo -Config $ca
    if ($caInfo -like '*No local Certification Authority*')
    {
        Write-Error 'No issuing CA found in the machines domain'
        return
    }
    $computerName = $ca.Split('\')[0]

    $start = Get-Date
    $done = $false
    $i = 0
    do
    {
        Write-Verbose -Message "Trying to publish '$TemplateName' on '$ca' at ($(Get-Date)), retry count $i"
        certutil.exe -Config $ca -SetCAtemplates "+$TemplateName" | Out-Null
        if (-not $LASTEXITCODE)
        {
            $done = $true
        }
        else
        {
            if ($i % 5 -eq 0)
            {
                Get-Service -Name CertSvc -ComputerName $computerName | Restart-Service
            }

            $ex = New-Object System.ComponentModel.Win32Exception($LASTEXITCODE)
            Write-Verbose -Message "Publishing the template '$TemplateName' failed: $($ex.Message)"

            Start-Sleep -Seconds 10
            $i++
        }
    }
    until ($done -or ((Get-Date) - $start).TotalMinutes -ge 10)
    Write-Verbose -Message "Certificate templete '$TemplateName' published successfully"

    if ($LASTEXITCODE)
    {
        $ex = New-Object System.ComponentModel.Win32Exception($LASTEXITCODE)
        Write-Error -Message "Publishing the template '$TemplateName' failed: $($ex.Message)" -Exception $ex
        return
    }

    Write-Verbose "Successfully published template '$TemplateName'"
}

function Request-Certificate
{
    [cmdletBinding()]
    param (
        [Parameter(Mandatory = $true, HelpMessage = 'Please enter the subject beginning with CN=')]
        [ValidatePattern('CN=')]
        [string]$Subject,

        [Parameter(HelpMessage = 'Please enter the SAN domains as a comma separated list')]
        [string[]]$SAN,

        [Parameter(HelpMessage = 'Please enter the Online Certificate Authority')]
        [string]$OnlineCA,

        [Parameter(Mandatory = $true, HelpMessage = 'Please enter the Online Certificate Authority')]
        [string]$TemplateName
    )

    $infFile = [System.IO.Path]::GetTempFileName()
    $requestFile = [System.IO.Path]::GetTempFileName()
    $certFile = [System.IO.Path]::GetTempFileName()
    $rspFile = [System.IO.Path]::GetTempFileName()

    ### INI file generation
    $iniContent = @'
[Version]
Signature="$Windows NT$"

[NewRequest]
Subject="{0}"
Exportable=TRUE
KeyLength=2048
KeySpec=1
KeyUsage=0xA0
MachineKeySet=True
ProviderName="Microsoft RSA SChannel Cryptographic Provider"
ProviderType=12
SMIME=FALSE
RequestType=PKCS10
[Strings]
szOID_ENHANCED_KEY_USAGE = "2.5.29.37"
szOID_PKIX_KP_SERVER_AUTH = "1.3.6.1.5.5.7.3.1"
szOID_PKIX_KP_CLIENT_AUTH = "1.3.6.1.5.5.7.3.2"
'@

    $iniContent = $iniContent -f $Subject

    Add-Content -Path $infFile -Value $iniContent
    Write-Verbose "ini file created '$infFile'"
 
    if ($SAN)
    {
        Write-Verbose 'Assing SAN section'
        Add-Content -Path $infFile -Value 'szOID_SUBJECT_ALT_NAME2 = "2.5.29.17"'
        Add-Content -Path $infFile -Value '[Extensions]'
        Add-Content -Path $infFile -Value '2.5.29.17 = "{text}"'
 
        foreach ($s in $SAN)
        {
            Write-Verbose "`t $s"
            $temp = '_continue_ = "dns={0}&"' -f $s
            Add-Content -Path $infFile -Value $temp
        }
    }
 
    ### Certificate request generation
    Remove-Item -Path $requestFile
    Write-Verbose "Calling 'certreq.exe -new $infFile $requestFile | Out-Null'"
    certreq.exe -new $infFile $requestFile | Out-Null
 
    ### Online certificate request and import
    if (-not $OnlineCA)
    {
        Write-Verbose 'No CA given, trying to find one...'
        $OnlineCA = Find-CertificateAuthority -ErrorAction Stop
        Write-Verbose "Found CA '$OnlineCA'"
    }
    
    if (-not $OnlineCA)
    {
        Write-Error "No OnlineCA given and no one could be found in the machine's domain"
        return
    }
       
    Remove-Item -Path $certFile
    Write-Verbose "Calling 'certreq.exe -q -submit -attrib CertificateTemplate:$TemplateName -config $OnlineCA $requestFile $certFile | Out-Null'"
    certreq.exe -submit -q -attrib "CertificateTemplate:$TemplateName" -config $OnlineCA $requestFile $certFile | Out-Null

    if ($LASTEXITCODE)
    {
        $ex = New-Object System.ComponentModel.Win32Exception($LASTEXITCODE)
        Write-Error -Message "Submitting the certificate request failed: $($ex.Message)" -Exception $ex 
        return
    }
 
    Write-Verbose "Calling 'certreq.exe -accept $certFile'"
    certreq.exe -q -accept $certFile
    if ($LASTEXITCODE)
    {
        $ex = New-Object System.ComponentModel.Win32Exception($LASTEXITCODE)
        Write-Error -Message "Accepting the certificate failed: $($ex.Message)" -Exception $ex
        return
    }

    Copy-Item -Path $certFile -Destination c:\cert.cer -Force
    Copy-Item -Path $infFile -Destination c:\request.inf -Force

    $certPrint = [System.Security.Cryptography.X509Certificates.X509Certificate2][System.Security.Cryptography.X509Certificates.X509Certificate2]::CreateFromCertFile('C:\cert.cer')
    $certPrint

    Remove-Item -Path $infFile, $requestFile, $certFile, $rspFile, 'C:\cert.cer' -Force
}

function Test-CATemplate
{
    [cmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$TemplateName
    )
    
    $tempates = certutil.exe -Template | Select-String -Pattern TemplatePropCommonName

    $template = $tempates -like "*$TemplateName"

    return [bool]$template
}

function Add-TfsAgentUserCapability
{
    param
    (
        [Parameter(Mandatory = $true)]
        [string]
        $InstanceName,

        [Parameter()]
        [string]
        $CollectionName = 'DefaultCollection',

        [Parameter(Mandatory = $true)]
        [string]
        $PoolName = '*',

        [Parameter(Mandatory = $true, ParameterSetName = 'CredId')]
        [Parameter(Mandatory = $true, ParameterSetName = 'PatId')]
        [uint16]
        $AgentId,

        [Parameter(Mandatory = $true, ParameterSetName = 'CredObject')]
        [Parameter(Mandatory = $true, ParameterSetName = 'PatObject')]
        [object]
        $Agent,

        [ValidateRange(1, 65535)]
        [uint32]
        $Port,

        [string]
        $ApiVersion = '5.1',

        [switch]
        $UseSsl,

        [Parameter(Mandatory = $true, ParameterSetName = 'CredId')]
        [Parameter(Mandatory = $true, ParameterSetName = 'CredObject')]
        [pscredential]
        $Credential,
        
        [Parameter(Mandatory = $true, ParameterSetName = 'PatId')]
        [Parameter(Mandatory = $true, ParameterSetName = 'PatObject')]
        [string]
        $PersonalAccessToken,

        [switch]
        $SkipCertificateCheck,

        [Parameter(Mandatory = $true)]
        [hashtable]
        $Capability
    )

    if ($SkipCertificateCheck.IsPresent)
    {
        $null = [ServerCertificateValidationCallback]::Ignore()
    }

    $poolParam = Sync-Parameter -Command (Get-Command Get-TfsAgentPool) -Parameter $PSBoundParameters
    $pool = Get-TfsAgentPool @poolParam

    if (-not $pool)
    {
        Write-Error -Message "Pool $PoolName could not be found!"
        return
    }

    if ($AgentId)
    {
        $agtParam = Sync-Parameter -Command (Get-Command Get-TfsAgent) -Parameter $PSBoundParameters
        $Agent = Get-TfsAgent @agtParam -Filter {$_.id -eq $AgentId}
    }

    if (-not $Agent)
    {
        Write-Error -Message "Agent could not be found!"
        return
    }

    $requestUrl = if ($UseSsl) {'https://' } else {'http://'}
    $requestUrl += if ( $Port  -gt 0)
    {
        '{0}{1}/{2}/_apis/distributedtask/pools/{3}/agents/{4}/usercapabilities' -f $InstanceName, ":$Port", $CollectionName, $pool.id, $Agent.Id
    }
    else
    {
        '{0}/{1}/_apis/distributedtask/pools/{2}/agents/{3}/usercapabilities' -f $InstanceName, $CollectionName, $pool.id, $Agent.Id
    }
    
    if ($ApiVersion)
    {
        $requestUrl += '?api-version={0}' -f $ApiVersion
    }

    $settableCapabilities = @{ }
    foreach ($prop in $Agent.usercapabilities.psobject.properties)
    {
        $settableCapabilities[$prop.Name] = $prop.Value
    }

    foreach ($kvp in $Capability.GetEnumerator())
    {
        $settableCapabilities[$kvp.Key] = $kvp.Value
    }

    $requestParameters = @{
        Uri             = $requestUrl
        Method          = 'Put'
        ContentType     = 'application/json'
        Body            = ($settableCapabilities | ConvertTo-Json)
        ErrorAction     = 'Stop'
        UseBasicParsing = $true
    }

    if ($PSEdition -eq 'Core' -and (Get-Command -Name Invoke-RestMethod).Parameters.ContainsKey('SkipCertificateCheck'))
    {
        $requestParameters.SkipCertificateCheck = $SkipCertificateCheck.IsPresent
    }

    if ($Credential)
    {
        $requestParameters.Credential = $Credential
    }
    else
    {
        $requestParameters.Headers = @{ Authorization = Get-TfsAccessTokenString -PersonalAccessToken $PersonalAccessToken }
    }

    try
    {
        $result = Invoke-RestMethod @requestParameters
    }
    catch
    {
        Write-Error -ErrorRecord $_
    }
    
    if ($result.value)
    {
        return $result.value
    }
    elseif ($result)
    {
        return $result
    }
}

function Get-TfsAccessTokenString
{
    [CmdletBinding()]
    [OutputType([String])]
    param
    (
        [Parameter(Mandatory = $True)]
        [String] $PersonalAccessToken
    )

    $tokenString = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f '',$PersonalAccessToken)))
    return ("Basic {0}" -f $tokenString)
}

function Get-TfsAgent
{
    param
    (
        [Parameter(Mandatory = $true)]
        [string]
        $InstanceName,

        [Parameter()]
        [string]
        $CollectionName = 'DefaultCollection',

        [Parameter(Mandatory = $true)]
        [string]
        $PoolName,

        [ValidateRange(1, 65535)]
        [uint32]
        $Port,

        [string]
        $ApiVersion,

        [switch]
        $UseSsl,

        [Parameter(Mandatory = $true, ParameterSetName = 'Cred')]
        [pscredential]
        $Credential,
        
        [Parameter(Mandatory = $true, ParameterSetName = 'Pat')]
        [string]
        $PersonalAccessToken,

        [switch]
        $SkipCertificateCheck,

        [scriptblock]
        $Filter = { $true }
    )

    if ($SkipCertificateCheck.IsPresent)
    {
        $null = [ServerCertificateValidationCallback]::Ignore()
    }

    $poolParam = Sync-Parameter -Command (Get-Command Get-TfsAgentPool) -Parameter $PSBoundParameters
    $pool = Get-TfsAgentPool @poolParam

    if (-not $pool)
    {
        Write-Error -Message "Pool $PoolName could not be found!"
        return
    }

    $requestUrl = if ($UseSsl) {'https://' } else {'http://'}
    $requestUrl += if ( $Port  -gt 0)
    {
        '{0}{1}/{2}/_apis/distributedtask/pools/{3}/agents?includeCapabilities=true' -f $InstanceName, ":$Port", $CollectionName, $pool.id
    }
    else
    {
        '{0}/{1}/_apis/distributedtask/pools/{2}/agents?includeCapabilities=true' -f $InstanceName, $CollectionName, $pool.id
    }
    
    if ($ApiVersion)
    {
        $requestUrl += '&api-version={0}' -f $ApiVersion
    }

    $requestParameters = @{
        Uri             = $requestUrl
        Method          = 'Get'
        ErrorAction     = 'Stop'
        UseBasicParsing = $true
    }

    if ($PSEdition -eq 'Core' -and (Get-Command -Name Invoke-RestMethod).Parameters.ContainsKey('SkipCertificateCheck'))
    {
        $requestParameters.SkipCertificateCheck = $SkipCertificateCheck.IsPresent
    }

    if ($Credential)
    {
        $requestParameters.Credential = $Credential
    }
    else
    {
        $requestParameters.Headers = @{ Authorization = Get-TfsAccessTokenString -PersonalAccessToken $PersonalAccessToken }
    }

    try
    {
        $result = Invoke-RestMethod @requestParameters
    }
    catch
    {
        Write-Error -ErrorRecord $_
    }
    
    if ($result.value)
    {
        return $result.value | Where-Object -FilterScript $Filter
    }
    elseif ($result)
    {
        return $result | Where-Object -FilterScript $Filter
    }
}

function Get-TfsAgentPool
{
    param
    (
        [Parameter(Mandatory = $true)]
        [string]
        $InstanceName,

        [Parameter()]
        [string]
        $CollectionName = 'DefaultCollection',

        [Parameter()]
        [string]
        $PoolName = '*',

        [ValidateRange(1, 65535)]
        [uint32]
        $Port,

        [string]
        $ApiVersion = '2.3-preview.1',

        [switch]
        $UseSsl,

        [Parameter(Mandatory = $true, ParameterSetName = 'Cred')]
        [pscredential]
        $Credential,
        
        [Parameter(Mandatory = $true, ParameterSetName = 'Pat')]
        [string]
        $PersonalAccessToken,

        [switch]
        $SkipCertificateCheck
    )

    if ($SkipCertificateCheck.IsPresent)
    {
        $null = [ServerCertificateValidationCallback]::Ignore()
    }

    $requestUrl = if ($UseSsl) {'https://' } else {'http://'}
    $requestUrl += if ( $Port  -gt 0)
    {
        '{0}{1}/{2}/_apis/distributedtask/pools' -f $InstanceName, ":$Port", $CollectionName
    }
    else
    {
        '{0}/{1}/_apis/distributedtask/pools' -f $InstanceName, $CollectionName
    }
    
    if ($ApiVersion)
    {
        $requestUrl += '?api-version={0}' -f $ApiVersion
    }

    $requestParameters = @{
        Uri             = $requestUrl
        Method          = 'Get'
        ErrorAction     = 'Stop'
        UseBasicParsing = $true
    }

    if ($PSEdition -eq 'Core' -and (Get-Command -Name Invoke-RestMethod).Parameters.ContainsKey('SkipCertificateCheck'))
    {
        $requestParameters.SkipCertificateCheck = $SkipCertificateCheck.IsPresent
    }

    if ($Credential)
    {
        $requestParameters.Credential = $Credential
    }
    else
    {
        $requestParameters.Headers = @{ Authorization = Get-TfsAccessTokenString -PersonalAccessToken $PersonalAccessToken }
    }

    try
    {
        $result = Invoke-RestMethod @requestParameters
    }
    catch
    {
        Write-Error -ErrorRecord $_
    }
    
    if ($result.value)
    {
        return $result.value | Where-Object -Property Name -like $PoolName
    }
    elseif ($result)
    {
        return $result | Where-Object -Property Name -like $PoolName
    }
}

function Get-TfsAgentQueue
{
    
    param
    (
        [Parameter(Mandatory = $true)]
        [string]
        $InstanceName,

        [Parameter()]
        [string]
        $CollectionName = 'DefaultCollection',

        [ValidateRange(1, 65535)]
        [uint32]
        $Port,

        [string]
        $ApiVersion = '3.0-preview',

        [Parameter(Mandatory = $true)]
        [string]
        $ProjectName,

        [string]
        $QueueName,

        [switch]
        $UseSsl,

        [Parameter(Mandatory = $true, ParameterSetName = 'Cred')]
        [pscredential]
        $Credential,
        
        [Parameter(Mandatory = $true, ParameterSetName = 'Pat')]
        [string]
        $PersonalAccessToken,

        [switch]
        $SkipCertificateCheck
    )

    if ($SkipCertificateCheck.IsPresent)
    {
        $null = [ServerCertificateValidationCallback]::Ignore()
    }

    $requestUrl = if ($UseSsl) {'https://' } else {'http://'}
    $requestUrl += if ( $Port -gt 0)
    {
        '{0}{1}/{2}/{3}/_apis/distributedtask/queues' -f $InstanceName, ":$Port", $CollectionName, $ProjectName
    }
    else
    {
        '{0}/{1}/{2}/_apis/distributedtask/queues' -f $InstanceName, $CollectionName, $ProjectName
    }
    
    if ($ApiVersion)
    {
        $requestUrl += '?api-version={0}' -f $ApiVersion
    }

    if ($QueueName)
    {
        $requestUrl += '&queueName={0}' -f $QueueName
    }

    $requestParameters = @{
        Uri             = $requestUrl
        Method          = 'Get'
        ErrorAction     = 'Stop'
        UseBasicParsing = $true
    }

    if ($Credential)
    {
        $requestParameters.Credential = $Credential
    }
    else
    {
        $requestParameters.Headers = @{ Authorization = Get-TfsAccessTokenString -PersonalAccessToken $PersonalAccessToken }
    }

    if ($PSEdition -eq 'Core' -and (Get-Command -Name Invoke-RestMethod).Parameters.ContainsKey('SkipCertificateCheck'))
    {
        $requestParameters.SkipCertificateCheck = $SkipCertificateCheck.IsPresent
    }

    try
    {
        $result = Invoke-RestMethod @requestParameters
    }
    catch
    {
        Write-Error -ErrorRecord $_
    }
    
    return $result.value
}

function Get-TfsBuildDefinition
{
    
    [CmdletBinding(DefaultParameterSetName = 'Cred')]
    param
    (
        [Parameter(Mandatory = $true)]
        [string]
        $InstanceName,

        [Parameter()]
        [string]
        $CollectionName = 'DefaultCollection',

        [ValidateRange(1, 65535)]
        [uint32]
        $Port,

        [string]
        $ApiVersion = '2.0',

        [Parameter(Mandatory = $true)]
        [string]
        $ProjectName,

        [string]
        $QueueName,

        [switch]
        $UseSsl,

        [Parameter(Mandatory = $true, ParameterSetName = 'Cred')]
        [pscredential]
        $Credential,
        
        [Parameter(Mandatory = $true, ParameterSetName = 'Pat')]
        [string]
        $PersonalAccessToken,

        [switch]
        $SkipCertificateCheck
    )

    if ($SkipCertificateCheck.IsPresent)
    {
        $null = [ServerCertificateValidationCallback]::Ignore()
    }
    
    $requestUrl = if ($UseSsl) {'https://' } else {'http://'}
    $requestUrl += if ( $Port -gt 0)
    {
        '{0}{1}/{2}/{3}/_apis/build/definitions' -f $InstanceName, ":$Port", $CollectionName, $ProjectName
    }
    else
    {
        '{0}/{1}/{2}/_apis/build/definitions' -f $InstanceName, $CollectionName, $ProjectName
    }
    
    if ($ApiVersion)
    {
        $requestUrl += '?api-version={0}' -f $ApiVersion
    }

    $requestParameters = @{
        Uri             = $requestUrl
        Method          = 'Get'
        ErrorAction     = 'Stop'
        UseBasicParsing = $true
    }

    if ($PSEdition -eq 'Core' -and (Get-Command -Name Invoke-RestMethod).Parameters.ContainsKey('SkipCertificateCheck'))
    {
        $requestParameters.SkipCertificateCheck = $SkipCertificateCheck.IsPresent
    }

    if ($Credential)
    {
        $requestParameters.Credential = $Credential
    }
    else
    {
        $requestParameters.Headers = @{ Authorization = Get-TfsAccessTokenString -PersonalAccessToken $PersonalAccessToken }
    }

    try
    {
        $result = Invoke-RestMethod @requestParameters
    }
    catch
    {
        if ($_.ErrorDetails.Message)
        {
            $errorDetails = $_.ErrorDetails.Message | ConvertFrom-Json
            if ($errorDetails.typeKey -eq 'ProjectDoesNotExistWithNameException')
            {
                return $null
            }
        }
        
        Write-Error -ErrorRecord $_
    }
    
    return $result.value
}

function Get-TfsBuildDefinitionTemplate
{
    
    [CmdletBinding(DefaultParameterSetName = 'Cred')]
    param
    (
        [Parameter(Mandatory = $true)]
        [string]
        $InstanceName,

        [Parameter()]
        [string]
        $CollectionName = 'DefaultCollection',

        [ValidateRange(1, 65535)]
        [uint32]
        $Port,

        [string]
        $ApiVersion = '2.0',

        [Parameter(Mandatory = $true)]
        [string]
        $ProjectName,

        [switch]
        $UseSsl,

        [Parameter(Mandatory = $true, ParameterSetName = 'Cred')]
        [pscredential]
        $Credential,
        
        [Parameter(Mandatory = $true, ParameterSetName = 'Pat')]
        [string]
        $PersonalAccessToken,

        [switch]
        $SkipCertificateCheck
    )

    if ($SkipCertificateCheck.IsPresent)
    {
        $null = [ServerCertificateValidationCallback]::Ignore()
    }

    $requestUrl = if ($UseSsl) {'https://' } else {'http://'}
    $requestUrl += if ( $Port  -gt 0)
    {
        '{0}{1}/{2}/{3}/_apis/build/definitions/templates' -f $InstanceName, ":$Port", $CollectionName, $ProjectName
    }
    else
    {
        '{0}/{1}/{2}/_apis/build/definitions/templates' -f $InstanceName, $CollectionName, $ProjectName
    }
    
    if ($ApiVersion)
    {
        $requestUrl += '?api-version={0}' -f $ApiVersion
    }

    $requestParameters = @{
        Uri             = $requestUrl
        Method          = 'Get'
        ErrorAction     = 'Stop'
        UseBasicParsing = $true
    }

    if ($PSEdition -eq 'Core' -and (Get-Command -Name Invoke-RestMethod).Parameters.ContainsKey('SkipCertificateCheck'))
    {
        $requestParameters.SkipCertificateCheck = $SkipCertificateCheck.IsPresent
    }

    if ($Credential)
    {
        $requestParameters.Credential = $Credential
    }
    else
    {
        $requestParameters.Headers = @{ Authorization = Get-TfsAccessTokenString -PersonalAccessToken $PersonalAccessToken }
    }

    try
    {
        $result = Invoke-RestMethod @requestParameters
    }
    catch
    {
        Write-Error -ErrorRecord $_
    }
    
    if ($result.value)
    {
        return $result.value
    }
    elseif ($result)
    {
        return $result
    }
}

function Get-TfsBuildStep
{
    [CmdletBinding(DefaultParameterSetName = 'Tfs')]
    param
    (
        [Parameter(Mandatory = $true)]
        [string]
        $InstanceName,

        [Parameter()]
        [string]
        $CollectionName = 'DefaultCollection',

        [ValidateRange(1, 65535)]
        [uint32]
        $Port,

        [Parameter(Mandatory = $true, ParameterSetName = 'TfsName')]
        [Parameter(Mandatory = $true, ParameterSetName = 'VstsName')]
        [SupportsWildcards()]
        [string]
        $FriendlyName,

        [Parameter(Mandatory = $true, ParameterSetName = 'TfsHashtable')]
        [Parameter(Mandatory = $true, ParameterSetName = 'VstsHashtable')]
        [hashtable]
        $FilterHashtable,

        [Parameter(Mandatory = $true, ParameterSetName = 'TfsScript')]
        [Parameter(Mandatory = $true, ParameterSetName = 'VstsScript')]
        [scriptblock]
        $FilterScript,

        [switch]
        $UseSsl,

        [Parameter(Mandatory = $true, ParameterSetName = 'Tfs')]
        [Parameter(Mandatory = $true, ParameterSetName = 'TfsName')]
        [Parameter(Mandatory = $true, ParameterSetName = 'TfsHashtable')]
        [Parameter(Mandatory = $true, ParameterSetName = 'TfsScript')]
        [pscredential]
        $Credential,
        
        [Parameter(Mandatory = $true, ParameterSetName = 'Vsts')]
        [Parameter(Mandatory = $true, ParameterSetName = 'VstsName')]
        [Parameter(Mandatory = $true, ParameterSetName = 'VstsHashtable')]
        [Parameter(Mandatory = $true, ParameterSetName = 'VstsScript')]
        [string]
        $PersonalAccessToken,

        [switch]
        $SkipCertificateCheck
    )

    if ($SkipCertificateCheck.IsPresent)
    {
        $null = [ServerCertificateValidationCallback]::Ignore()
    }

    $requestUrl = if ($UseSsl) {'https://' } else {'http://'}
    $requestUrl += if ( $Port -gt 0)
    {
        '{0}{1}/{2}/_apis/distributedtask/tasks' -f $InstanceName, ":$Port", $CollectionName
    }
    else
    {
        '{0}/{1}/_apis/distributedtask/tasks' -f $InstanceName, $CollectionName
    }

    $requestParameters = @{
        Uri         = $requestUrl
        Method      = 'Get'
        ErrorAction = 'Stop'
    }

    if ($PSEdition -eq 'Core' -and (Get-Command -Name Invoke-RestMethod).Parameters.ContainsKey('SkipCertificateCheck'))
    {
        $requestParameters.SkipCertificateCheck = $SkipCertificateCheck.IsPresent
    }

    if ($Credential)
    {
        $requestParameters.Credential = $Credential
    }
    else
    {
        $requestParameters.Headers = @{ Authorization = Get-TfsAccessTokenString -PersonalAccessToken $PersonalAccessToken }
    }

    try
    {
        $result = Invoke-RestMethod @requestParameters -UseBasicParsing
    }
    catch
    {
        Write-Error -ErrorRecord $_
    }
    
    $steps = if ($result.value)
    {
        $result.value | Where-Object -Property visibility -Contains 'Build'
    }
    elseif ($result -is [string])
    {
        ($result | ConvertFrom-Json -AsHashtable).value | Where-Object -Property visibility -Contains 'Build'
    }
    elseif ($result)
    {
        $result | Where-Object -Property visibility -Contains 'Build'
    }

    if ($FriendlyName)
    {
        $steps = if ($FriendlyName -match '(\?|\*)')
        {
            $steps | Where-Object -Property friendlyName -like $FriendlyName
        }
        else
        {
            $steps | Where-Object -Property friendlyName -eq $FriendlyName
        }
    }

    if ($FilterHashtable)
    {
        $steps = foreach ( $kvp in $FilterHashtable.GetEnumerator())
        {
            if ($kvp.Value -match '(\?|\*)')
            {
                $steps | Where-Object -Property $kvp.Key -like $kvp.Value
            }
            else
            {
                $steps | Where-Object -Property $kvp.Key -eq $kvp.Value    
            }            
        }
    }

    if ($FilterScript)
    {
        $steps = $steps | Where-Object -FilterScript $FilterScript
    }

    '@('
    foreach ($step in $steps)
    {
        "
        @{
            enabled         = $true
            continueOnError = $false
            alwaysRun       = $false
            displayName     = 'YOUR OWN DISPLAY NAME HERE' # e.g. $($step.instanceNameFormat) or $($step.friendlyName)
            task            = @{
                id          = '$($step.id)'
                versionSpec = '*'
            }
            inputs          = @{"
        foreach ($input in $step.inputs)
        {
            $required = if ($input.required) {$true}else {$false}
            "`t`t`t`t{0} = 'VALUE' # Type: {1}, Default: {2}, Mandatory: {3}" -f $input.name, $input.type, $input.defaultValue, $required
        }
        '
            }
        }
        '
    }
    ')'
}

function Get-TfsFeed
{
    
    param
    (
        [Parameter(Mandatory = $true)]
        [string]
        $InstanceName,
 
        [Parameter()]
        [string]
        $CollectionName = 'DefaultCollection',
 
        [ValidateRange(1, 65535)]
        [uint32]
        $Port,
 
        [string]
        $ApiVersion = '1.0',
 
        [string]
        $FeedName,
 
        [switch]
        $UseSsl,
 
        [Parameter(ParameterSetName = 'Tfs')]
        [pscredential]
        $Credential,
         
        [Parameter(ParameterSetName = 'Vsts')]
        [string]
        $PersonalAccessToken,

        [switch]
        $SkipCertificateCheck
    )

    if ($SkipCertificateCheck.IsPresent)
    {
        $null = [ServerCertificateValidationCallback]::Ignore()
    }
 
    $requestUrl = if ($UseSsl) {'https://' } else {'http://'}
    $requestUrl += if ($Port -gt 0)
    {
        '{0}{1}/{2}/_apis/packaging/feeds' -f $InstanceName, ":$Port", $CollectionName
    }
    else
    {
        '{0}/{1}/_apis/packaging/feeds' -f $InstanceName, $CollectionName
    }
    
    if ($ApiVersion)
    {
        $requestUrl += '?api-version={0}' -f $ApiVersion
    }
 
    $requestParameters = @{
        Uri         = $requestUrl
        Method      = 'Get'
        ErrorAction = 'Stop'
    }

    if ($PSEdition -eq 'Core' -and (Get-Command -Name Invoke-RestMethod).Parameters.ContainsKey('SkipCertificateCheck'))
    {
        $requestParameters.SkipCertificateCheck = $SkipCertificateCheck.IsPresent
    }
 
    if ($Credential)
    {
        $requestParameters.Credential = $Credential
    }
    else
    {
        $requestParameters.Headers = @{ Authorization = Get-TfsAccessTokenString -PersonalAccessToken $PersonalAccessToken }
    }
 
    try
    {
        $result = Invoke-RestMethod @requestParameters
    }
    catch
    {
        if ($_.ErrorDetails.Message)
        {
            $errorDetails = $_.ErrorDetails.Message | ConvertFrom-Json
            if ($errorDetails.typeKey -eq 'ProjectDoesNotExistWithNameException')
            {
                return $null
            }
        }
        Write-Error -ErrorRecord $_
    }
     
    $data = if ($result.value)
    {
        $result.value
    }
    elseif ($result)
    {
        $result
    }

    if ($FeedName)
    {
        $data | Where-Object name -eq $FeedName
    }
    else
    {
        $data
    }
}

function Get-TfsFeedPermission
{
    
    param
    (
        [Parameter(Mandatory = $true)]
        [string]
        $InstanceName,
 
        [Parameter()]
        [string]
        $CollectionName = 'DefaultCollection',
 
        [ValidateRange(1, 65535)]
        [uint32]
        $Port,
 
        [string]
        $ApiVersion = '1.0',
 
        [string]
        $FeedName,
 
        [switch]
        $UseSsl,
 
        [Parameter(ParameterSetName = 'Tfs')]
        [pscredential]
        $Credential,
         
        [Parameter(ParameterSetName = 'Vsts')]
        [string]
        $PersonalAccessToken,

        [switch]
        $SkipCertificateCheck
    )

    if ($SkipCertificateCheck.IsPresent)
    {
        $null = [ServerCertificateValidationCallback]::Ignore()
    }

    $feed = Get-TfsFeed @PSBoundParameters
    if (-not $feed)
    {
        Write-Warning "The feed '$FeedName' does not exist."
        return
    }
 
    $requestUrl = if ($UseSsl) {'https://' } else {'http://'}
    $requestUrl += if ($Port -gt 0)
    {
        '{0}{1}/{2}/_apis/packaging/feeds/{3}/permissions' -f $InstanceName, ":$Port", $CollectionName, $feed.id
    }
    else
    {
        '{0}/{1}/_apis/packaging/feeds/{2}/permissions' -f $InstanceName, $CollectionName, $feed.id
    }
    
    if ($ApiVersion)
    {
        $requestUrl += '?api-version={0}' -f $ApiVersion
    }
 
    $requestParameters = @{
        Uri         = $requestUrl
        Method      = 'Get'
        ErrorAction = 'Stop'
    }

    if ($PSEdition -eq 'Core' -and (Get-Command -Name Invoke-RestMethod).Parameters.ContainsKey('SkipCertificateCheck'))
    {
        $requestParameters.SkipCertificateCheck = $SkipCertificateCheck.IsPresent
    }
 
    if ($Credential)
    {
        $requestParameters.Credential = $Credential
    }
    else
    {
        $requestParameters.Headers = @{ Authorization = Get-TfsAccessTokenString -PersonalAccessToken $PersonalAccessToken }
    }
 
    try
    {
        $result = Invoke-RestMethod @requestParameters
    }
    catch
    {
        if ($_.ErrorDetails.Message)
        {
            $errorDetails = $_.ErrorDetails.Message | ConvertFrom-Json
            if ($errorDetails.typeKey -eq 'ProjectDoesNotExistWithNameException')
            {
                return $null
            }
        }
        Write-Error -ErrorRecord $_
    }
     
    if ($result.value)
    {
        $result.value
    }
    elseif ($result)
    {
        $result
    }
}

function Get-TfsGitRepository
{
    param
    (
        [Parameter(Mandatory = $true)]
        [string]
        $InstanceName,

        [Parameter(Mandatory = $true)]
        [string]
        $CollectionName,

        [ValidateRange(1, 65535)]
        [uint32]
        $Port,

        [string]
        $ApiVersion = '1.0',

        [string]
        $ProjectName,

        [switch]
        $UseSsl,

        [Parameter(ParameterSetName = 'Tfs')]
        [pscredential]
        $Credential,
        
        [Parameter(ParameterSetName = 'Vsts')]
        [string]
        $PersonalAccessToken,

        [switch]
        $SkipCertificateCheck
    )

    if ($SkipCertificateCheck.IsPresent)
    {
        $null = [ServerCertificateValidationCallback]::Ignore()
    }

    $requestUrl = if ($UseSsl) {'https://' } else {'http://'}
    $requestUrl += if ( $Port -gt 0)
    {
        '{0}{1}/{2}/{3}/_apis/git/repositories' -f $InstanceName, ":$Port", $CollectionName, $ProjectName, $ApiVersion
    }
    else
    {
        '{0}/{1}/{2}/_apis/git/repositories' -f $InstanceName, $CollectionName, $ProjectName, $ApiVersion
    }
    
    if ($ApiVersion)
    {
        $requestUrl += '?api-version={0}' -f $ApiVersion
    }

    $requestParameters = @{
        Uri         = $requestUrl
        Method      = 'Get'
        ErrorAction = 'Stop'
    }

    if ($PSEdition -eq 'Core' -and (Get-Command -Name Invoke-RestMethod).Parameters.ContainsKey('SkipCertificateCheck'))
    {
        $requestParameters.SkipCertificateCheck = $SkipCertificateCheck.IsPresent
    }

    if ($Credential)
    {
        $requestParameters.Credential = $Credential
    }
    else
    {
        $requestParameters.Headers = @{ Authorization = Get-TfsAccessTokenString -PersonalAccessToken $PersonalAccessToken }
    }

    try
    {
        $result = Invoke-RestMethod @requestParameters
    }
    catch
    {
        if ($_.ErrorDetails.Message)
        {
            $errorDetails = $_.ErrorDetails.Message | ConvertFrom-Json
            if ($errorDetails.typeKey -eq 'ProjectDoesNotExistWithNameException')
            {
                return $null
            }
        }
        
        Write-Error -ErrorRecord $_
    }
    
    return $result.value
}

function Get-TfsProcessTemplate
{
    
    param
    (
        [Parameter(Mandatory = $true)]
        [string]
        $InstanceName,

        [Parameter()]
        [string]
        $CollectionName = 'DefaultCollection',

        [ValidateRange(1, 65535)]
        [uint32]
        $Port,

        [string]
        $ApiVersion = '1.0',

        [switch]
        $UseSsl,

        [Parameter(Mandatory = $true, ParameterSetName = 'Tfs')]
        [pscredential]
        $Credential,
        
        [Parameter(Mandatory = $true, ParameterSetName = 'Vsts')]
        [string]
        $PersonalAccessToken,

        [switch]
        $SkipCertificateCheck
    )

    if ($SkipCertificateCheck.IsPresent)
    {
        $null = [ServerCertificateValidationCallback]::Ignore()
    }

    $requestUrl = if ($UseSsl) {'https://' } else {'http://'}
    $requestUrl += if ($Port -gt 0)
    {
        '{0}{1}/{2}/_apis/process/processes' -f $InstanceName, ":$Port", $CollectionName
    }
    else
    {
        '{0}/{1}/_apis/process/processes' -f $InstanceName, $CollectionName
    }
    
    if ($ApiVersion)
    {
        $requestUrl += '?api-version={0}' -f $ApiVersion
    }
    
    $requestParameters = @{
        Uri    = $requestUrl
        Method = 'Get'
    }

    if ($PSEdition -eq 'Core' -and (Get-Command -Name Invoke-RestMethod).Parameters.ContainsKey('SkipCertificateCheck'))
    {
        $requestParameters.SkipCertificateCheck = $SkipCertificateCheck.IsPresent
    }

    if ( $Credential)
    {
        $requestParameters.Credential = $Credential
    }
    else
    {
        $requestParameters.Headers = @{ Authorization = Get-TfsAccessTokenString -PersonalAccessToken $PersonalAccessToken }
    }

    try
    {
        $result = Invoke-RestMethod @requestParameters
    }
    catch
    {
        Write-Error -ErrorRecord $_
    }

    if ($result.value)
    {
        return $result.value
    }
    elseif ($result)
    {
        return $result
    }
}

function Get-TfsProject
{
    
    param
    (
        [Parameter(Mandatory = $true)]
        [string]
        $InstanceName,
 
        [Parameter()]
        [string]
        $CollectionName = 'DefaultCollection',
 
        [ValidateRange(1, 65535)]
        [uint32]
        $Port,
 
        [string]
        $ApiVersion = '1.0',
 
        [string]
        $ProjectName,
 
        [switch]
        $UseSsl,
 
        [Parameter(ParameterSetName = 'Tfs')]
        [pscredential]
        $Credential,
         
        [Parameter(ParameterSetName = 'Vsts')]
        [string]
        $PersonalAccessToken,

        [switch]
        $SkipCertificateCheck
    )

    if ($SkipCertificateCheck.IsPresent)
    {
        $null = [ServerCertificateValidationCallback]::Ignore()
    }
 
    $requestUrl = if ($UseSsl) {'https://' } else {'http://'}
    $requestUrl += if ( $Port -gt 0)
    {
        '{0}{1}/{2}/_apis/projects{3}' -f $InstanceName, ":$Port", $CollectionName, "/$ProjectName"
    }
    else
    {
        '{0}/{1}/_apis/projects{2}' -f $InstanceName, $CollectionName, "/$ProjectName"
    }
    
    if ($ApiVersion)
    {
        $requestUrl += '?api-version={0}' -f $ApiVersion
    }
 
    $requestParameters = @{
        Uri         = $requestUrl
        Method      = 'Get'
        ErrorAction = 'Stop'
    }

    if ($PSEdition -eq 'Core' -and (Get-Command -Name Invoke-RestMethod).Parameters.ContainsKey('SkipCertificateCheck'))
    {
        $requestParameters.SkipCertificateCheck = $SkipCertificateCheck.IsPresent
    }
 
    if ($Credential)
    {
        $requestParameters.Credential = $Credential
    }
    else
    {
        $requestParameters.Headers = @{ Authorization = Get-TfsAccessTokenString -PersonalAccessToken $PersonalAccessToken }
    }
 
    try
    {
        $result = Invoke-RestMethod @requestParameters
    }
    catch
    {
        if ($_.ErrorDetails.Message)
        {
            $errorDetails = $_.ErrorDetails.Message | ConvertFrom-Json
            if ($errorDetails.typeKey -eq 'ProjectDoesNotExistWithNameException')
            {
                return $null
            }
        }
        Write-Error -ErrorRecord $_
    }
     
    if ($result.value)
    {
        return $result.value
    }
    elseif ($result)
    {
        return $result
    }
}

function Get-TfsReleaseDefinition
{
    [CmdletBinding(DefaultParameterSetName = 'Cred')]
    param
    (
        [Parameter(Mandatory = $true)]
        [string]
        $InstanceName,

        [Parameter()]
        [string]
        $CollectionName = 'DefaultCollection',

        [ValidateRange(1, 65535)]
        [uint32]
        $Port,

        [string]
        $ApiVersion,

        [Parameter(Mandatory = $true)]
        [string]
        $ProjectName,

        [switch]
        $UseSsl,

        [Parameter(Mandatory = $true, ParameterSetName = 'Cred')]
        [pscredential]
        $Credential,
        
        [Parameter(Mandatory = $true, ParameterSetName = 'Pat')]
        [string]
        $PersonalAccessToken,

        [switch]
        $SkipCertificateCheck
    )

    if ($SkipCertificateCheck.IsPresent)
    {
        $null = [ServerCertificateValidationCallback]::Ignore()
    }

    $requestUrl = if ($UseSsl) {'https://' } else {'http://'}
    $requestUrl += if ( $Port -gt 0)
    {
        '{0}{1}/{2}/{3}/_apis/release/definitions' -f $InstanceName, ":$Port", $CollectionName, $ProjectName
    }
    else
    {
        '{0}/{1}/{2}/_apis/release/definitions' -f $InstanceName, $CollectionName, $ProjectName
    }

    if ($ApiVersion)
    {
        $requestUrl += '?api-version={0}' -f $ApiVersion
    }

    $requestParameters = @{
        Uri             = $requestUrl
        Method          = 'Get'
        ErrorAction     = 'Stop'
        UseBasicParsing = $true
    }

    if ($PSEdition -eq 'Core' -and (Get-Command -Name Invoke-RestMethod).Parameters.ContainsKey('SkipCertificateCheck'))
    {
        $requestParameters.SkipCertificateCheck = $SkipCertificateCheck.IsPresent
    }

    if ($Credential)
    {
        $requestParameters.Credential = $Credential
    }
    else
    {
        $requestParameters.Headers = @{ Authorization = Get-TfsAccessTokenString -PersonalAccessToken $PersonalAccessToken }
    }

    try
    {
        $result = Invoke-RestMethod @requestParameters
    }
    catch
    {
        if ($_.ErrorDetails.Message)
        {
            $errorDetails = $_.ErrorDetails.Message | ConvertFrom-Json
            if ($errorDetails.typeKey -eq 'ProjectDoesNotExistWithNameException')
            {
                return $null
            }
        }
        
        Write-Error -ErrorRecord $_
    }
    
    return $result.value
}

function Get-TfsReleaseStep
{
    [CmdletBinding(DefaultParameterSetName = 'Tfs')]
    param
    (
        [Parameter(Mandatory = $true)]
        [string]
        $InstanceName,

        [Parameter()]
        [string]
        $CollectionName = 'DefaultCollection',

        [ValidateRange(1, 65535)]
        [uint32]
        $Port,

        [Parameter(Mandatory = $true, ParameterSetName = 'TfsName')]
        [Parameter(Mandatory = $true, ParameterSetName = 'VstsName')]
        [SupportsWildcards()]
        [string]
        $FriendlyName,

        [Parameter(Mandatory = $true, ParameterSetName = 'TfsHashtable')]
        [Parameter(Mandatory = $true, ParameterSetName = 'VstsHashtable')]
        [hashtable]
        $FilterHashtable,

        [Parameter(Mandatory = $true, ParameterSetName = 'TfsScript')]
        [Parameter(Mandatory = $true, ParameterSetName = 'VstsScript')]
        [scriptblock]
        $FilterScript,

        [switch]
        $UseSsl,

        [Parameter(Mandatory = $true, ParameterSetName = 'Tfs')]
        [Parameter(Mandatory = $true, ParameterSetName = 'TfsName')]
        [Parameter(Mandatory = $true, ParameterSetName = 'TfsHashtable')]
        [Parameter(Mandatory = $true, ParameterSetName = 'TfsScript')]
        [pscredential]
        $Credential,
        
        [Parameter(Mandatory = $true, ParameterSetName = 'Vsts')]
        [Parameter(Mandatory = $true, ParameterSetName = 'VstsName')]
        [Parameter(Mandatory = $true, ParameterSetName = 'VstsHashtable')]
        [Parameter(Mandatory = $true, ParameterSetName = 'VstsScript')]
        [string]
        $PersonalAccessToken,

        [switch]
        $SkipCertificateCheck
    )

    if ($SkipCertificateCheck.IsPresent)
    {
        $null = [ServerCertificateValidationCallback]::Ignore()
    }

    $requestUrl = if ($UseSsl) {'https://' } else {'http://'}
    $requestUrl += if ( $Port -gt 0)
    {
        '{0}{1}/{2}/_apis/distributedtask/tasks' -f $InstanceName, ":$Port", $CollectionName
    }
    else
    {
        '{0}/{1}/_apis/distributedtask/tasks' -f $InstanceName, $CollectionName
    }

    $requestParameters = @{
        Uri         = $requestUrl
        Method      = 'Get'
        ErrorAction = 'Stop'
    }

    if ($PSEdition -eq 'Core' -and (Get-Command -Name Invoke-RestMethod).Parameters.ContainsKey('SkipCertificateCheck'))
    {
        $requestParameters.SkipCertificateCheck = $SkipCertificateCheck.IsPresent
    }

    if ($Credential)
    {
        $requestParameters.Credential = $Credential
    }
    else
    {
        $requestParameters.Headers = @{ Authorization = Get-TfsAccessTokenString -PersonalAccessToken $PersonalAccessToken }
    }

    try
    {
        $result = Invoke-RestMethod @requestParameters
    }
    catch
    {
        Write-Error -ErrorRecord $_
    }
    
    $steps = if ($result.value)
    {
        $result.value | Where-Object -Property visibility -Contains 'Release'
    }
    elseif ($result -is [string])
    {
        ($result | ConvertFrom-Json -AsHashtable).value | Where-Object -Property visibility -Contains 'Release'
    }
    elseif ($result)
    {
        $result | Where-Object -Property visibility -Contains 'Release'
    }

    if ($FriendlyName)
    {
        $steps = if ($FriendlyName -match '(\?|\*)')
        {
            $steps | Where-Object -Property friendlyName -like $FriendlyName
        }
        else
        {
            $steps | Where-Object -Property friendlyName -eq $FriendlyName
        }
    }

    if ($FilterHashtable)
    {
        $steps = foreach ( $kvp in $FilterHashtable.GetEnumerator())
        {
            if ($kvp.Value -match '(\?|\*)')
            {
                $steps | Where-Object -Property $kvp.Key -like $kvp.Value
            }
            else
            {
                $steps | Where-Object -Property $kvp.Key -eq $kvp.Value    
            }            
        }
    }

    if ($FilterScript)
    {
        $steps = $steps | Where-Object -FilterScript $FilterScript
    }

    '@('
    foreach ($step in $steps)
    {
        "
        @{
            enabled          = `$true
            continueOnError  = `$false
            alwaysRun        = `$false
            timeoutInMinutes = 0
            definitionType   = 'task'
            version          = '*'
            name             = 'YOUR OWN DISPLAY NAME HERE' # e.g. $($step.instanceNameFormat) or $($step.friendlyName)
            taskid           = '$($step.id)'
            inputs           = @{"
        foreach ($input in $step.inputs)
        {
            $required = if ($input.required) {$true}else {$false}
            "`t`t`t`t{0} = 'VALUE' # Type: {1}, Default: {2}, Mandatory: {3}" -f $input.name, $input.type, $input.defaultValue, $required
        }
        '
            }
        }
        '
    }
    ')'
}

function New-TfsAgentQueue
{
    
    param
    (
        [Parameter(Mandatory = $true)]
        [string]
        $InstanceName,

        [Parameter()]
        [string]
        $CollectionName = 'DefaultCollection',

        [ValidateRange(1, 65535)]
        [uint32]
        $Port,

        [string]
        $ApiVersion = '3.0-preview.1',

        [Parameter(Mandatory = $true)]
        [string]
        $ProjectName,

        [switch]
        $UseSsl,

        [string]
        $QueueName,

        [Parameter(Mandatory = $true, ParameterSetName = 'Cred')]
        [pscredential]
        $Credential,
        
        [Parameter(Mandatory = $true, ParameterSetName = 'Pat')]
        [string]
        $PersonalAccessToken,

        [switch]
        $SkipCertificateCheck
    )

    if ($SkipCertificateCheck.IsPresent)
    {
        $null = [ServerCertificateValidationCallback]::Ignore()
    }

    $existingQueue = Get-TfsAgentQueue @PSBoundParameters
    if ($existingQueue) { return $existingQueue }
    
    $requestUrl = if ($UseSsl) {'https://' } else {'http://'}
    $requestUrl += if ( $Port -gt 0)
    {
        '{0}{1}/{2}/{3}/_apis/distributedTask/queues' -f $InstanceName, ":$Port", $CollectionName, $ProjectName
    }
    else
    {
        '{0}/{1}/{2}/_apis/distributedTask/queues' -f $InstanceName, $CollectionName, $ProjectName
    }
    
    if ($ApiVersion)
    {
        $requestUrl += '?api-version={0}' -f $ApiVersion
    }

    $poolParameter = Sync-Parameter -Command (Get-Command Get-TfsAgentPool) -Parameters $PSBoundParameters
    $pools = Get-TfsAgentPool @poolParameter

    $useablePool = $pools | Where-Object -Property size -gt 0 | Select-Object -First 1
    if (-not $useablePool) { $useablePool = $pools | Select-Object -First 1}
    if (-not $useablePool) { Write-Error -Message 'No agent pools available to form queue'; return}

    $payload = [ordered]@{
        "name" = $QueueName
        "pool" = @{
            "id" = $useablePool.id
        }
    }

    $requestParameters = @{
        Uri         = $requestUrl
        Method      = 'Post'
        ContentType = 'application/json'
        Body        = ($payload | ConvertTo-Json)
        ErrorAction = 'Stop'
    }

    if ($PSEdition -eq 'Core' -and (Get-Command -Name Invoke-RestMethod).Parameters.ContainsKey('SkipCertificateCheck'))
    {
        $requestParameters.SkipCertificateCheck = $SkipCertificateCheck.IsPresent
    }

    if ($Credential)
    {
        $requestParameters.Credential = $Credential
    }
    else
    {
        $requestParameters.Headers = @{ Authorization = Get-TfsAccessTokenString -PersonalAccessToken $PersonalAccessToken }
    }

    try
    {
        $result = Invoke-RestMethod @requestParameters
    }
    catch
    {
        Write-Error -ErrorRecord $_
    }
}

function New-TfsBuildDefinition
{
    
    [CmdletBinding(DefaultParameterSetName = 'Cred')]
    param
    (
        [Parameter(Mandatory = $true)]
        [string]
        $InstanceName,

        [Parameter()]
        [string]
        $CollectionName = 'DefaultCollection',

        [ValidateRange(1, 65535)]
        [uint32]
        $Port,

        [string]
        $ApiVersion = '2.0',

        [Parameter(Mandatory = $true)]
        [string]
        $ProjectName,

        [Parameter(Mandatory = $true)]
        [string]
        $DefinitionName,

        [string]
        $QueueName,

        [hashtable[]]
        $BuildTasks, # Not very nice and needs to be replaced as soon as I find out how to retrieve all build step guids

        [hashtable[]]
        $Phases,

        [string[]]
        $CiTriggerRefs,

        [hashtable]
        $Variables,

        [switch]
        $UseSsl,

        [Parameter(Mandatory = $true, ParameterSetName = 'Cred')]
        [pscredential]
        $Credential,
        
        [Parameter(Mandatory = $true, ParameterSetName = 'Pat')]
        [string]
        $PersonalAccessToken,
        
        [switch]$Clean,
        
        [int]$CleanOptions = 0,

        [switch]
        $SkipCertificateCheck
    )

    if ($SkipCertificateCheck.IsPresent)
    {
        $null = [ServerCertificateValidationCallback]::Ignore()
    }

    $requestUrl = if ($UseSsl) {'https://' } else {'http://'}
    $requestUrl += if ($Port -gt 0)
    {
        '{0}{1}/{2}/{3}/_apis/build/definitions' -f $InstanceName, ":$Port", $CollectionName, $ProjectName
    }
    else
    {
        '{0}/{1}/{2}/_apis/build/definitions' -f $InstanceName, $CollectionName, $ProjectName
    }
    
    if ($ApiVersion)
    {
        $requestUrl += '?api-version={0}' -f $ApiVersion
    }

    $exBuildParam = Sync-Parameter -Command (Get-Command Get-TfsBuildDefinition) -Parameters $PSBoundParameters
    $exBuildParam.Remove('Version')
    $existingBuild = Get-TfsBuildDefinition @exBuildParam
    if ($existingBuild | Where-Object name -eq $DefinitionName)
    { 
        Write-Warning -Message ('Build definition {0} in {1} already exists.' -f $DefinitionName, $ProjectName)
        return 
    }

    $qparameters = Sync-Parameter -Command (Get-Command Get-TfsAgentQueue) -Parameters $PSBoundParameters
    $qparameters.Remove('ApiVersion') # preview-API is called
    $qparameters.ErrorAction = 'SilentlyContinue'
    $queue = Get-TfsAgentQueue @qparameters | Select-Object -First 1

    if (-not $queue)
    {
        Write-Verbose -Message ('No existing queue found for project {0}. Creating new queue.' -f $ProjectName)
        $parameters = Sync-Parameter -Command (Get-Command New-TfsAgentQueue) -Parameters $PSBoundParameters
        $parameters.Remove('ApiVersion') # preview-API is called
        $parameters.ErrorAction = 'Stop'
        $qparameters.ErrorAction = 'Stop'
        try
        {
            New-TfsAgentQueue @parameters
            $queue = Get-TfsAgentQueue @qparameters | Select-Object -First 1
        }
        catch
        {
            Write-Error -ErrorRecord $_
        }
    }

    $projectParameters = Sync-Parameter -Command (Get-Command Get-TfsProject) -Parameters $PSBoundParameters
    $projectParameters.ErrorAction = 'Stop'
    
    try
    {
        $project = Get-TfsProject @projectParameters
    }
    catch
    {
        Write-Error -ErrorRecord $_
    }

    $repoParameters = Sync-Parameter -Command (Get-Command Get-TfsGitRepository) -Parameters $PSBoundParameters
    $repoParameters.ErrorAction = 'Stop'

    try
    {
        $repo = Get-TfsGitRepository @repoParameters
    }
    catch
    {
        Write-Error -ErrorRecord $_
    }

    $buildDefinition = if ($ApiVersion -gt '4.0')
    {
        @{
            name       = $DefinitionName
            type       = "build"
            quality    = "definition"
            queue      = @{
                id = $queue.id
            }
            process      = @{ }
            repository = @{
                id            = $repo.id
                type          = "TfsGit"
                name          = $repo.name
                defaultBranch = "refs/heads/master"
                url           = $repo.remoteUrl                
                clean         = $Clean.ToBool()
                properties = @{
                    cleanOptions = "$CleanOptions"
                }
            }
            options    = @(
                @{
                    enabled    = $true
                    definition = @{
                        id = (New-Guid).Guid
                    }
                    inputs     = @{
                        parallel  = $false
                        multipliers = '["config","platform"]'
                    }
                }
            )
            variables  = @{
                forceClean = @{
                    value         = $false
                    allowOverride = $true
                }
                config     = @{
                    value         = "debug, release"
                    allowOverride = $true
                }
                platform   = @{
                    value         = "any cpu"
                    allowOverride = $true
                }
            }
        }
    }
    else
    {
        @{
            name       = $DefinitionName
            type       = "build"
            quality    = "definition"
            queue      = @{
                id = $queue.id
            }
            build      = $BuildTasks
            repository = @{
                id            = $repo.id
                type          = "TfsGit"
                name          = $repo.name
                defaultBranch = "refs/heads/master"
                url           = $repo.remoteUrl
                clean         = $false
            }
            options    = @(
                @{
                    enabled    = $true
                    definition = @{
                        id = (New-Guid).Guid
                    }
                    inputs     = @{
                        parallel  = $false
                        multipliers = '["config","platform"]'
                    }
                }
            )
            variables  = @{
                forceClean = @{
                    value         = $false
                    allowOverride = $true
                }
                config     = @{
                    value         = "debug, release"
                    allowOverride = $true
                }
                platform   = @{
                    value         = "any cpu"
                    allowOverride = $true
                }
            }
        }
    }

    if (-not $Phases -and $ApiVersion -ge '4.0')
    {
        $Phases =  @(
            @{
                name = 'Phase 1'
                condition = 'succeeded()'
            }
        )
        $buildDefinition.process.Add('phases', $Phases)

        if ($BuildTasks)
        {
            $buildDefinition.process.phases[0].Add('steps', $BuildTasks)
        }
    }

    $refs = @()
    if ($CiTriggerRefs)
    {
        foreach ($ref in $CiTriggerRefs)
        {
            $refs += "+$ref"
        }
        $trigger = @{
            branchFilters = $refs
            maxConcurrentBuildsPerBranch = 1
            pollingInterval = 0
            triggerType = 2
        }

        $buildDefinition.triggers = @($trigger)
    }

    if ($Variables)
    {
        foreach ($variable in $Variables.GetEnumerator())
        {
            $variableContent = @{
                value = $variable.Value
                allowOverrise = $true
            }
            $buildDefinition.variables.Add($variable.Key, $variableContent)
        }
    }

    $requestParameters = @{
        Uri         = $requestUrl
        Method      = 'Post'
        ContentType = 'application/json'
        Body        = ($buildDefinition | ConvertTo-Json -Depth 42)
        ErrorAction = 'Stop'
    }

    if ($PSEdition -eq 'Core' -and (Get-Command -Name Invoke-RestMethod).Parameters.ContainsKey('SkipCertificateCheck'))
    {
        $requestParameters.SkipCertificateCheck = $SkipCertificateCheck.IsPresent
    }

    if ($Credential)
    {
        $requestParameters.Credential = $Credential
    }
    else
    {
        $requestParameters.Headers = @{ Authorization = Get-TfsAccessTokenString -PersonalAccessToken $PersonalAccessToken }
    }

    try
    {
        $result = Invoke-RestMethod @requestParameters
        Write-Verbose -Message ('New build definition {0} created for project {1}' -f $DefinitionName, $ProjectName)
    }
    catch
    {
        Write-Error -ErrorRecord $_
    }
}

function New-TfsFeed
{
    
    [CmdletBinding(DefaultParameterSetName = 'NameCred')]
    param
    (
        [Parameter(Mandatory = $true)]
        [string]
        $InstanceName,

        [Parameter()]
        [string]
        $CollectionName = 'DefaultCollection',

        [ValidateRange(1, 65535)]
        [uint32]
        $Port,

        [string]
        $ApiVersion = '2.0',

        [Parameter(Mandatory = $true)]
        [string]
        $FeedName,

        [string]
        $Description,

        [switch]
        $UseSsl,

        [Parameter(Mandatory = $true, ParameterSetName = 'GuidCred')]
        [Parameter(Mandatory = $true, ParameterSetName = 'NameCred')]
        [pscredential]
        $Credential,
        
        [Parameter(Mandatory = $true, ParameterSetName = 'NamePat')]
        [Parameter(Mandatory = $true, ParameterSetName = 'GuidPat')]
        [string]
        $PersonalAccessToken,

        [timespan]
        $Timeout = (New-TimeSpan -Seconds 30),

        [switch]
        $SkipCertificateCheck
    )

    if ($SkipCertificateCheck.IsPresent)
    {
        $null = [ServerCertificateValidationCallback]::Ignore()
    }

    $requestUrl = if ($UseSsl) {'https://' } else {'http://'}
    $requestUrl += if ( $Port -gt 0)
    {
        '{0}{1}/{2}/_apis/packaging/feeds' -f $InstanceName, ":$Port", $CollectionName
    }
    else
    {
        '{0}/{1}/_apis/packaging/feeds' -f $InstanceName, $CollectionName
    }
    
    if ($ApiVersion)
    {
        $requestUrl += '?api-version={0}' -f $ApiVersion
    }

    $feedParameters = Sync-Parameter -Command (Get-Command -Name Get-TfsFeed) -Parameters $PSBoundParameters
    $feedParameters.ErrorAction = 'SilentlyContinue'
    if (Get-TfsFeed @feedParameters)
    {
        Write-Error -Message "The Feed '$FeedName' already exists"
        return
    }

    $payload = @{
        name         = $FeedName
        description  = $Description
    }

    $requestParameters = @{
        Uri         = $requestUrl
        Method      = 'Post'
        ContentType = 'application/json'
        Body        = ($payload | ConvertTo-Json)
        ErrorAction = 'Stop'
    }

    if ($PSEdition -eq 'Core' -and (Get-Command -Name Invoke-RestMethod).Parameters.ContainsKey('SkipCertificateCheck'))
    {
        $requestParameters.SkipCertificateCheck = $SkipCertificateCheck.IsPresent
    }

    if ($Credential)
    {
        $requestParameters.Credential = $Credential
    }
    else
    {
        $requestParameters.Headers = @{ Authorization = Get-TfsAccessTokenString -PersonalAccessToken $PersonalAccessToken }
    }

    try
    {
        $result = Invoke-RestMethod @requestParameters
    }
    catch
    {
        Write-Error -ErrorRecord $_
    }
}

function New-TfsProject
{
    
    [CmdletBinding(DefaultParameterSetName = 'NameCred')]
    param
    (
        [Parameter(Mandatory = $true)]
        [string]
        $InstanceName,

        [Parameter()]
        [string]
        $CollectionName = 'DefaultCollection',

        [ValidateRange(1, 65535)]
        [uint32]
        $Port,

        [string]
        $ApiVersion = '2.0',

        [Parameter(Mandatory = $true)]
        [string]
        $ProjectName,

        [string]
        $ProjectDescription,

        [ValidateSet('Git', 'Tfvc')]
        $SourceControlType = 'Git',

        [Parameter(Mandatory = $true, ParameterSetName = 'GuidPat')]
        [Parameter(Mandatory = $true, ParameterSetName = 'GuidCred')]
        [guid]
        $TemplateGuid,

        [Parameter(Mandatory = $true, ParameterSetName = 'NamePat')]
        [Parameter(Mandatory = $true, ParameterSetName = 'NameCred')]
        [string]
        $TemplateName,

        [switch]
        $UseSsl,

        [Parameter(Mandatory = $true, ParameterSetName = 'GuidCred')]
        [Parameter(Mandatory = $true, ParameterSetName = 'NameCred')]
        [pscredential]
        $Credential,
        
        [Parameter(Mandatory = $true, ParameterSetName = 'NamePat')]
        [Parameter(Mandatory = $true, ParameterSetName = 'GuidPat')]
        [string]
        $PersonalAccessToken,

        [timespan]
        $Timeout = (New-TimeSpan -Seconds 30),

        [switch]
        $SkipCertificateCheck
    )

    if ($SkipCertificateCheck.IsPresent)
    {
        $null = [ServerCertificateValidationCallback]::Ignore()
    }

    $requestUrl = if ($UseSsl) {'https://' } else {'http://'}
    $requestUrl += if ( $Port -gt 0)
    {
        '{0}{1}/{2}/_apis/projects' -f $InstanceName, ":$Port", $CollectionName
    }
    else
    {
        '{0}/{1}/_apis/projects' -f $InstanceName, $CollectionName
    }
    
    if ($ApiVersion)
    {
        $requestUrl += '?api-version={0}' -f $ApiVersion
    }

    if ($PSCmdlet.ParameterSetName -like 'Name*')
    {
        $parameters = Sync-Parameter -Command (Get-Command Get-TfsProcessTemplate) -Parameters $PSBoundParameters
        $TemplateGuid = (Get-TfsProcessTemplate @parameters | Where-Object -Property name -eq $TemplateName).id
        if (-not $TemplateGuid) {Write-Error -Message "Could not locate $TemplateName. Try Get-TfsProcessTemplate to see all available templates"; return}
    }

    $projectParameters = Sync-Parameter -Command (Get-Command Get-TfsProject) -Parameters $PSBoundParameters
    $projectParameters.ErrorAction = 'SilentlyContinue'
    if (Get-TfsProject @projectParameters)
    {
        Write-Verbose -Message ('Project {0} already exists' -f $ProjectName)
        return
    }

    $payload = @{
        name         = $ProjectName
        description  = $ProjectDescription
        capabilities = @{
            versioncontrol  = @{
                sourceControlType = $SourceControlType                
            }
            processTemplate = @{
                templateTypeId = $TemplateGuid.Guid
            }
        }
    }

    $requestParameters = @{
        Uri         = $requestUrl
        Method      = 'Post'
        ContentType = 'application/json'
        Body        = ($payload | ConvertTo-Json)
        ErrorAction = 'Stop'
    }

    if ($PSEdition -eq 'Core' -and (Get-Command -Name Invoke-RestMethod).Parameters.ContainsKey('SkipCertificateCheck'))
    {
        $requestParameters.SkipCertificateCheck = $SkipCertificateCheck.IsPresent
    }

    if ($Credential)
    {
        $requestParameters.Credential = $Credential
    }
    else
    {
        $requestParameters.Headers = @{ Authorization = Get-TfsAccessTokenString -PersonalAccessToken $PersonalAccessToken }
    }

    try
    {
        $result = Invoke-RestMethod @requestParameters
    }
    catch
    {
        Write-Error -ErrorRecord $_
    }

    $start = Get-Date
    $projectStatus = Get-TfsProject @projectParameters

    while ($projectStatus.State -ne 'wellFormed' -and ((Get-Date) - $start -lt $Timeout))
    {
        Write-Verbose -Message ('Waiting {0} for {1} to enter status wellFormed' -f $Timeout, $ProjectName)
        Start-Sleep -Seconds 1
        $projectStatus = Get-TfsProject @projectParameters
    }

    if (-not $projectStatus.State -eq 'wellFormed')
    {
        Write-Error -Message ('Unable to create new project in {0}' -f $Timeout) -TargetObject $ProjectName
        return
    }

    return $projectStatus
}

function New-TfsReleaseDefinition
{
    [CmdletBinding(DefaultParameterSetName = 'Cred')]
    param
    (
        [Parameter(Mandatory = $true)]
        [string]
        $InstanceName,

        [Parameter()]
        [string]
        $CollectionName = 'DefaultCollection',

        [ValidateRange(1, 65535)]
        [uint32]
        $Port,

        [string]
        $ApiVersion = '3.0-preview.3',

        [Parameter(Mandatory = $true)]
        [string]
        $ProjectName,

        [Parameter(Mandatory = $true)]
        [string]
        $ReleaseName,

        [hashtable[]]
        $ReleaseTasks,

        [hashtable[]]
        $Environments,

        [switch]
        $UseSsl,

        [Parameter(Mandatory = $true, ParameterSetName = 'Cred')]
        [pscredential]
        $Credential,
        
        [Parameter(Mandatory = $true, ParameterSetName = 'Pat')]
        [string]
        $PersonalAccessToken,

        [switch]
        $SkipCertificateCheck
    )

    if ($SkipCertificateCheck.IsPresent)
    {
        $null = [ServerCertificateValidationCallback]::Ignore()
    }

    $requestUrl = if ($UseSsl) {'https://' } else {'http://'}
    $requestUrl += if ( $Port -gt 0)
    {
        '{0}{1}/{2}/{3}/_apis/release/definitions' -f $InstanceName, ":$Port", $CollectionName, $ProjectName
    }
    else
    {
        '{0}/{1}/{2}/_apis/release/definitions' -f $InstanceName, $CollectionName, $ProjectName
    }

    if ($ApiVersion)
    {
        $requestUrl += '?api-version={0}' -f $ApiVersion
    }

    $exReleaseParam = Sync-Parameter -Command (Get-Command Get-TfsReleaseDefinition) -Parameters $PSBoundParameters
    $exReleaseParam.Remove('ApiVersion')
    $existingRelease = Get-TfsReleaseDefinition @exReleaseParam
    if ($existingRelease | Where-Object name -eq $ReleaseName)
    {
        Write-Verbose -Message ('Release definition {0} in {1} already exists.' -f $ReleaseName, $ProjectName);
        return 
    }

    $qparameters = Sync-Parameter -Command (Get-Command Get-TfsAgentQueue) -Parameters $PSBoundParameters
    $qparameters.Remove('ApiVersion') # preview-API is called
    $qparameters.ErrorAction = 'SilentlyContinue'
    $queue = Get-TfsAgentQueue @qparameters | Select-Object -First 1

    if (-not $queue)
    {
        $parameters = Sync-Parameter -Command (Get-Command New-TfsAgentQueue) -Parameters $PSBoundParameters
        $parameters.Remove('ApiVersion') # preview-API is called
        $parameters.ErrorAction = 'Stop'
        $qparameters.ErrorAction = 'Stop'
        try
        {
            New-TfsAgentQueue @parameters
            $queue = Get-TfsAgentQueue @qparameters | Select-Object -First 1
        }
        catch
        {
            Write-Error -ErrorRecord $_
        }
    }

    $projectParameters = Sync-Parameter -Command (Get-Command Get-TfsProject) -Parameters $PSBoundParameters
    $projectParameters.ErrorAction = 'Stop'
    
    try
    {
        $project = Get-TfsProject @projectParameters
    }
    catch
    {
        Write-Error -ErrorRecord $_
    }

    $buildParameters = Sync-Parameter -Command (Get-Command Get-TfsBuildDefinition) -Parameters $PSBoundParameters
    $buildParameters.ErrorAction = 'Stop'

    try
    {
        $build = Get-TfsBuildDefinition @buildParameters
    }
    catch
    {
        Write-Error -ErrorRecord $_
    }

    if (-not $Environments)
    {
        $Environments = @(
            @{
                "id"                      = 1 
                "name"                    = "Environment" 
                "rank"                    = 1 
                "deployStep"              = @{
                    "id"    = 0 
                    "tasks" = @()
                } 
                "deployPhases"            = @(
                    @{
                        "name"            = "Run on agent" 
                        "phaseType"       = 1 
                        "rank"            = 1 
                        "workflowTasks"   = $ReleaseTasks
                        "deploymentInput" = @{
                            "demands"               = @() 
                            "queueId"               = $queue.id 
                            "enableAccessToken"     = $false 
                            "skipArtifactsDownload" = $false 
                            "timeoutInMinutes"      = 0
                        } 
                        "controlOptions"  = @{
                            "alwaysRun"       = $false 
                            "continueOnError" = $false 
                            "enabled"         = $true
                        }
                    }
                ) 
                "queueId"                 = $queue.id 
                "demands"                 = @() 
                "conditions"              = @(
                    @{
                        "name"          = "ReleaseStarted" 
                        "conditionType" = 1 
                        "value"         = ""
                    }
                ) 
                "environmentOptions"      = @{
                    "emailNotificationType" = "OnlyOnFailure" 
                    "emailRecipients"       = "release.environment.owner;release.creator" 
                    "skipArtifactsDownload" = $false 
                    "timeoutInMinutes"      = 0 
                    "enableAccessToken"     = $false
                } 
                "executionPolicy"         = @{
                    "concurrencyCount" = 0 
                    "queueDepthCount"  = 0
                } 
                "releaseId"               = $null 
                "definitionEnvironmentId" = $null 
                "preDeployApprovals"      = @{
                    "approvals"       = @(
                        @{
                            "rank"             = 1 
                            "isAutomated"      = $true
                            "isNotificationOn" = $false 
                            "id"               = 0
                        }
                    ) 
                    "approvalOptions" = $null
                } 
                "postDeployApprovals"     = @{
                    "approvals"       = @(
                        @{
                            "rank"             = 1 
                            "isAutomated"      = $true 
                            "isNotificationOn" = $false 
                            "id"               = 0
                        }
                    ) 
                    "approvalOptions" = $null
                } 
                "schedules"               = @() 
                "retentionPolicy"         = @{
                    "daysToKeep"     = 30 
                    "releasesToKeep" = 3 
                    "retainBuild"    = $true
                }
            }
        )
    }

    $payload = @{
        "id"                = 0 
        "name"              = $ReleaseName
        "comment"           = $null 
        "createdOn"         = (Get-Date).ToString('yyyy-MM-ddThh:mm:ss.fffZ')
        "createdBy"         = $null 
        "modifiedBy"        = $null 
        "modifiedOn"        = $null 
        "environments"      = $Environments
        "artifacts"         = @(
            @{
                "id"                  = 0 
                "definitionReference" = @{
                    "project"    = @{
                        "id"   = $project.id
                        "name" = $project.name
                    } 
                    "definition" = @{
                        "id"   = $build.id
                        "name" = $build.name
                    }
                } 
                "alias"               = $build.name
                "type"                = "Build" 
                "artifactTypeName"    = "Build" 
                "sourceId"            = "" 
                "isPrimary"           = $true
            }
        )  
        "triggers"          = @(
            @{
                "triggerType"   = 1 
                "artifactAlias" = $build.name
            }
        ) 
        "releaseNameFormat" = 'Release-$(rev:r)'
    }

    $requestParameters = @{
        Uri         = $requestUrl
        Method      = 'Post'
        ContentType = 'application/json'
        Body        = ($payload | ConvertTo-Json -Depth 42)
        ErrorAction = 'Stop'
    }

    if ($PSEdition -eq 'Core' -and (Get-Command -Name Invoke-RestMethod).Parameters.ContainsKey('SkipCertificateCheck'))
    {
        $requestParameters.SkipCertificateCheck = $SkipCertificateCheck.IsPresent
    }

    if ($Credential)
    {
        $requestParameters.Credential = $Credential
    }
    else
    {
        $requestParameters.Headers = @{ Authorization = Get-TfsAccessTokenString -PersonalAccessToken $PersonalAccessToken }
    }

    try
    {
        $result = Invoke-RestMethod @requestParameters
        Write-Verbose -Message ('New release definition {0} created for project {1}' -f $ReleaseName, $ProjectName)
    }
    catch
    {
        Write-Error -ErrorRecord $_
    }
}

function Remove-TfsAgentUserCapability
{
    param
    (
        [Parameter(Mandatory = $true)]
        [string]
        $InstanceName,

        [Parameter()]
        [string]
        $CollectionName = 'DefaultCollection',

        [Parameter(Mandatory = $true)]
        [string]
        $PoolName = '*',

        [Parameter(Mandatory = $true, ParameterSetName = 'CredId')]
        [Parameter(Mandatory = $true, ParameterSetName = 'PatId')]
        [uint16]
        $AgentId,

        [Parameter(Mandatory = $true, ParameterSetName = 'CredObject')]
        [Parameter(Mandatory = $true, ParameterSetName = 'PatObject')]
        [object]
        $Agent,

        [ValidateRange(1, 65535)]
        [uint32]
        $Port,

        [string]
        $ApiVersion = '5.1',

        [switch]
        $UseSsl,

        [Parameter(Mandatory = $true, ParameterSetName = 'CredId')]
        [Parameter(Mandatory = $true, ParameterSetName = 'CredObject')]
        [pscredential]
        $Credential,
        
        [Parameter(Mandatory = $true, ParameterSetName = 'PatId')]
        [Parameter(Mandatory = $true, ParameterSetName = 'PatObject')]
        [string]
        $PersonalAccessToken,

        [switch]
        $SkipCertificateCheck,

        [Parameter(Mandatory = $true)]
        [string[]]
        $Capability
    )

    if ($SkipCertificateCheck.IsPresent)
    {
        $null = [ServerCertificateValidationCallback]::Ignore()
    }

    $poolParam = Sync-Parameter -Command (Get-Command Get-TfsAgentPool) -Parameter $PSBoundParameters
    $pool = Get-TfsAgentPool @poolParam

    if (-not $pool)
    {
        Write-Error -Message "Pool $PoolName could not be found!"
        return
    }

    if ($AgentId)
    {
        $agtParam = Sync-Parameter -Command (Get-Command Get-TfsAgent) -Parameter $PSBoundParameters
        $Agent = Get-TfsAgent @agtParam -Filter {$_.id -eq $AgentId}
    }

    if (-not $Agent)
    {
        Write-Error -Message "Agent could not be found!"
        return
    }

    $requestUrl = if ($UseSsl) {'https://' } else {'http://'}
    $requestUrl += if ( $Port  -gt 0)
    {
        '{0}{1}/{2}/_apis/distributedtask/pools/{3}/agents/{4}/usercapabilities' -f $InstanceName, ":$Port", $CollectionName, $pool.id, $Agent.Id
    }
    else
    {
        '{0}/{1}/_apis/distributedtask/pools/{2}/agents/{3}/usercapabilities' -f $InstanceName, $CollectionName, $pool.id, $Agent.Id
    }
    
    if ($ApiVersion)
    {
        $requestUrl += '?api-version={0}' -f $ApiVersion
    }

    $settableCapabilities = @{ }
    foreach ($prop in $Agent.usercapabilities.psobject.properties)
    {
        if ($prop.Name -in $Capability) { continue }
        $settableCapabilities[$prop.Name] = $prop.Value
    }

    $requestParameters = @{
        Uri             = $requestUrl
        Method          = 'Put'
        ContentType     = 'application/json'
        Body            = ($settableCapabilities | ConvertTo-Json)
        ErrorAction     = 'Stop'
        UseBasicParsing = $true
    }

    if ($PSEdition -eq 'Core' -and (Get-Command -Name Invoke-RestMethod).Parameters.ContainsKey('SkipCertificateCheck'))
    {
        $requestParameters.SkipCertificateCheck = $SkipCertificateCheck.IsPresent
    }

    if ($Credential)
    {
        $requestParameters.Credential = $Credential
    }
    else
    {
        $requestParameters.Headers = @{ Authorization = Get-TfsAccessTokenString -PersonalAccessToken $PersonalAccessToken }
    }

    try
    {
        $result = Invoke-RestMethod @requestParameters
    }
    catch
    {
        Write-Error -ErrorRecord $_
    }
    
    if ($result.value)
    {
        return $result.value
    }
    elseif ($result)
    {
        return $result
    }
}

function Remove-TfsFeed
{
    
    param
    (
        [Parameter(Mandatory = $true)]
        [string]
        $InstanceName,
 
        [Parameter()]
        [string]
        $CollectionName = 'DefaultCollection',
 
        [ValidateRange(1, 65535)]
        [uint32]
        $Port,
 
        [string]
        $ApiVersion = '1.0',
 
        [Parameter(Mandatory = $true)]
        [string]
        $FeedName,
 
        [switch]
        $UseSsl,
 
        [Parameter(ParameterSetName = 'Tfs')]
        [pscredential]
        $Credential,
         
        [Parameter(ParameterSetName = 'Vsts')]
        [string]
        $PersonalAccessToken,

        [switch]
        $SkipCertificateCheck
    )

    if ($SkipCertificateCheck.IsPresent)
    {
        $null = [ServerCertificateValidationCallback]::Ignore()
    }

    $feed = Get-TfsFeed @PSBoundParameters
    if (-not $feed)
    {
        Write-Warning "The feed '$FeedName' does not exist."
        return
    }
 
    $requestUrl = if ($UseSsl) {'https://' } else {'http://'}
    $requestUrl += if ($Port -gt 0)
    {
        '{0}{1}/{2}/_apis/packaging/feeds/{3}' -f $InstanceName, ":$Port", $CollectionName, $feed.id
    }
    else
    {
        '{0}/{1}/_apis/packaging/feeds/{2}' -f $InstanceName, $CollectionName, $feed.id
    }
    
    if ($ApiVersion)
    {
        $requestUrl += '?api-version={0}' -f $ApiVersion
    }
 
    $requestParameters = @{
        Uri         = $requestUrl
        Method      = 'Delete'
        ErrorAction = 'Stop'
    }

    if ($PSEdition -eq 'Core' -and (Get-Command -Name Invoke-RestMethod).Parameters.ContainsKey('SkipCertificateCheck'))
    {
        $requestParameters.SkipCertificateCheck = $SkipCertificateCheck.IsPresent
    }
 
    if ($Credential)
    {
        $requestParameters.Credential = $Credential
    }
    else
    {
        $requestParameters.Headers = @{ Authorization = Get-TfsAccessTokenString -PersonalAccessToken $PersonalAccessToken }
    }
 
    try
    {
        $result = Invoke-RestMethod @requestParameters
    }
    catch
    {
        if ($_.ErrorDetails.Message)
        {
            $errorDetails = $_.ErrorDetails.Message | ConvertFrom-Json
            if ($errorDetails.typeKey -eq 'ProjectDoesNotExistWithNameException')
            {
                return $null
            }
        }
        Write-Error -ErrorRecord $_
    }
     
    $data = if ($result.value)
    {
        $result.value
    }
    elseif ($result)
    {
        $result
    }

    if ($FeedName)
    {
        $data | Where-Object name -eq $FeedName
    }
    else
    {
        $data
    }
}

function Set-TfsAgentUserCapability
{
    param
    (
        [Parameter(Mandatory = $true)]
        [string]
        $InstanceName,

        [Parameter()]
        [string]
        $CollectionName = 'DefaultCollection',

        [Parameter(Mandatory = $true)]
        [string]
        $PoolName = '*',

        [Parameter(Mandatory = $true, ParameterSetName = 'CredId')]
        [Parameter(Mandatory = $true, ParameterSetName = 'PatId')]
        [uint16]
        $AgentId,

        [Parameter(Mandatory = $true, ParameterSetName = 'CredObject')]
        [Parameter(Mandatory = $true, ParameterSetName = 'PatObject')]
        [object]
        $Agent,

        [ValidateRange(1, 65535)]
        [uint32]
        $Port,

        [string]
        $ApiVersion = '5.1',

        [switch]
        $UseSsl,

        [Parameter(Mandatory = $true, ParameterSetName = 'CredId')]
        [Parameter(Mandatory = $true, ParameterSetName = 'CredObject')]
        [pscredential]
        $Credential,
        
        [Parameter(Mandatory = $true, ParameterSetName = 'PatId')]
        [Parameter(Mandatory = $true, ParameterSetName = 'PatObject')]
        [string]
        $PersonalAccessToken,

        [switch]
        $SkipCertificateCheck,

        [Parameter(Mandatory = $true)]
        [hashtable]
        $Capability
    )

    if ($SkipCertificateCheck.IsPresent)
    {
        $null = [ServerCertificateValidationCallback]::Ignore()
    }

    $poolParam = Sync-Parameter -Command (Get-Command Get-TfsAgentPool) -Parameter $PSBoundParameters
    $pool = Get-TfsAgentPool @poolParam

    if (-not $pool)
    {
        Write-Error -Message "Pool $PoolName could not be found!"
        return
    }

    if ($AgentId)
    {
        $agtParam = Sync-Parameter -Command (Get-Command Get-TfsAgent) -Parameter $PSBoundParameters
        $Agent = Get-TfsAgent @agtParam -Filter {$_.id -eq $AgentId}
    }

    if (-not $Agent)
    {
        Write-Error -Message "Agent could not be found!"
        return
    }

    $requestUrl = if ($UseSsl) {'https://' } else {'http://'}
    $requestUrl += if ( $Port  -gt 0)
    {
        '{0}{1}/{2}/_apis/distributedtask/pools/{3}/agents/{4}/usercapabilities' -f $InstanceName, ":$Port", $CollectionName, $pool.id, $Agent.Id
    }
    else
    {
        '{0}/{1}/_apis/distributedtask/pools/{2}/agents/{3}/usercapabilities' -f $InstanceName, $CollectionName, $pool.id, $Agent.Id
    }
    
    if ($ApiVersion)
    {
        $requestUrl += '?api-version={0}' -f $ApiVersion
    }

    $requestParameters = @{
        Uri             = $requestUrl
        Method          = 'Put'
        ContentType     = 'application/json'
        Body            = ($Capability | ConvertTo-Json)
        ErrorAction     = 'Stop'
        UseBasicParsing = $true
    }

    if ($PSEdition -eq 'Core' -and (Get-Command -Name Invoke-RestMethod).Parameters.ContainsKey('SkipCertificateCheck'))
    {
        $requestParameters.SkipCertificateCheck = $SkipCertificateCheck.IsPresent
    }

    if ($Credential)
    {
        $requestParameters.Credential = $Credential
    }
    else
    {
        $requestParameters.Headers = @{ Authorization = Get-TfsAccessTokenString -PersonalAccessToken $PersonalAccessToken }
    }

    try
    {
        $result = Invoke-RestMethod @requestParameters
    }
    catch
    {
        Write-Error -ErrorRecord $_
    }
    
    if ($result.value)
    {
        return $result.value
    }
    elseif ($result)
    {
        return $result
    }
}

function Set-TfsFeedPermission
{
    param
    (
        [Parameter(Mandatory = $true)]
        [string]
        $InstanceName,

        [Parameter()]
        [string]
        $CollectionName = 'DefaultCollection',

        [ValidateRange(1, 65535)]
        [uint32]
        $Port,

        [string]
        $ApiVersion = '2.0',

        [Parameter(Mandatory = $true)]
        [string]
        $FeedName,

        [Parameter(Mandatory = $true)]
        [object[]]
        $Permissions,

        [switch]
        $UseSsl,

        [Parameter(ParameterSetName = 'Tfs')]
        [pscredential]
        $Credential,
        
        [Parameter(ParameterSetName = 'Vsts')]
        [string]
        $PersonalAccessToken,

        [switch]
        $SkipCertificateCheck
    )

    if ($SkipCertificateCheck.IsPresent)
    {
        $null = [ServerCertificateValidationCallback]::Ignore()
    }
        
    $feedParameters = Sync-Parameter -Command (Get-Command -Name Get-TfsFeed) -Parameters $PSBoundParameters
    $feedParameters.ErrorAction = 'SilentlyContinue'
    $feed = Get-TfsFeed @feedParameters
    if (-not $feed)
    {
        Write-Error -Message "The Feed '$FeedName' does not exist"
        return
    }

    $requestUrl = if ($UseSsl) {'https://' } else {'http://'}
    $requestUrl += if ( $Port -gt 0)
    {
        '{0}{1}/{2}/_apis/packaging/Feeds/{3}/permissions' -f $InstanceName, ":$Port", $CollectionName, $feed.id
    }
    else
    {
        '{0}/{1}/_apis/packaging/Feeds/{2}/permissions' -f $InstanceName, $CollectionName, $feed.id
    }
    
    if ($ApiVersion)
    {
        $requestUrl += '?api-version={0}' -f $ApiVersion
    }

    $payload = @{
        body        = $Permissions
    }

    $requestParameters = @{
        Uri         = $requestUrl
        Method      = 'Patch'
        ContentType = 'application/json'
        Body        = ($Permissions | ConvertTo-Json)
        ErrorAction = 'Stop'
    }

    if ($PSEdition -eq 'Core' -and (Get-Command -Name Invoke-RestMethod).Parameters.ContainsKey('SkipCertificateCheck'))
    {
        $requestParameters.SkipCertificateCheck = $SkipCertificateCheck.IsPresent
    }

    if ($Credential)
    {
        $requestParameters.Credential = $Credential
    }
    else
    {
        $requestParameters.Headers = @{ Authorization = Get-TfsAccessTokenString -PersonalAccessToken $PersonalAccessToken }
    }

    try
    {
        $result = Invoke-RestMethod @requestParameters
    }
    catch
    {
        Write-Error -ErrorRecord $_
    }
}

function Set-TfsFeedPermissions
{
    
    param
    (
        [Parameter(Mandatory = $true)]
        [string]
        $InstanceName,

        [Parameter()]
        [string]
        $CollectionName = 'DefaultCollection',

        [ValidateRange(1, 65535)]
        [uint32]
        $Port,

        [string]
        $ApiVersion = '2.0',

        [Parameter(Mandatory = $true)]
        [string]
        $FeedName,

        [Parameter(Mandatory = $true)]
        [object[]]
        $Permissions,

        [switch]
        $UseSsl,

        [Parameter(ParameterSetName = 'Tfs')]
        [pscredential]
        $Credential,
        
        [Parameter(ParameterSetName = 'Vsts')]
        [string]
        $PersonalAccessToken,

        [switch]
        $SkipCertificateCheck
    )

    if ($SkipCertificateCheck.IsPresent)
    {
        $null = [ServerCertificateValidationCallback]::Ignore()
    }
        
    $feedParameters = Sync-Parameter -Command (Get-Command -Name Get-TfsFeed) -Parameters $PSBoundParameters
    $feedParameters.ErrorAction = 'SilentlyContinue'
    $feed = Get-TfsFeed @feedParameters
    if (-not $feed)
    {
        Write-Error -Message "The Feed '$FeedName' does not exist"
        return
    }

    $requestUrl = if ($UseSsl) {'https://' } else {'http://'}
    $requestUrl += if ( $Port -gt 0)
    {
        '{0}{1}/{2}/_apis/packaging/Feeds/{3}/permissions' -f $InstanceName, ":$Port", $CollectionName, $feed.id
    }
    else
    {
        '{0}/{1}/_apis/packaging/Feeds/{2}/permissions' -f $InstanceName, $CollectionName, $feed.id
    }
    
    if ($ApiVersion)
    {
        $requestUrl += '?api-version={0}' -f $ApiVersion
    }

    $payload = @{
        body        = $Permissions
    }

    $requestParameters = @{
        Uri         = $requestUrl
        Method      = 'Patch'
        ContentType = 'application/json'
        Body        = ($Permissions | ConvertTo-Json)
        ErrorAction = 'Stop'
    }

    if ($PSEdition -eq 'Core' -and (Get-Command -Name Invoke-RestMethod).Parameters.ContainsKey('SkipCertificateCheck'))
    {
        $requestParameters.SkipCertificateCheck = $SkipCertificateCheck.IsPresent
    }

    if ($Credential)
    {
        $requestParameters.Credential = $Credential
    }
    else
    {
        $requestParameters.Headers = @{ Authorization = Get-TfsAccessTokenString -PersonalAccessToken $PersonalAccessToken }
    }

    try
    {
        $result = Invoke-RestMethod @requestParameters
    }
    catch
    {
        Write-Error -ErrorRecord $_
    }
}

function Set-TfsProject
{
    
    param
    (
        [Parameter(Mandatory = $true)]
        [string]
        $InstanceName,

        [Parameter()]
        [string]
        $CollectionName = 'DefaultCollection',

        [ValidateRange(1, 65535)]
        [uint32]
        $Port,

        [string]
        $ApiVersion = '2.0',

        [Parameter(Mandatory = $true)]
        [string]
        $ProjectGuid,

        [string]
        $NewName,

        [string]
        $NewDescription,

        [switch]
        $UseSsl,

        [Parameter(ParameterSetName = 'Tfs')]
        [pscredential]
        $Credential,
        
        [Parameter(ParameterSetName = 'Vsts')]
        [string]
        $PersonalAccessToken,

        [switch]
        $SkipCertificateCheck
    )

    if ($SkipCertificateCheck.IsPresent)
    {
        $null = [ServerCertificateValidationCallback]::Ignore()
    }

    $requestUrl = if ($UseSsl) {'https://' } else {'http://'}
    $requestUrl += if ( $Port -gt 0)
    {
        '{0}{1}/{2}/_apis/projects/{3}' -f $InstanceName, ":$Port", $CollectionName, $ProjectGuid
    }
    else
    {
        '{0}/{1}/_apis/projects/{2}' -f $InstanceName, $CollectionName, $ProjectGuid
    }
    
    if ($ApiVersion)
    {
        $requestUrl += '?api-version={0}' -f $ApiVersion
    }

    $payload = @{
        name        = $NewName
        description = $NewDescription
    }

    $requestParameters = @{
        Uri         = $requestUrl
        Method      = 'Patch'
        ContentType = 'application/json'
        Body        = ($payload | ConvertTo-Json)
        ErrorAction = 'Stop'
    }

    if ($PSEdition -eq 'Core' -and (Get-Command -Name Invoke-RestMethod).Parameters.ContainsKey('SkipCertificateCheck'))
    {
        $requestParameters.SkipCertificateCheck = $SkipCertificateCheck.IsPresent
    }

    if ($Credential)
    {
        $requestParameters.Credential = $Credential
    }
    else
    {
        $requestParameters.Headers = @{ Authorization = Get-TfsAccessTokenString -PersonalAccessToken $PersonalAccessToken }
    }

    try
    {
        $result = Invoke-RestMethod @requestParameters
        Write-Verbose ('Project {0} renamed to {1}' -f $ProjectGuid, $NewName)
    }
    catch
    {
        Write-Error -ErrorRecord $_
    }
}

Export-ModuleMember -Function Add-AccountPrivilege,Add-FunctionToPSSession,Add-StringIncrement,Add-VariableToPSSession,Get-ConsoleText,Get-DotNetFrameworkVersion,Get-FullMesh,Get-ModuleDependency,Get-RunspacePool,Get-StringSection,Get-Type,Install-SoftwarePackage,Invoke-Ternary,New-RunspacePool,Read-Choice,Read-HashTable,Receive-RunspaceJob,Remove-RunspacePool,Send-ModuleToPsSession,Split-Array,Start-RunspaceJob,Sync-Parameter,Test-HashtableKeys,Test-IsAdministrator,Wait-RunspaceJob,Get-DscConfigurationImportedResource,Get-RequiredModulesFromMOF,Add-HostEntry,Clear-HostFile,ConvertTo-BinaryIp,ConvertTo-DecimalIp,ConvertTo-DottedDecimalIp,ConvertTo-Mask,ConvertTo-MaskLength,Get-BroadcastAddress,Get-HostEntry,Get-HostFile,Get-NetworkAddress,Get-NetworkRange,Get-NetworkSummary,Get-PublicIpAddress,Remove-HostEntry,Test-Port,Get-PerformanceCounterID,Get-PerformanceCounterLocalName,Get-PerformanceDataCollectorSet,New-PerformanceDataCollectorSet,Remove-PerformanceDataCollectorSet,Start-PerformanceDataCollectorSet,Stop-PerformanceDataCollectorSet,Add-CATemplateStandardPermission,Add-Certificate2,Enable-AutoEnrollment,Find-CertificateAuthority,Get-CaTemplate,Get-Certificate2,Get-NextOid,New-CaTemplate,Publish-CaTemplate,Request-Certificate,Test-CaTemplate,Add-TfsAgentUserCapability,Get-TfsAccessTokenString,Get-TfsAgent,Get-TfsAgentPool,Get-TfsAgentQueue,Get-TfsBuildDefinition,Get-TfsBuildDefinitionTemplate,Get-TfsBuildStep,Get-TfsFeed,Get-TfsFeedPermission,Get-TfsGitRepository,Get-TfsProcessTemplate,Get-TfsProject,Get-TfsReleaseDefinition,Get-TfsReleaseStep,New-TfsAgentQueue,New-TfsBuildDefinition,New-TfsFeed,New-TfsProject,New-TfsReleaseDefinition,Remove-TfsAgentUserCapability,Remove-TfsFeed,Set-TfsAgentUserCapability,Set-TfsFeedPermission,Set-TfsFeedPermissions,Set-TfsProject

