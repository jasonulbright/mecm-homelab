function Add-UnattendedWindowsNetworkAdapter
{
	param (
		[string]$Interfacename,

		[AutomatedLab.IPNetwork[]]$IpAddresses,

		[AutomatedLab.IPAddress[]]$Gateways,

		[AutomatedLab.IPAddress[]]$DnsServers,

        [string]$ConnectionSpecificDNSSuffix,

        [string]$DnsDomain,

        [string]$UseDomainNameDevolution,

        [string]$DNSSuffixSearchOrder,

        [string]$EnableAdapterDomainNameRegistration,

        [string]$DisableDynamicUpdate,

        [string]$NetbiosOptions
	)

    function Add-XmlGroup
    {
        param
        (
            [string]$XPath,
            [string]$ElementName,
            [string]$Action,
            [string]$KeyValue
        )

        Write-Debug -Message "XPath=$XPath"
        Write-Debug -Message "ElementName=$ElementName"

        #$ns = @{ un = 'urn:schemas-microsoft-com:unattend' }
        #$wcmNamespaceUrl = 'http://schemas.microsoft.com/WMIConfig/2002/State'

        $rootElement = $script:un | Select-Xml -XPath $XPath -Namespace $script:ns | Select-Object -ExpandProperty Node

        $element = $script:un.CreateElement($ElementName, $script:un.DocumentElement.NamespaceURI)
        [Void]$rootElement.AppendChild($element)
        #[Void]$element.SetAttribute('action', $script:wcmNamespaceUrl, 'add')
        if ($Action)   { [Void]$element.SetAttribute('action', $script:wcmNamespaceUrl, $Action) }
        if ($KeyValue) { [Void]$element.SetAttribute('keyValue', $script:wcmNamespaceUrl, $KeyValue) }
    }

    function Add-XmlElement
    {
        param
        (
            [string]$XPath,
            [string]$ElementName,
            [string]$Text,
            [string]$Action,
            [string]$KeyValue
        )

        Write-Debug -Message "XPath=$XPath"
        Write-Debug -Message "ElementName=$ElementName"
        Write-Debug -Message "Text=$Text"

        #$ns = @{ un = 'urn:schemas-microsoft-com:unattend' }
        #$wcmNamespaceUrl = 'http://schemas.microsoft.com/WMIConfig/2002/State'

        $rootElement = $script:un | Select-Xml -XPath $xPath -Namespace $script:ns | Select-Object -ExpandProperty Node

        $element = $script:un.CreateElement($elementName, $script:un.DocumentElement.NamespaceURI)
        [Void]$rootElement.AppendChild($element)
        if ($Action)   { [Void]$element.SetAttribute('action', $script:wcmNamespaceUrl, $Action) }
        if ($KeyValue) { [Void]$element.SetAttribute('keyValue', $script:wcmNamespaceUrl, $KeyValue) }
        $element.InnerText = $Text
    }

    $TCPIPInterfacesNode = '//un:settings[@pass = "specialize"]/un:component[@name = "Microsoft-Windows-TCPIP"]'


    if (-not ($script:un | Select-Xml -XPath "$TCPIPInterfacesNode/un:Interfaces" -Namespace $script:ns | Select-Object -ExpandProperty Node))
    {
        Add-XmlGroup -XPath "$TCPIPInterfacesNode" -ElementName 'Interfaces'
        $order = 1
    }

    Add-XmlGroup -XPath "$TCPIPInterfacesNode/un:Interfaces" -ElementName 'Interface' -Action 'add'
    Add-XmlGroup -XPath "$TCPIPInterfacesNode/un:Interfaces/un:Interface" -ElementName 'Ipv4Settings'
    Add-XmlElement -XPath "$TCPIPInterfacesNode/un:Interfaces/un:Interface/un:Ipv4Settings" -ElementName 'DhcpEnabled' -Text "$(([string](-not ([boolean]($ipAddresses -match '\.')))).ToLower())"
    #Add-XmlElement -XPath "$TCPIPInterfacesNode/un:Interfaces/un:Interface/un:Ipv4Settings" -ElementName 'Metric' -Text '10'
    #Add-XmlElement -XPath "$TCPIPInterfacesNode/un:Interfaces/un:Interface/un:Ipv4Settings" -ElementName 'RouterDiscoveryEnabled' -Text 'false'

    Add-XmlGroup -XPath "$TCPIPInterfacesNode/un:Interfaces/un:Interface" -ElementName 'Ipv6Settings'
    Add-XmlElement -XPath "$TCPIPInterfacesNode/un:Interfaces/un:Interface/un:Ipv6Settings" -ElementName 'DhcpEnabled' -Text "$(([string](-not ([boolean]($ipAddresses -match ':')))).ToLower())"
    #Add-XmlElement -XPath "$TCPIPInterfacesNode/un:Interfaces/un:Interface/un:Ipv6Settings" -ElementName 'Metric' -Text '10'
    #Add-XmlElement -XPath "$TCPIPInterfacesNode/un:Interfaces/un:Interface/un:Ipv6Settings" -ElementName 'RouterDiscoveryEnabled' -Text 'false'

    Add-XmlElement -XPath "$TCPIPInterfacesNode/un:Interfaces/un:Interface" -ElementName 'Identifier' -Text "$Interfacename"

    if ($IpAddresses)
	{
        Add-XmlGroup -XPath "$TCPIPInterfacesNode/un:Interfaces/un:Interface" -ElementName 'UnicastIpAddresses'
        $ipCount = 1
        foreach ($ipAddress in $IpAddresses)
        {
            Add-XmlElement -XPath "$TCPIPInterfacesNode/un:Interfaces/un:Interface/un:UnicastIpAddresses" -ElementName 'IpAddress' -Text "$($ipAddress.IpAddress.AddressAsString)/$($ipAddress.Cidr)" -Action 'add' -KeyValue "$(($ipCount++))"
        }
    }

    if ($gateways)
	{
        Add-XmlGroup -XPath "$TCPIPInterfacesNode/un:Interfaces/un:Interface" -ElementName 'Routes'
        $gatewayCount = 0
        foreach ($gateway in $gateways)
        {
            Add-XmlGroup -XPath "$TCPIPInterfacesNode/un:Interfaces/un:Interface/un:Routes" -ElementName 'Route' -Action 'add'
            Add-XmlElement -XPath "$TCPIPInterfacesNode/un:Interfaces/un:Interface/un:Routes/un:Route" -ElementName 'Identifier' -Text "$(($gatewayCount++))"
            #Add-XmlElement -XPath "$TCPIPInterfacesNode/un:Interfaces/un:Interface/un:Routes/un:Route" -ElementName 'Metric' -Text '0'
            Add-XmlElement -XPath "$TCPIPInterfacesNode/un:Interfaces/un:Interface/un:Routes/un:Route" -ElementName 'NextHopAddress' -Text $gateway
            if ($gateway -match ':')
            {
                $prefix = '::/0'
            }
            else
            {
                $prefix = '0.0.0.0/0'
            }
            Add-XmlElement -XPath "$TCPIPInterfacesNode/un:Interfaces/un:Interface/un:Routes/un:Route" -ElementName 'Prefix' -Text $prefix
        }

    }

    $DNSClientNode = '//un:settings[@pass = "specialize"]/un:component[@name = "Microsoft-Windows-DNS-Client"]'

    #if ($UseDomainNameDevolution)
    #{
    #    Add-XmlElement -XPath "$DNSClientNode" -ElementName 'UseDomainNameDevolution' -Text "$UseDomainNameDevolution"
    #}

    if ($DNSSuffixSearchOrder)
    {
        if (-not ($script:un | Select-Xml -XPath "$DNSClientNode/un:DNSSuffixSearchOrder" -Namespace $script:ns | Select-Object -ExpandProperty Node))
        {
            Add-XmlGroup -XPath "$DNSClientNode" -ElementName 'DNSSuffixSearchOrder' -Action 'add'
            $order = 1
        }
        else
        {
            $nodes = ($script:un | Select-Xml -XPath "$DNSClientNode/un:DNSSuffixSearchOrder" -Namespace $script:ns  | Select-Object -ExpandProperty Node).childnodes
            $order = ($nodes | Measure-Object).count+1
        }

        foreach ($DNSSuffix in $DNSSuffixSearchOrder)
        {
            Add-XmlElement -XPath "$DNSClientNode/un:DNSSuffixSearchOrder" -ElementName 'DomainName' -Text $DNSSuffix -Action 'add' -KeyValue "$(($order++))"
        }
    }

    if (-not ($script:un | Select-Xml -XPath "$DNSClientNode/un:Interfaces" -Namespace $script:ns | Select-Object -ExpandProperty Node))
    {
        Add-XmlGroup -XPath "$DNSClientNode" -ElementName 'Interfaces'
        $order = 1
    }

    Add-XmlGroup -XPath "$DNSClientNode/un:Interfaces" -ElementName 'Interface' -Action 'add'
    Add-XmlElement -XPath "$DNSClientNode/un:Interfaces/un:Interface" -ElementName 'Identifier' -Text "$Interfacename"

    if ($DnsDomain)
    {
        Add-XmlElement -XPath "$DNSClientNode/un:Interfaces/un:Interface" -ElementName 'DNSDomain' -Text "$DnsDomain"
    }

    if ($dnsServers)
	{
        Add-XmlGroup -XPath "$DNSClientNode/un:Interfaces/un:Interface" -ElementName 'DNSServerSearchOrder'
        $dnsServersCount = 1
        foreach ($dnsServer in $dnsServers)
        {
            Add-XmlElement -XPath "$DNSClientNode/un:Interfaces/un:Interface/un:DNSServerSearchOrder" -ElementName 'IpAddress' -Text $dnsServer -Action 'add' -KeyValue "$(($dnsServersCount++))"
        }
    }

    Add-XmlElement -XPath "$DNSClientNode/un:Interfaces/un:Interface" -ElementName 'EnableAdapterDomainNameRegistration' -Text $EnableAdapterDomainNameRegistration

    Add-XmlElement -XPath "$DNSClientNode/un:Interfaces/un:Interface" -ElementName 'DisableDynamicUpdate' -Text $DisableDynamicUpdate


    $NetBTNode = '//un:settings[@pass = "specialize"]/un:component[@name = "Microsoft-Windows-NetBT"]'

    if (-not ($script:un | Select-Xml -XPath "$NetBTNode/un:Interfaces" -Namespace $script:ns | Select-Object -ExpandProperty Node))
    {
        Add-XmlGroup -XPath "$NetBTNode" -ElementName 'Interfaces'
    }

    Add-XmlGroup -XPath "$NetBTNode/un:Interfaces" -ElementName 'Interface' -Action 'add'
    Add-XmlElement -XPath "$NetBTNode/un:Interfaces/un:Interface" -ElementName 'NetbiosOptions' -Text $NetbiosOptions
    Add-XmlElement -XPath "$NetBTNode/un:Interfaces/un:Interface" -ElementName 'Identifier' -Text "$Interfacename"

}

function Add-UnattendedWindowsPreinstallationCommand {
    [CmdletBinding()]
    param ()

    Write-PSFMessage -Message "No Preinstall implemented yet with Windows"
}

function Add-UnattendedWindowsRenameNetworkAdapters
{
    function Add-XmlGroup
    {
        param
        (
            $XPath,
            $ElementName,
            $Action,
            $KeyValue
        )

        Write-Debug -Message "XPath=$XPath"
        Write-Debug -Message "ElementName=$ElementName"

        #$ns = @{ un = 'urn:schemas-microsoft-com:unattend' }
        #$wcmNamespaceUrl = 'http://schemas.microsoft.com/WMIConfig/2002/State'

        $rootElement = $script:un | Select-Xml -XPath $xPath -Namespace $script:ns | Select-Object -ExpandProperty Node

        $element = $script:un.CreateElement($elementName)
        [Void]$rootElement.AppendChild($element)
        #[Void]$element.SetAttribute('action', $script:wcmNamespaceUrl, 'add')
        if ($Action)   { [Void]$element.SetAttribute('action', $script:wcmNamespaceUrl, $Action) }
        if ($KeyValue) { [Void]$element.SetAttribute('keyValue', $script:wcmNamespaceUrl, $KeyValue) }
    }

    function Add-XmlElement
    {
        param
        (
            $rootElement,
            $ElementName,
            $Text,
            $Action,
            $KeyValue
        )

        Write-Debug -Message "XPath=$XPath"
        Write-Debug -Message "ElementName=$ElementName"
        Write-Debug -Message "Text=$Text"

        #$ns = @{ un = 'urn:schemas-microsoft-com:unattend' }
        #$wcmNamespaceUrl = 'http://schemas.microsoft.com/WMIConfig/2002/State'

        #$rootElement = $script:un | Select-Xml -XPath $xPath -Namespace $script:ns | Select-Object -ExpandProperty Node

        $element = $script:un.CreateElement($elementName)
        [Void]$rootElement.AppendChild($element)
        if ($Action)   { [Void]$element.SetAttribute('action', $script:wcmNamespaceUrl, $Action) }
        if ($KeyValue) { [Void]$element.SetAttribute('keyValue', $script:wcmNamespaceUrl, $KeyValue) }
        $element.InnerText = $Text
    }

    $order = (($script:un | Select-Xml -XPath "$WinPENode/un:RunSynchronousCommand" -Namespace $script:ns).node.childnodes.order | Measure-Object -Maximum).maximum
    $order++

    Add-XmlGroup -XPath '//un:settings[@pass = "oobeSystem"]/un:component[@name = "Microsoft-Windows-Shell-Setup"]/un:FirstLogonCommands' -ElementName 'SynchronousCommand' -Action 'add'

    $nodes = ($script:un | Select-Xml -XPath '//un:settings[@pass = "oobeSystem"]/un:component[@name = "Microsoft-Windows-Shell-Setup"]/un:FirstLogonCommands' -Namespace $script:ns  |
	Select-Object -ExpandProperty Node).childnodes

    $order = ($nodes | Measure-Object).count
    $rootElement = $nodes[$order-1]

    Add-XmlElement -RootElement $rootElement -ElementName 'Description' -Text 'Rename network adapters'
    Add-XmlElement -RootElement $rootElement -ElementName 'Order' -Text "$order"
    Add-XmlElement -RootElement $rootElement -ElementName 'CommandLine' -Text 'powershell.exe -executionpolicy bypass -file "c:\RenameNetworkAdapters.ps1"'

}

function Add-UnattendedWindowsSynchronousCommand
{
    param (
        [Parameter(Mandatory)]
        [string]$Command,

        [Parameter(Mandatory)]
        [string]$Description
    )

    $highestOrder = ($un | Select-Xml -Namespace $ns -XPath //un:RunSynchronous).Node.RunSynchronousCommand.Order |
    Sort-Object -Property { [int]$_ } -Descending |
    Select-Object -First 1

    $runSynchronousNode = ($un | Select-Xml -Namespace $ns -XPath //un:RunSynchronous).Node

    $runSynchronousCommandNode = $un.CreateElement('RunSynchronousCommand')

    [Void]$runSynchronousCommandNode.SetAttribute('action', $wcmNamespaceUrl, 'add')

    $runSynchronousCommandDescriptionNode = $un.CreateElement('Description')
    $runSynchronousCommandDescriptionNode.InnerText = $Description

    $runSynchronousCommandOrderNode = $un.CreateElement('Order')
    $runSynchronousCommandOrderNode.InnerText = ([int]$highestOrder + 1)

    $runSynchronousCommandPathNode = $un.CreateElement('Path')
    $runSynchronousCommandPathNode.InnerText = $Command

    [void]$runSynchronousCommandNode.AppendChild($runSynchronousCommandDescriptionNode)
    [void]$runSynchronousCommandNode.AppendChild($runSynchronousCommandOrderNode)
    [void]$runSynchronousCommandNode.AppendChild($runSynchronousCommandPathNode)

    [void]$runSynchronousNode.AppendChild($runSynchronousCommandNode)
}

function Add-UnattendedWinSshPublicKey
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [string]
        $PublicKey
    )

    Write-PSFMessage -Message "No unattended ssh key import on Windows yet, we're using synchronous command"
}


function Export-UnattendedWindowsFile
{
    param (
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $script:un.Save($Path)
}

function Import-UnattendedWindowsContent
{
    param
    (
        [Parameter(Mandatory = $true)]
        [xml]
        $Content
    )

    $script:un = $Content
    $script:ns = @{ un = 'urn:schemas-microsoft-com:unattend' }
    $Script:wcmNamespaceUrl = 'http://schemas.microsoft.com/WMIConfig/2002/State'
}

function Set-UnattendedWindowsDomain
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$DomainName,

		[Parameter(Mandatory = $true)]
		[string]$Username,

		[Parameter(Mandatory = $true)]
		[string]$Password,

		[Parameter()]
		[string]$OrganizationalUnit
	)

	$idNode = $script:un |
	Select-Xml -XPath '//un:settings[@pass = "specialize"]/un:component[@name = "Microsoft-Windows-UnattendedJoin"]/un:Identification' -Namespace $ns |
	Select-Object -ExpandProperty Node

	$idNode.RemoveAll()

	$joinDomainNode = $script:un.CreateElement('JoinDomain')
	$joinDomainNode.InnerText = $DomainName

	$credentialsNode = $script:un.CreateElement('Credentials')
	$domainNode = $script:un.CreateElement('Domain')
	$domainNode.InnerText = $DomainName
	$userNameNode = $script:un.CreateElement('Username')
	$userNameNode.InnerText = $Username
	$passwordNode = $script:un.CreateElement('Password')
	$passwordNode.InnerText = $Password

	if ($OrganizationalUnit)
	{
		$ouNode = $script:un.CreateElement('MachineObjectOU')
		$ouNode.InnerText = $OrganizationalUnit
		$null = $idNode.AppendChild($ouNode)
	}

	[Void]$credentialsNode.AppendChild($domainNode)
	[Void]$credentialsNode.AppendChild($userNameNode)
	[Void]$credentialsNode.AppendChild($passwordNode)

	[Void]$idNode.AppendChild($credentialsNode)
	[Void]$idNode.AppendChild($joinDomainNode)
}


function Set-UnattendedWindowsAdministratorName
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$Name
	)

	$shellNode = $script:un |
	Select-Xml -XPath '//un:settings[@pass = "oobeSystem"]/un:component[@name = "Microsoft-Windows-Shell-Setup"]' -Namespace $ns |
	Select-Object -ExpandProperty Node

	$shellNode.UserAccounts.LocalAccounts.LocalAccount.Name = $Name
	$shellNode.UserAccounts.LocalAccounts.LocalAccount.DisplayName = $Name
}

function Set-UnattendedWindowsAdministratorPassword
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$Password
	)

	$shellNode = $script:un |
	Select-Xml -XPath '//un:settings[@pass = "oobeSystem"]/un:component[@name = "Microsoft-Windows-Shell-Setup"]' -Namespace $ns |
	Select-Object -ExpandProperty Node

	$shellNode.UserAccounts.AdministratorPassword.Value = $Password
	$shellNode.UserAccounts.AdministratorPassword.PlainText = 'true'

	$shellNode.UserAccounts.LocalAccounts.LocalAccount.Password.Value = $Password
}

function Set-UnattendedWindowsAntiMalware
{
    param (
        [Parameter(Mandatory = $true)]
        [bool]$Enabled
    )

    $node = $script:un |
    Select-Xml -XPath '//un:settings[@pass = "specialize"]/un:component[@name = "Security-Malware-Windows-Defender"]' -Namespace $ns |
    Select-Object -ExpandProperty Node

    if ($Enabled)
    {
        $node.DisableAntiSpyware = 'true'
    }
    else
    {
        $node.DisableAntiSpyware = 'false'
    }
}

function Set-UnattendedWindowsAutoLogon
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$DomainName,

		[Parameter(Mandatory = $true)]
		[string]$Username,

		[Parameter(Mandatory = $true)]
		[string]$Password
	)

	$shellNode = $script:un |
	Select-Xml -XPath '//un:settings[@pass = "specialize"]/un:component[@name = "Microsoft-Windows-Shell-Setup"]' -Namespace $ns |
	Select-Object -ExpandProperty Node

	$autoLogonNode = $script:un.CreateElement('AutoLogon')

	$passwordNode = $script:un.CreateElement('Password')
	$passwordValueNode = $script:un.CreateElement('Value')
	$passwordValueNode.InnerText = $Password

	$domainNode = $script:un.CreateElement('Domain')
	$domainNode.InnerText = $DomainName

	$enabledNode = $script:un.CreateElement('Enabled')
	$enabledNode.InnerText = 'true'

	$logonCount = $script:un.CreateElement('LogonCount')
	$logonCount.InnerText = '9999'

	$userNameNode = $script:un.CreateElement('Username')
	$userNameNode.InnerText = $Username

	[Void]$autoLogonNode.AppendChild($passwordNode)
	[Void]$passwordNode.AppendChild($passwordValueNode)
	[Void]$autoLogonNode.AppendChild($domainNode)
	[Void]$autoLogonNode.AppendChild($enabledNode)
	[Void]$autoLogonNode.AppendChild($logonCount)
	[Void]$autoLogonNode.AppendChild($userNameNode)

	[Void]$shellNode.AppendChild($autoLogonNode)
}


function Set-UnattendedWindowsComputerName
{
    param (
        [Parameter(Mandatory = $true)]
        [string]$ComputerName
    )
    $component = $script:un |
	Select-Xml -XPath '//un:settings[@pass = "specialize"]/un:component[@name = "Microsoft-Windows-Shell-Setup"]' -Namespace $ns |
	Select-Object -ExpandProperty Node

	$component.ComputerName = $ComputerName
}

function Set-UnattendedWindowsFirewallState
{
	param (
		[Parameter(Mandatory = $true)]
		[boolean]$State
	)

    $setupNode = $script:un |
	Select-Xml -XPath '//un:settings[@pass = "specialize"]/un:component[@name = "Networking-MPSSVC-Svc"]' -Namespace $ns |
	Select-Object -ExpandProperty Node

	$WindowsFirewallStateNode = $script:un.CreateElement('DomainProfile_EnableFirewall')
	$WindowsFirewallStateNode.InnerText = ([string]$State).ToLower()
	[Void]$setupNode.AppendChild($WindowsFirewallStateNode)

	$WindowsFirewallStateNode = $script:un.CreateElement('PrivateProfile_EnableFirewall')
	$WindowsFirewallStateNode.InnerText = ([string]$State).ToLower()
	[Void]$setupNode.AppendChild($WindowsFirewallStateNode)

	$WindowsFirewallStateNode = $script:un.CreateElement('PublicProfile_EnableFirewall')
	$WindowsFirewallStateNode.InnerText = ([string]$State).ToLower()
	[Void]$setupNode.AppendChild($WindowsFirewallStateNode)
}

function Set-UnattendedWindowsIpSettings
{
	param (
		[string]$IpAddress,

		[string]$Gateway,

		[String[]]$DnsServers,

        [string]$DnsDomain
	)

    $ethernetInterface = $script:un |
	Select-Xml -XPath '//un:settings[@pass = "specialize"]/un:component[@name = "Microsoft-Windows-TCPIP"]/un:Interfaces/un:Interface[un:Identifier = "Ethernet"]' -Namespace $ns |
	Select-Object -ExpandProperty Node

	if (-not $ethernetInterface)
	{
		$ethernetInterface = $script:un |
		Select-Xml -XPath '//un:settings[@pass = "specialize"]/un:component[@name = "Microsoft-Windows-TCPIP"]/un:Interfaces/un:Interface[un:Identifier = "Local Area Connection"]' -Namespace $ns |
		Select-Object -ExpandProperty Node
	}

	if ($IpAddress)
	{
		$ethernetInterface.Ipv4Settings.DhcpEnabled = 'false'
		$ethernetInterface.UnicastIpAddresses.IpAddress.InnerText = $IpAddress
	}

	if ($Gateway)
	{
		$InterfaceElement = $script:un |
		Select-Xml -XPath '//un:settings[@pass = "specialize"]/un:component[@name = "Microsoft-Windows-TCPIP"]/un:Interfaces/un:Interface' -Namespace $ns |
		Select-Object -ExpandProperty Node

		$RoutesNode = $script:un.CreateElement('Routes')
		[Void]$InterfaceElement.AppendChild($RoutesNode)

		$routes = $script:un |
		Select-Xml -XPath '//un:settings[@pass = "specialize"]/un:component[@name = "Microsoft-Windows-TCPIP"]/un:Interfaces/un:Interface/un:Routes' -Namespace $ns |
		Select-Object -ExpandProperty Node

		$routeElement = $script:un.CreateElement('Route')
		$identifierElement = $script:un.CreateElement('Identifier')
		$prefixElement = $script:un.CreateElement('Prefix')
		$nextHopAddressElement = $script:un.CreateElement('NextHopAddress')
		[void]$routeElement.AppendChild($identifierElement)
		[void]$routeElement.AppendChild($prefixElement)
		[void]$routeElement.AppendChild($nextHopAddressElement)

		[Void]$routeElement.SetAttribute('action', $wcmNamespaceUrl, 'add')
		$identifierElement.InnerText = '0'
		$prefixElement.InnerText = '0.0.0.0/0'
		$nextHopAddressElement.InnerText = $Gateway

		[void]$RoutesNode.AppendChild($routeElement)
	}

  <#
    <Routes>
    <Route wcm:action="add">
    <Identifier>0</Identifier>
    <Prefix>0.0.0.0/0</Prefix>
    <NextHopAddress></NextHopAddress>
    </Route>
    </Routes>
  #>

	if ($DnsServers)
	{
		$ethernetInterface = $script:un |
		Select-Xml -XPath '//un:settings[@pass = "specialize"]/un:component[@name = "Microsoft-Windows-DNS-Client"]/un:Interfaces/un:Interface[un:Identifier = "Ethernet"]' -Namespace $ns |
		Select-Object -ExpandProperty Node -ErrorAction SilentlyContinue

		if (-not $ethernetInterface)
		{
			$ethernetInterface = $script:un |
			Select-Xml -XPath '//un:settings[@pass = "specialize"]/un:component[@name = "Microsoft-Windows-DNS-Client"]/un:Interfaces/un:Interface[un:Identifier = "Local Area Connection"]' -Namespace $ns |
			Select-Object -ExpandProperty Node -ErrorAction SilentlyContinue
		}

    <#
        <DNSServerSearchOrder>
        <IpAddress wcm:action="add" wcm:keyValue="1">10.0.0.10</IpAddress>
        </DNSServerSearchOrder>
    #>

		$dnsServerSearchOrder = $script:un.CreateElement('DNSServerSearchOrder')
		$i = 1
		foreach ($dnsServer in $DnsServers)
		{
			$ipAddressElement = $script:un.CreateElement('IpAddress')
			[Void]$ipAddressElement.SetAttribute('action', $wcmNamespaceUrl, 'add')
			[Void]$ipAddressElement.SetAttribute('keyValue', $wcmNamespaceUrl, "$i")
			$ipAddressElement.InnerText = $dnsServer

			[Void]$dnsServerSearchOrder.AppendChild($ipAddressElement)
			$i++
		}

		[Void]$ethernetInterface.AppendChild($dnsServerSearchOrder)
	}

    <#
        <DNSDomain>something.com</DNSDomain>
    #>
    if ($DnsDomain)
    {
        $ethernetInterface = $script:un |
		Select-Xml -XPath '//un:settings[@pass = "specialize"]/un:component[@name = "Microsoft-Windows-DNS-Client"]/un:Interfaces/un:Interface[un:Identifier = "Ethernet"]' -Namespace $ns |
		Select-Object -ExpandProperty Node -ErrorAction SilentlyContinue

		if (-not $ethernetInterface)
		{
			$ethernetInterface = $script:un |
			Select-Xml -XPath '//un:settings[@pass = "specialize"]/un:component[@name = "Microsoft-Windows-DNS-Client"]/un:Interfaces/un:Interface[un:Identifier = "Local Area Connection"]' -Namespace $ns |
			Select-Object -ExpandProperty Node -ErrorAction SilentlyContinue
		}

		$dnsDomainElement = $script:un.CreateElement('DNSDomain')
		$dnsDomainElement.InnerText = $DnsDomain

		[Void]$ethernetInterface.AppendChild($dnsDomainElement)
    }
}

function Set-UnattendedWindowsLocalIntranetSites
{
	param (
		[Parameter(Mandatory = $true)]
		[string[]]$Values
	)

    $ieNode = $script:un |
	Select-Xml -XPath '//un:settings[@pass = "specialize"]/un:component[@name = "Microsoft-Windows-IE-InternetExplorer"]' -Namespace $ns |
	Select-Object -ExpandProperty Node

    $ieNode.LocalIntranetSites = $Values -join ';'
}

function Set-UnattendedWindowsPackage
{
    param
    (
        [string[]]$Package
    )
}

function Set-UnattendedWindowsProductKey
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$ProductKey
	)

	$setupNode = $script:un |
	Select-Xml -XPath '//un:settings[@pass = "specialize"]/un:component[@name = "Microsoft-Windows-Shell-Setup"]' -Namespace $ns |
	Select-Object -ExpandProperty Node

	$productKeyNode = $script:un.CreateElement('ProductKey')
	$productKeyNode.InnerText = $ProductKey
	[Void]$setupNode.AppendChild($productKeyNode)
}

function Set-UnattendedWindowsTimeZone
{
    param
    (
        [Parameter(Mandatory = $true)]
        [string]$TimeZone
    )

    $component = $script:un |
        Select-Xml -XPath '//un:settings[@pass = "specialize"]/un:component[@name = "Microsoft-Windows-Shell-Setup"]' -Namespace $ns |
        Select-Object -ExpandProperty Node

    $component.TimeZone = $TimeZone
}


function Set-UnattendedWindowsUserLocale {
	param (
		[Parameter(Mandatory = $true)]
		[string]$UserLocale
	)

	if (-not $script:un) {
		Write-Error 'No unattended file imported. Please use Import-UnattendedFile first'
		return
	}

	$component = $script:un |
	Select-Xml -XPath '//un:settings[@pass = "oobeSystem"]/un:component[@name = "Microsoft-Windows-International-Core"]' -Namespace $ns |
	Select-Object -ExpandProperty Node

	#this is for getting the input locale strings like '0409:00000409'
	$component.UserLocale = $UserLocale

	$inputLocale = [System.Collections.Generic.List[string]]::new()
	$inputLocale.Add($languageList[$UserLocale])
	$inputLocale.Add($languageList['en-us'])
	$component.InputLocale = ($inputLocale -join ';')
}


function Set-UnattendedWindowsWorkgroup
{
    param
    (
		[Parameter(Mandatory = $true)]
        [string]
        $WorkgroupName
    )

    $idNode = $script:un |
	Select-Xml -XPath '//un:settings[@pass = "specialize"]/un:component[@name = "Microsoft-Windows-UnattendedJoin"]/un:Identification' -Namespace $ns |
	Select-Object -ExpandProperty Node

	$idNode.RemoveAll()

	$workGroupNode = $script:un.CreateElement('JoinWorkgroup')
	$workGroupNode.InnerText = $WorkgroupName
    [Void]$idNode.AppendChild($workGroupNode)
}

function Write-UnattendedWindowsFile
{
    param
    (
        [string]
        $Content,

        [string]
        $DestinationPath,

        [switch]
        $Append
    )
   Write-PSFMessage -Message 'Unattended Windows File Not Implemented Yet'
}


function Add-UnattendedCloudInitNetworkAdapter
{
    param (
        [string]$InterfaceName,

        [AutomatedLab.IPNetwork[]]$IpAddresses,

        [AutomatedLab.IPAddress[]]$Gateways,

        [AutomatedLab.IPAddress[]]$DnsServers
    )

    $macAddress = ($Interfacename -replace '-', ':').ToLower()
    if (-not $script:un['autoinstall']['network'].ContainsKey('ethernets`'))
    {
        $script:un['autoinstall']['network']['ethernets'] = @{ }
    }

    if ($script:un['autoinstall']['network']['ethernets'].Keys.Count -eq 0)
    {
        $ifName = 'en0'
    }
    else
    {
        [int]$lastIfIndex = ($script:un['autoinstall']['network']['ethernets'].Keys.GetEnumerator() | Sort-Object | Select-Object -Last 1) -replace 'en'
        $lastIfIndex++
        $ifName = 'en{0}' -f $lastIfIndex
    }

    $script:un['autoinstall']['network']['ethernets'][$ifName] = @{
        match      = @{
            macaddress = $macAddress
        }
        'set-name' = $ifName
    }

    $adapterAddress = $IpAddresses | Select-Object -First 1

    if (-not $adapterAddress)
    {
        $script:un['autoinstall']['network']['ethernets'][$ifName]['dhcp4'] = 'yes'
        $script:un['autoinstall']['network']['ethernets'][$ifName]['dhcp6'] = 'yes'
    }
    else
    {
        $script:un['autoinstall']['network']['ethernets'][$ifName]['addresses'] = @()
        foreach ($ip in $IpAddresses)
        {
            $script:un['autoinstall']['network']['ethernets'][$ifName]['addresses'] += '{0}/{1}' -f $ip.IPAddress.AddressAsString, $ip.SerializationCidr
        }
    }

    if ($Gateways -and -not $script:un['autoinstall']['network']['ethernets'][$ifName].ContainsKey('routes')) { $script:un['autoinstall']['network']['ethernets'][$ifName].routes = @() }
    foreach ($gw in $Gateways)
    {
        $script:un['autoinstall']['network']['ethernets'][$ifName]['routes'] += @{
            to  = 'default'
            via = $gw.AddressAsString
        }
    }

    if ($DnsServers)
    {
        $script:un['autoinstall']['network']['ethernets'][$ifName]['nameservers'] = @{ addresses = [string[]]($DnsServers.AddressAsString) }
    }
}


function Add-UnattendedCloudInitPreinstallationCommand {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$Command,

        [Parameter(Mandatory)]
        [string]$Description
    )

    # Ensure that installer runs to completion by returning with exit code 0
    if (-not $script:un['autoinstall'].Contains('early-commands')) {
        $script:un['autoinstall']['early-commands'] = [System.Collections.Generic.List[string]]::new()
    }

    $Command = "$Command; exit 0"
    $script:un['autoinstall']['early-commands'].Add($Command)
}


function Add-UnattendedCloudInitRenameNetworkAdapters
{
	[CmdletBinding()]
    param
    (
    )
	
    Write-PSFMessage -Message 'Method not required on Ubuntu/Cloudinit'
}


function Add-UnattendedCloudInitSshPublicKey
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [string]
        $PublicKey
    )

    foreach ($user in $script:un['autoinstall']['user-data']['users'])
    {
        if ($user -eq 'default') { continue }
        if (-not $user.ContainsKey('ssh_authorized_keys'))
        {
            $user.Add('ssh_authorized_keys', @())
        }

        $user['ssh_authorized_keys'] += $PublicKey
    }
}


function Add-UnattendedCloudInitSynchronousCommand
{
	[CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$Command,

        [Parameter(Mandatory)]
        [string]$Description
    )

    # Ensure that installer runs to completion by returning with exit code 0
    if ($Command -notlike 'curtin in-target --target=/target -*')
    {
        $Command = "curtin in-target --target=/target -- $Command"
    }

    $Command = "$Command; exit 0"
    $script:un['autoinstall']['late-commands'] += $Command
}


function Export-UnattendedCloudInitFile
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateScript({ Test-Path -Path $_ -PathType Container })]
        [string]$Path
    )

    # Cloud-init -> User Data cannot contain networking information
    # 
    $metadataDictionary = @{
        'instance-id'    = $Script:un['autoinstall']['user-data']['hostname']
        'local-hostname' = $Script:un['autoinstall']['user-data']['hostname']
    }

    ("#cloud-config`n{0}" -f ($script:un | ConvertTo-Yaml -Options DisableAliases)) | Set-Content -Path (Join-Path -Path $Path -ChildPath user-data) -Force
    ("#cloud-config`n{0}" -f ($metadataDictionary | ConvertTo-Yaml -Options DisableAliases)) | Set-Content -Path (Join-Path -Path $Path -ChildPath meta-data) -Force
}

function Import-UnattendedCloudInitContent
{
	[CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string[]]
        $Content
    )

    $script:un = $Content -join "`r`n" | ConvertFrom-Yaml
}


function Set-UnattendedCloudInitAdministratorName
{
    param
    (
        [Parameter(Mandatory = $true)]
        [string]
        $Name
    )

    $usr = @{
        name        = $Name
        shell       = '/bin/bash'
        lock_passwd = $false
        sudo        = 'ALL=(ALL) NOPASSWD:ALL'
    }

    if (-not $script:un['autoinstall']['user-data'].ContainsKey('users')) { $script:un['autoinstall']['user-data']['users'] = @() }

    if ($script:un['autoinstall']['user-data']['users']['name'] -notcontains $Name)
    {
        $script:un['autoinstall']['user-data']['users'] += $usr
    }
}


function Set-UnattendedCloudInitAdministratorPassword
{
    param
    (
        [Parameter(Mandatory = $true)]
        [string]
        $Password
    )

    $Script:un['autoinstall']['user-data']['password'] = $Password

    foreach ($user in $script:un['autoinstall']['user-data']['users'])
    {
        if ($user -eq 'default') { continue }
        $user['plain_text_passwd'] = $Password
    }
}


function Set-UnattendedCloudInitAntiMalware
{
	[CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [bool]$Enabled
    )

    Write-PSFMessage -Message "No anti-malware settings for CloudInit/Ubuntu"
}

function Set-UnattendedCloudInitAutoLogon
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$DomainName,

		[Parameter(Mandatory = $true)]
		[string]$Username,

		[Parameter(Mandatory = $true)]
		[string]$Password
    )
	
    Write-PSFMessage -Message "Auto-logon not implemented yet for CloudInit/Ubuntu"
}

function Set-UnattendedCloudInitComputerName
{
    param
    (
        [Parameter(Mandatory = $true)]
        [string]
        $ComputerName
    )

    $Script:un['autoinstall']['user-data']['hostname'] = $ComputerName.ToLower()
}


function Set-UnattendedCloudInitDomain
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$DomainName,

		[Parameter(Mandatory = $true)]
		[string]$Username,

		[Parameter(Mandatory = $true)]
		[string]$Password,

		[Parameter()]
		[string]$OrganizationalUnit
	)

	if ($script:un['autoinstall']['user-data']['hostname'])
	{
		$script:un['autoinstall']['user-data']['fqdn'] = '{0}.{1}' -f $script:un['autoinstall']['user-data']['hostname'].ToLower(), $DomainName
	}
	
	$script:un['autoinstall']['user-data']['write_files'] += @{
		append  = $false
		path    = '/etc/cron.d/00realmjoin'
		content = if ($OrganizationalUnit)
		{
			"@reboot root echo '{0}' | realm join --computer-ou='{2}' -U {3} {1}`n@reboot root pam-auth-update --enable mkhomedir`n" -f $Password, $DomainName, $OrganizationalUnit, $UserName
		}
		else
		{
			"@reboot root echo '{0}' | realm join -U {2} {1}`n@reboot root pam-auth-update --enable mkhomedir`n" -f $Password, $DomainName, $UserName
		}
	}
	$script:un['autoinstall']['user-data']['write_files'] += @{
		append  = $false
		path    = '/etc/sudoers.d/domainadmins'
		content = @"
# Allow Domain Admins
%Domain\ Admins@$($DomainName.ToUpper()) ALL=(ALL:ALL) NOPASSWD:ALL

"@
	}
	$script:un['autoinstall']['user-data']['write_files'] += @{
		append  = $false
		path    = '/etc/cron.d/99realmjoin'
		content = "@reboot root rm -rf /etc/cron.d/00realmjoin`n"
	}
}


function Set-UnattendedCloudInitFirewallState
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[boolean]$State
	)

    $script:un['autoinstall']['late-commands'] += 'curtin in-target --target=/target -- ufw enable 2>/dev/null || true'
    $script:un['autoinstall']['late-commands'] += 'curtin in-target --target=/target -- ufw allow 22 2>/dev/null || true'
}

function Set-UnattendedCloudInitIpSettings
{
    [CmdletBinding()]
    param (
        [string]$IpAddress,

        [string]$Gateway,

        [String[]]$DnsServers,

        [string]$DnsDomain
    )

    $ifName = 'en0'

    $script:un['autoinstall']['network']['ethernets'][$ifName] = @{
        match      = @{
            macAddress = $macAddress
        }
        'set-name' = $ifName
    }

    $adapterAddress = $IpAddress

    if (-not $adapterAddress)
    {
        $script:un['autoinstall']['network']['ethernets'][$ifName]['dhcp4'] = 'yes'
        $script:un['autoinstall']['network']['ethernets'][$ifName]['dhcp6'] = 'yes'
    }
    else
    {
        $script:un['autoinstall']['network']['ethernets'][$ifName]['addresses'] = @(
            $IpAddress
        )
    }

    if ($Gateway -and -not $script:un['autoinstall']['network']['ethernets'][$ifName].ContainsKey('routes')) 
    {
        $script:un['autoinstall']['network']['ethernets'][$ifName]['routes'] = @(
            @{
                to  = 'default'
                via = $Gateway
            })
    }

    if ($DnsServers)
    {
        $script:un['autoinstall']['network']['ethernets'][$ifName]['nameservers'] = @{ addresses = $DnsServers }
    }
}

function Set-UnattendedCloudInitLocalIntranetSites
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string[]]$Values
	)
	
	Write-PSFMessage -Message 'No local intranet sites for CloudInit/Ubuntu'
}

function Set-UnattendedCloudInitPackage
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string[]]$Package,

        [bool]
        $IsSnap = $false
    )

    if ($IsSnap -and -not $script:un['autoinstall'].Contains('snaps')) {
        $script:un['autoinstall']['snaps'] = [System.Collections.Generic.List[hashtable]]::new()
    }

    foreach ($pack in $Package)
    {
        if ($pack -in $script:un['autoinstall']['packages'] -or $pack -in $script:un['autoinstall']['snaps'].name) { continue }
        
        if ($IsSnap) {
            $script:un['autoinstall']['snaps'].Add(@{
                name = $pack
            })
            continue
        }

        $script:un['autoinstall']['packages'] += $pack
    }
}

function Set-UnattendedProductKey
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$ProductKey
	)

	Write-PSFMessage "No product key required on CloudInit/Ubuntu"
}

function Set-UnattendedCloudInitTimeZone
{
	[CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [string]$TimeZone
    )

    $tzInfo = Get-TimeZone -Id $TimeZone -ErrorAction SilentlyContinue

    if (-not $tzInfo) { Get-TimeZone }

    Write-PSFMessage -Message ('Since non-standard timezone names are used, we revert to Etc/GMT{0}' -f $tzInfo.BaseUtcOffset.TotalHours)
    $tzname = if ($tzInfo.BaseUtcOffset.TotalHours -gt 0)
    {
        'Etc/GMT+{0}' -f $tzInfo.BaseUtcOffset.TotalHours
    }
    elseif ($tzInfo.BaseUtcOffset.TotalHours -eq 0)
    {
        'Etc/GMT'
    }
    else
    {
        'Etc/GMT{0}' -f $tzInfo.BaseUtcOffset.TotalHours
    }

    $script:un['autoinstall']['timezone'] = $tzname
}

function Set-UnattendedCloudInitUserLocale
{
    param (
        [Parameter(Mandatory = $true)]
        [string]$UserLocale
    )

    try
    {
        $ci = [cultureinfo]::new($UserLocale)
    }
    catch
    {
        Write-PSFMessage -Message "Could not determine culture from $UserLocale. Assuming en_us"
        $script:un['autoinstall']['locale'] = 'en_US.UTF-8'
        $script:un['autoinstall']['keyboard'] = @{
            layout = 'us'
        }
        return
    }

    $weirdLinuxCultureName = if ($ci.IsNeutralCulture) { $ci.TwoLetterISOLanguageName } else {$ci.Name -split '-' | Select-Object -Last 1}
    $script:un['autoinstall']['locale'] = "$($ci.IetfLanguageTag -replace '-','_').UTF-8"
    $script:un['autoinstall']['keyboard'] = @{
        layout = $weirdLinuxCultureName.ToLower()
    }
}


function Set-UnattendedCloudInitWorkgroup
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$WorkgroupName
	)
    
	$script:un['autoinstall']['late-commands'] += "sed -i 's|[#]*workgroup = WORKGROUP|workgroup = {0}|g' /etc/samba/smb.conf" -f $WorkgroupName
}


function Write-UnattendedCloudInitFile
{
    param
    (
        [string]
        $Content,

        [string]
        $DestinationPath,

        [switch]
        $Append
    )
   
    $script:un['autoinstall']['user-data']['write_files'] += @{
        append  = $Append.IsPresent
        path    = $DestinationPath
        content = "{0}`n" -f $Content
    }
}


function Add-UnattendedKickstartNetworkAdapter {
    param (
        [string]$Interfacename,

        [AutomatedLab.IPNetwork[]]$IpAddresses,

        [AutomatedLab.IPAddress[]]$Gateways,

        [AutomatedLab.IPAddress[]]$DnsServers
    )

    $linuxInterfaceName = ($Interfacename -replace '-', ':').ToLower()
    $adapterAddress = $IpAddresses | Select-Object -First 1

    if (-not $adapterAddress) {
        $configurationItem = "network --bootproto=dhcp --device={0}" -f $linuxInterfaceName
    }
    else {
        $configurationItem = "network --bootproto=static --device={0} --ip={1} --netmask={2}" -f $linuxInterfaceName, $adapterAddress.IPAddress.AddressAsString, $adapterAddress.Netmask
    }

    if ($Gateways) {
        $configurationItem += ' --gateway={0}' -f ($Gateways.AddressAsString -join ',')
    }

    $configurationItem += if ($DnsServers | Where-Object AddressAsString -ne '0.0.0.0') {
        ' --nameserver={0}' -f ($DnsServers.AddressAsString -join ',')
    }
    else {
        ' --nodns'
    }

    $existingLine = $script:un | Where-Object { $_ -match 'network' }

    if ($existingLine -like '*bootproto*') {
        $index = $script:un.IndexOf($existingLine)
        $null = $existingLine -match '(?<HostName>--hostname=\w+)'
        $script:un[$index] = '{0} {1}' -f $configurationItem, $Matches.HostName
        return
    }

    $script:un.Add($configurationItem)
}


function Add-UnattendedKickstartPreinstallationCommand {
    [CmdletBinding()]
    param ()

    Write-PSFMessage -Message "No Preinstall implemented yet with Kickstart"
}

function Add-UnattendedKickstartRenameNetworkAdapters
{
    [CmdletBinding()]
    param ( )
    Write-PSFMessage -Message 'Method not yet implemented for RHEL/CentOS/Fedora'
}


function Add-UnattendedKickstartSshPublicKey
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [string]
        $PublicKey
    )

    Write-PSFMessage -Message "No unattended ssh key import on kickstart yet, we're using %post%"
}


function Add-UnattendedKickstartSynchronousCommand
{
    param (
        [Parameter(Mandatory)]
        [string]$Command,

        [Parameter(Mandatory)]
        [string]$Description
    )

    Write-PSFMessage -Message "Adding command to %post section to $Description"

    $idx = $script:un.IndexOf('%post')

    if ($idx -eq -1)
    {
        $script:un.Add('%post')
        $idx = $script:un.IndexOf('%post')
    }

    $script:un.Insert($idx + 1, $Command)
}


function Export-UnattendedKickstartFile {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [string]$Version
    )

    $idx = $script:un.IndexOf('%post')

    if ($idx -eq -1) {
        $script:un.Add('%post')
        $idx = $script:un.IndexOf('%post')
    }

    $repoIp = try {
        ([System.Net.Dns]::GetHostByName('packages.microsoft.com').AddressList | Where-Object AddressFamily -eq InterNetwork).IPAddressToString
    }
    catch
    { '104.214.230.139' }

    try {
        $repoContent = (Invoke-RestMethod -Method Get -Uri "https://packages.microsoft.com/config/rhel/$Version/prod.repo" -ErrorAction Stop) -split "`n"
    }
    catch { }

    $pwshRelease = ((Invoke-RestMethod -Uri 'https://api.github.com/repos/PowerShell/PowerShell/releases/latest' -ErrorAction SilentlyContinue).assets | Where-Object Name -match 'rh\.x86_64\.rpm').browser_download_url
    if (-not $pwshRelease) {
        $pwshRelease = 'https://github.com/PowerShell/PowerShell/releases/download/v7.5.2/powershell-7.5.2-1.rh.x86_64.rpm'
    }


    if ($script:un[$idx + 1] -ne '#start') {
        @(
            '#start'
            '. /etc/os-release'
            foreach ($line in $repoContent) {
                if (-not $line) { continue }
                if ($line -like '*gpgcheck*') { $line = 'gpgcheck=0' }
                'echo "{0}" >> /etc/yum.repos.d/microsoft.repo' -f $line
            }
            'echo "{0} packages.microsoft.com" >> /etc/hosts' -f $repoIp
            'yum install -y openssl'
            'yum install -y omi'
            'yum install -y powershell'
            'yum install -y omi-psrp-server'
            'yum list installed "powershell" > /ksPowerShell'
            'yum list installed "omi-psrp-server" > /ksOmi'
            'rm /etc/yum.repos.d/microsoft.repo'
            foreach ($line in $repoContent) {
                if (-not $line) { continue }
                'echo "{0}" >> /etc/yum.repos.d/microsoft.repo' -f $line
            }
            'authselect select sssd with-mkhomedir -f'
            'systemctl restart sssd'
            'echo "Subsystem powershell /usr/bin/pwsh -sshs -NoLogo" >> /etc/ssh/sshd_config'
            'systemctl restart sshd'
            'if (! command -v pwsh >/dev/null 2>&1)'
            'then'
            '    sudo dnf install -y {0} >/dev/null 2>&1' -f $pwshRelease
            '    sudo yum install -y {0} >/dev/null 2>&1' -f $pwshRelease
            'fi'
        ) | ForEach-Object -Process {
            $idx++
            $script:un.Insert($idx, $_)
        }

        # When index of end is greater then index of package end: add %end to EOF
        # else add %end before %packages

        $idxPackage = $script:un.IndexOf('%packages --ignoremissing')
        $idxPost = $script:un.IndexOf('%post')

        $idxEnd = if (-1 -ne $idxPackage -and $idxPost -lt $idxPackage) {
            $idxPackage
        }
        else {
            $script:un.Count
        }

        $script:un.Insert($idxEnd, '%end')
    }

    ($script:un | Out-String) -replace "`r`n", "`n" | Set-Content -Path $Path -Force
}


function Import-UnattendedKickstartContent
{
    param
    (
        [Parameter(Mandatory = $true)]
        [System.Collections.Generic.List[string]]
        $Content
    )
    $script:un = $Content
}


function Set-UnattendedKickstartAdministratorName
{
    param
    (
        $Name
    )

    $script:un.Add("user --name=$Name --groups=wheel --password='%PASSWORD%'")
}


function Set-UnattendedKickstartAdministratorPassword
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$Password
    )

		$Script:un.Add("rootpw '$Password'")
		$Script:un = [System.Collections.Generic.List[string]]($Script:un.Replace('%PASSWORD%', $Password))
}


function Set-UnattendedKickstartAntiMalware
{
    param (
        [Parameter(Mandatory = $true)]
        [bool]$Enabled
    )

    if ($Enabled)
    {
        $Script:un.Add("selinux --enforcing")
    }
    else
    {
        $Script:un.Add("selinux --permissive") # Not a great idea to disable selinux alltogether
    }
}


function Set-UnattendedKickstartAutoLogon
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$DomainName,

		[Parameter(Mandatory = $true)]
		[string]$Username,

		[Parameter(Mandatory = $true)]
		[string]$Password
    )
    Write-PSFMessage -Message "Auto-logon not implemented yet for RHEL/CentOS/Fedora"
}


function Set-UnattendedKickstartComputerName {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ComputerName
    )

    $existingLine = $script:un | Where-Object { $_ -match 'network' }
    if ($existingLine -like '*--hostname*') {
        $index = $script:un.IndexOf($existingLine)
        $script:un[$index] = $existingLine -replace '--hostname=\w+', "--hostname=$ComputerName"
        return
    }

    if ($existingLine) {
        $index = $script:un.IndexOf($existingLine)
        $script:un[$index] = '{0} {1}' -f $existingLine, "--hostname=$ComputerName"
        return
    }

    $script:un.Add("network --hostname=$ComputerName")
}


function Set-UnattendedKickstartDomain {
	param (
		[Parameter(Mandatory = $true)]
		[string]$DomainName,

		[Parameter(Mandatory = $true)]
		[string]$Username,

		[Parameter(Mandatory = $true)]
		[string]$Password,

		[Parameter()]
		[string]$OrganizationalUnit
	)

	$idx = $script:un.IndexOf('%post')

	if ($idx -eq -1) {
		$idx = $script:un.Count
	}

	if ($OrganizationalUnit) {
		$script:un.Insert($idx , ("realm join --computer-ou='{2}' --one-time-password='{0}' {1}" -f $Password, $DomainName, $OrganizationalUnit))

	}
	else {
		$script:un.Insert($idx , ("realm join --one-time-password='{0}' {1}" -f $Password, $DomainName))
	}

	$existingLine = $script:un | Where-Object { $_ -match 'network' }

	if ($existingLine -like '*--ipv4-dns-search*') {
		$index = $script:un.IndexOf($existingLine)
		$script:un[$index] = $existingLine -replace 'ipv4-dns-search=[\w\.]+', "--ipv4-dns-search=$DomainName"
		return
	}

	if ($existingLine) {
		$index = $script:un.IndexOf($existingLine)
		$script:un[$index] = '{0} {1}' -f $existingLine, "--ipv4-dns-search=$DomainName"
		return
	}

	$script:un.Add($idx , "network --ipv4-dns-search=$DomainName")
}


function Set-UnattendedKickstartFirewallState
{
    param
    (
        [Parameter(Mandatory = $true)]
        [boolean]$State
    )

    if ($State)
    {
        $script:un.Add('firewall --enabled')
    }
    else
    {
        $script:un.Add('firewall --disabled')
    }
}


function Set-UnattendedKickstartIpSettings
{
    param (
        [string]$IpAddress,

        [string]$Gateway,

        [String[]]$DnsServers,

        [string]$DnsDomain
    )

    if (-not $IpAddress)
    {
        $configurationItem = "network --bootproto=dhcp"
    }
    else
    {
        $configurationItem = "network --bootproto=static --ip={0}" -f $IpAddress
    }

    if ($Gateway)
    {
        $configurationItem += ' --gateway={0}' -f $Gateway
    }

    $configurationItem += if ($DnsServers)
    {
        ' --nameserver={0} --ipv4-dns-search={1}' -f ($DnsServers.AddressAsString -join ','), $DnsDomain
    }
    else
    {
        ' --nodns'
    }

     $existingLine = $script:un | Where-Object { $_ -match 'network' }

    if ($existingLine -like '*bootproto*') {
        $index = $script:un.IndexOf($existingLine)
        $null = $existingLine -match '(?<HostName>--hostname=\w+)'
        $script:un[$index] = '{0} {1}' -f $configurationItem, $Matches.HostName
        return
    }

    $script:un.Add($configurationItem)
}


function Set-UnattendedKickstartLocalIntranetSites
{
	param (
		[Parameter(Mandatory = $true)]
		[string[]]$Values
	)

	Write-PSFMessage -Message 'No local intranet sites for RHEL/CentOS/Fedora'
}


function Set-UnattendedKickstartPackage
{
    param
    (
        [string[]]$Package
    )

    if ($Package -like '*Gnome*')
    {
        $script:un.Add('xconfig --startxonboot --defaultdesktop=GNOME')
    }
    elseif ($Package -like '*KDE*')
    {
        Write-PSFMessage -Level Warning -Message 'Adding KDE UI to RHEL/CentOS via kickstart file is not supported. Please configure your UI manually.'
    }

    $script:un.Add('%packages --ignoremissing')
    $script:un.Add('@^server-product-environment')
    $required = @(
        'oddjob'
        'oddjob-mkhomedir'
        'sssd'
        'adcli'
        'krb5-workstation'
        'realmd'
        'samba-common'
        'samba-common-tools'
        'authselect-compat'
        'sshd'
    )

    foreach ($p in $Package)
    {
        if ($p -eq '@^server-product-environment' -or $p -in $required) { continue }

        $null = $script:un.Add($p)

        if ($p -like '*gnome*' -and $script:un -notcontains '@^graphical-server-environment') { $script:un.Add('@^graphical-server-environment')}
    }

    foreach ($p in $required)
    {
        $script:un.Add($p)
    }

    $script:un.Add('%end')
}


function Set-UnattendedKickstartProductKey
{
    param (
        [Parameter(Mandatory = $true)]
        [string]$ProductKey
    )

    Write-PSFMessage -Message 'No product key necessary for RHEL/CentOS/Fedora'
}


function Set-UnattendedKickstartTimeZone
{
    param
    (
        [Parameter(Mandatory = $true)]
        [string]$TimeZone
    )

    $tzInfo = Get-TimeZone -Id $TimeZone -ErrorAction SilentlyContinue

    if (-not $tzInfo) { Get-TimeZone }

    Write-PSFMessage -Message ('Since non-standard timezone names are used, we revert to Etc/GMT{0}' -f $tzInfo.BaseUtcOffset.TotalHours)
    if ($tzInfo.BaseUtcOffset.TotalHours -gt 0)
    {
        $script:un.Add(('timezone Etc/GMT+{0}' -f $tzInfo.BaseUtcOffset.TotalHours))
    }
    elseif ($tzInfo.BaseUtcOffset.TotalHours -eq 0)
    {
        $script:un.Add('timezone Etc/GMT')
    }
    else
    {
        $script:un.Add(('timezone Etc/GMT{0}' -f $tzInfo.BaseUtcOffset.TotalHours))
    }
}


function Set-UnattendedKickstartUserLocale
{
    param (
        [Parameter(Mandatory = $true)]
        [string]$UserLocale
    )

    try
    {
        $ci = [cultureinfo]::new($UserLocale)
    }
    catch
    {
        Write-PSFMessage -Message "Could not determine culture from $UserLocale. Assuming en_us"
        $script:un.Add("keyboard 'us'")
        $script:un.Add('lang en_us')
        return
    }

    $weirdLinuxCultureName = if ($ci.IsNeutralCulture) { $ci.TwoLetterISOLanguageName } else {$ci.Name -split '-' | Select-Object -Last 1}
    $script:un.Add("keyboard '$($weirdLinuxCultureName.ToLower())'")
    $script:un.Add("lang $($ci.IetfLanguageTag -replace '-','_')")
}


function Set-UnattendedKickstartWorkgroup
{
    param
    (
		[Parameter(Mandatory = $true)]
        [string]
        $WorkgroupName
    )

    $script:un.Add(('auth --smbworkgroup={0}' -f $WorkgroupName))
}


function Write-UnattendedKickstartFile
{
    param
    (
        [string]
        $Content,

        [string]
        $DestinationPath,

        [switch]
        $Append
    )
   Write-PSFMessage -Message 'Unattended Kickstart File Not Implemented Yet'
}


function Add-UnattendedYastPreinstallationCommand {
    [CmdletBinding()]
    param ()

    Write-PSFMessage -Message "No Preinstall implemented yet with YAST"
}

function Add-UnattendedYastNetworkAdapter
{
    param (
        [string]$Interfacename,

        [AutomatedLab.IPNetwork[]]$IpAddresses,

        [AutomatedLab.IPAddress[]]$Gateways,

        [AutomatedLab.IPAddress[]]$DnsServers,

        [string]$ConnectionSpecificDNSSuffix,

        [string]$DnsDomain,

        [string]$DNSSuffixSearchOrder
    )

    $networking = $script:un.SelectSingleNode('/un:profile/un:networking', $script:nsm)
    $interfaceList = $script:un.SelectSingleNode('/un:profile/un:networking/un:interfaces', $script:nsm)
    $udevList = $script:un.SelectSingleNode('/un:profile/un:networking/un:net-udev', $script:nsm)
    $dns = $script:un.SelectSingleNode('/un:profile/un:networking/un:dns', $script:nsm)
    $nameServers = $script:un.SelectSingleNode('/un:profile/un:networking/un:dns/un:nameservers', $script:nsm)
    $routes = $script:un.SelectSingleNode('/un:profile/un:networking/un:routing/un:routes', $script:nsm)
    $hostName = $script:un.CreateElement('hostname', $script:nsm.LookupNamespace('un'))
    $null = $dns.AppendChild($hostName)

    if ($DnsDomain)
    {
        $domain = $script:un.CreateElement('domain', $script:nsm.LookupNamespace('un'))
        $domain.InnerText = $DnsDomain
        $null = $dns.AppendChild($domain)
    }

    if ($DnsServers)
    {
        foreach ($ns in $DnsServers)
        {
            $nameserver = $script:un.CreateElement('nameserver', $script:nsm.LookupNamespace('un'))
            $nameserver.InnerText = $ns
            $null = $nameservers.AppendChild($nameserver)
        }

        if ($DNSSuffixSearchOrder)
        {
            $searchlist = $script:un.CreateElement('searchlist', $script:nsm.LookupNamespace('un'))
            $nsAttr = $script:un.CreateAttribute('config', 'type', $script:nsm.LookupNamespace('config'))
            $nsAttr.InnerText = 'list'
            $null = $searchlist.Attributes.Append($nsAttr)

            foreach ($suffix in ($DNSSuffixSearchOrder -split ','))
            {
                $suffixEntry = $script:un.CreateElement('search', $script:nsm.LookupNamespace('un'))
                $suffixEntry.InnerText = $suffix
                $null = $searchlist.AppendChild($suffixEntry)
            }

            $null = $dns.AppendChild($searchlist)
        }
    }

    $null = $networking.AppendChild($dns)

    $interface = 'eth0'
    $lastInterface = $script:un.SelectNodes('/un:profile/un:networking/un:interfaces/un:interface/un:device', $script:nsm).InnerText | Sort-Object | Select-Object -Last 1
    if ($lastInterface) { $interface = 'eth{0}' -f ([int]$lastInterface.Substring($lastInterface.Length - 1, 1) + 1) }

    $interfaceNode = $script:un.CreateElement('interface', $script:nsm.LookupNamespace('un'))
    $bootproto = $script:un.CreateElement('bootproto', $script:nsm.LookupNamespace('un'))
    $bootproto.InnerText = if ($IpAddresses.Count -eq 0) { 'dhcp' } else { 'static' }
    $deviceNode = $script:un.CreateElement('device', $script:nsm.LookupNamespace('un'))
    $deviceNode.InnerText = $interface
    $firewallnode = $script:un.CreateElement('firewall', $script:nsm.LookupNamespace('un'))
    $firewallnode.InnerText = 'no'

    if ($IpAddresses.Count -gt 0)
    {
        $ipaddr = $script:un.CreateElement('ipaddr', $script:nsm.LookupNamespace('un'))
        $netmask = $script:un.CreateElement('netmask', $script:nsm.LookupNamespace('un'))
        $network = $script:un.CreateElement('network', $script:nsm.LookupNamespace('un'))
        $ipaddr.InnerText = $IpAddresses[0].IpAddress.AddressAsString
        $netmask.InnerText = $IpAddresses[0].Netmask.AddressAsString
        $network.InnerText = $IpAddresses[0].Network.AddressAsString
        $null = $interfaceNode.AppendChild($ipaddr)
        $null = $interfaceNode.AppendChild($netmask)
        $null = $interfaceNode.AppendChild($network)
    }
    $startmode = $script:un.CreateElement('startmode', $script:nsm.LookupNamespace('un'))

    $startmode.InnerText = 'auto'

    $null = $interfaceNode.AppendChild($bootproto)
    $null = $interfaceNode.AppendChild($deviceNode)
    $null = $interfaceNode.AppendChild($firewallnode)
    $null = $interfaceNode.AppendChild($startmode)

    if ($IpAddresses.Count -gt 1)
    {
        $aliases = $script:un.CreateElement('aliases', $script:nsm.LookupNamespace('un'))
        $count = 0

        foreach ($additionalAdapter in ($IpAddresses | Select-Object -Skip 1))
        {
            $alias = $script:un.CreateElement("alias$count", $script:nsm.LookupNamespace('un'))
            $ipaddr = $script:un.CreateElement('IPADDR', $script:nsm.LookupNamespace('un'))
            $label = $script:un.CreateElement('LABEL', $script:nsm.LookupNamespace('un'))
            $netmask = $script:un.CreateElement('NETMASK', $script:nsm.LookupNamespace('un'))
            $ipaddr.InnerText = $additionalAdapter.IpAddress.AddressAsString
            $netmask.InnerText = $additionalAdapter.Netmask.AddressAsString
            $label.InnerText = "ip$count"
            $null = $alias.AppendChild($ipaddr)
            $null = $alias.AppendChild($label)
            $null = $alias.AppendChild($netmask)
            $null = $aliases.AppendChild($alias)
            $count++
        }

        $null = $interfaceNode.AppendChild($aliases)
    }

    $null = $interfaceList.AppendChild($interfaceNode)

    $udevRuleNode = $script:un.CreateElement('rule', $script:nsm.LookupNamespace('un'))
    $udevRuleNameNode = $script:un.CreateElement('name', $script:nsm.LookupNamespace('un'))
    $udevRuleNameNode.InnerText = $interface
    $udevRuleRuleNode = $script:un.CreateElement('rule', $script:nsm.LookupNamespace('un'))
    $udevRuleRuleNode.InnerText = 'ATTR{address}' # No joke. They really mean it to be written this way
    $udevRuleValueNode = $script:un.CreateElement('value', $script:nsm.LookupNamespace('un'))
    $udevRuleValueNode.InnerText = ($Interfacename -replace '-', ':').ToUpper()
    $null = $udevRuleNode.AppendChild($udevRuleNameNode)
    $null = $udevRuleNode.AppendChild($udevRuleRuleNode)
    $null = $udevRuleNode.AppendChild($udevRuleValueNode)
    $null = $udevList.AppendChild($udevRuleNode)

    if ($Gateways)
    {
        foreach ($gateway in $Gateways)
        {
            $routeNode = $script:un.CreateElement('route', $script:nsm.LookupNamespace('un'))
            $mapAttr = $script:un.CreateAttribute('t')
            $mapAttr.InnerText = 'map'
            $null = $routeNode.Attributes.Append($mapAttr)
            $destinationNode = $script:un.CreateElement('destination', $script:nsm.LookupNamespace('un'))
            $deviceNode = $script:un.CreateElement('device', $script:nsm.LookupNamespace('un'))
            $gatewayNode = $script:un.CreateElement('gateway', $script:nsm.LookupNamespace('un'))
            $netmask = $script:un.CreateElement('netmask', $script:nsm.LookupNamespace('un'))

            $destinationNode.InnerText = 'default' # should work for both IPV4 and IPV6 routes

            $devicenode.InnerText = $interface
            $gatewayNode.InnerText = $gateway.AddressAsString
            $netmask.InnerText = '-'

            $null = $routeNode.AppendChild($destinationNode)
            $null = $routeNode.AppendChild($devicenode)
            $null = $routeNode.AppendChild($gatewayNode)
            $null = $routeNode.AppendChild($netmask)
            $null = $routes.AppendChild($routeNode)
        }
    }
}


function Add-UnattendedYastRenameNetworkAdapters
{

}

function Add-UnattendedYastSshPublicKey
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [string]
        $PublicKey
    )

    <#
    <authorized_keys config:type="list">
    <listentry>ssh-rsa ...</listentry>
    </authorized_keys>
    #>
    $userNode = $script:un.SelectSingleNode('/un:profile/un:users', $script:nsm)
    foreach ($user in $userNode.ChildNodes)
    {
        if (-not $user.authorized_keys)
        {
            $keyNode = $script:un.CreateElement('authorized_keys', $script:nsm.LookupNamespace('un'))
            $keyNode.SetAttribute('type', $script:nsm.LookupNamespace('config'), 'list')
            $null = $user.AppendChild($keyNode)
        }

        $keyNode = $user.authorized_keys
        if ($keyNode.listentry -contains $PublicKey) { continue }
        $listEntry = $script:un.CreateElement('listentry', $script:nsm.LookupNamespace('un'))
        $listEntry.InnerText = $PublicKey
        $null = $keyNode.AppendChild($listEntry)
    }
}


function Add-UnattendedYastSynchronousCommand {
    param (
        [Parameter(Mandatory)]
        [string]$Command,

        [Parameter(Mandatory)]
        [string]$Description
    )

    # Init Scripts - run after the system is up and running
    $scriptsNode = $script:un.SelectSingleNode('/un:profile/un:scripts/un:init-scripts', $script:nsm)

    # Add new script with GUID as filename (mandatory if more than one script)
    $scriptNode = $script:un.CreateElement('script', $script:nsm.LookupNamespace('un'))
    $mapAttr = $script:un.CreateAttribute('t')
    $mapAttr.InnerText = 'map'
    $null = $scriptNode.Attributes.Append($mapAttr)
    
    $fileNameNode = $script:un.CreateElement('filename', $script:nsm.LookupNamespace('un'))
    $fileNameNode.InnerText = [guid]::NewGuid().ToString()
    $null = $scriptNode.AppendChild($fileNameNode)

    # Add "source" node with CDATA content of $Command
    $sourceNode = $script:un.CreateElement('source', $script:nsm.LookupNamespace('un'))
    $cdata = $script:un.CreateCDataSection($Command)
    $null = $sourceNode.AppendChild($cdata)
    $null = $scriptNode.AppendChild($sourceNode)

    # Append the script node to the scripts node
    $null = $scriptsNode.AppendChild($scriptNode)
}

function Export-UnattendedYastFile
{
    param (
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $script:un.Save($Path)
}

function Import-UnattendedYastContent {
    param
    (
        [Parameter(Mandatory = $true)]
        [xml]
        $Content
    )

    $script:un = $Content
    $script:ns = @{
        xmlns  = "http://www.suse.com/1.0/yast2ns"
        config = "http://www.suse.com/1.0/configns"
    }
    $script:nsm = [System.Xml.XmlNamespaceManager]::new($script:un.NameTable)
    $script:nsm.AddNamespace('un', "http://www.suse.com/1.0/yast2ns")
    $script:nsm.AddNamespace('config', "http://www.suse.com/1.0/configns" )
}


function Set-UnattendedYastAdministratorName
{
    param
    (
        $Name
    )

    $userNode = $script:un.SelectSingleNode('/un:profile/un:users', $script:nsm)

    $user = $script:un.CreateElement('user', $script:nsm.LookupNamespace('un'))
    $username = $script:un.CreateElement('username', $script:nsm.LookupNamespace('un'))
    $pw = $script:un.CreateElement('user_password', $script:nsm.LookupNamespace('un'))
    $encrypted = $script:un.CreateElement('encrypted', $script:nsm.LookupNamespace('un'))
    $boolAttr = $script:un.CreateAttribute('t')
    $boolAttr.InnerText = 'boolean'
    $null = $encrypted.Attributes.Append($boolAttr)

    $encrypted.InnerText = 'false'
    $pw.InnerText = 'none'
    $username.InnerText = $Name

    $null = $user.AppendChild($pw)
    $null = $user.AppendChild($encrypted)
    $null = $user.AppendChild($username)

    $null = $userNode.AppendChild($user)
}

function Set-UnattendedYastAdministratorPassword
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$Password
    )

		$passwordNodes = $script:un.SelectNodes('/un:profile/un:users/un:user/un:user_password', $script:nsm)

		foreach ($node in $passwordNodes)
		{
			$node.InnerText = $Password
		}
}

function Set-UnattendedYastAntiMalware
{
    param (
        [Parameter(Mandatory = $true)]
        [bool]$Enabled
    )
}

function Set-UnattendedYastAutoLogon
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$DomainName,

		[Parameter(Mandatory = $true)]
		[string]$Username,

		[Parameter(Mandatory = $true)]
		[string]$Password
    )

	$logonNode = $script:un.CreateElement('login_settings', $script:nsm.LookupNamespace('un'))
	$autoLogon = $script:un.CreateElement('autologin_user', $script:nsm.LookupNamespace('un'))
	$autologon.InnerText = '{0}\{1}' -f $DomainName, $Username
	$null = $logonNode.AppendChild($autoLogon)
	$null = $script:un.DocumentElement.AppendChild($logonNode)
}

function Set-UnattendedYastComputerName
{
    param (
        [Parameter(Mandatory = $true)]
        [string]$ComputerName
    )
    $component = $script:un.SelectSingleNode('/un:profile/un:networking/un:dns/un:hostname', $script:nsm)
    $component.InnerText = $ComputerName
}

function Set-UnattendedYastDomain {
    param (
        [Parameter(Mandatory = $true)]
        [string]$DomainName,

        [Parameter(Mandatory = $true)]
        [string]$Username,

        [Parameter(Mandatory = $true)]
        [string]$Password,

        [Parameter()]
        [string]$OrganizationalUnit
    )

    $component = $script:un.SelectSingleNode('/un:profile/un:networking/un:dns/un:hostname', $script:nsm)
    $component.InnerText = '{0}.{1}' -f $component.InnerText, $DomainName

    <# SSSD configuration JSON - generated on running OpenSUSE client
    According to what docs I found this is also valid for older editions
    #>
    $sssdHash = @{
        "sssd" = @{
            "conf"    = @{
                "sssd"               = @{
                    "config_file_version" = "2"
                    "services"            = @(
                        "nss",
                        "pam"
                    )
                    "domains"             = @(
                        $DomainName
                    )
                }
                "nss"                = @{}
                "pam"                = @{}
                "domain/contoso.com" = @{
                    "id_provider"       = "ad"
                    "auth_provider"     = "ad"
                    "enumerate"         = "false"
                    "cache_credentials" = "false"
                    "case_sensitive"    = "true"
                }
            }
            "pam"     = $true
            "nss"     = @(
                "passwd"
                "group"
            )
            "enabled" = $true
        }
        "ldap" = @{
            "pam" = $false
            "nss" = @()
        }
        "krb"  = @{
            "conf" = @{
                "include"      = @()
                "libdefaults"  = @{
                    "dns_canonicalize_hostname" = "false"
                    "rdns"                      = "false"
                    dns_lookup_kdc              = "true"
                    "verify_ap_req_nofail"      = "true"
                    "default_ccache_name"       = "KEYRING:persistent:%{uid}"
                    "default_realm"             = $DomainName.ToUpper()
                    "clockskew"                 = "300"
                }
                "realms"       = @{
                    $DomainName.ToUpper() = @{
                        "default_domain" = $DomainName
                        "admin_server"   = $DomainName
                    }
                }
                "domain_realm" = @{
                    ".$DomainName" = $DomainName.ToUpper()
                }
                "logging"      = @{
                    "kdc"          = "FILE:/var/log/krb5/krb5kdc.log"
                    "admin_server" = "FILE:/var/log/krb5/kadmind.log"
                    "default"      = "SYSLOG:NOTICE:DAEMON"
                }
                "appdefaults"  = @{
                    "pam" = @{
                        "ticket_lifetime" = "1d"
                        "renew_lifetime"  = "1d"
                        "forwardable"     = "true"
                        "proxiable"       = "false"
                        "minimum_uid"     = "1"
                    }
                }
            }
            "pam"  = $false
        }
        "aux"  = @{
            "autofs"    = $false
            "nscd"      = $false
            "mkhomedir" = $true
        }
        "ad"   = @{
            "domain"             = $DomainName
            "user"               = $Username
            "ou"                 = $OrganizationalUnit
            "pass"               = $Password
            "overwrite_smb_conf" = $false
            "update_dns"         = $true
            "dnshostname"        = ""
        }
    }

    $authClientNode = $script:un.CreateElement('auth-client', $script:nsm.LookupNamespace('un'))
    $mapAttr = $script:un.CreateAttribute('t')
    $mapAttr.InnerText = 'map'
    $null = $authClientNode.Attributes.Append($mapAttr)
    $sssdConf = $script:un.CreateElement('conf_json', $script:nsm.LookupNamespace('un'))
    $sssdConf.InnerText = $sssdHash | ConvertTo-Json -Depth 42 -Compress
    $null = $authClientNode.AppendChild($sssdConf)
    $null = $script:un.DocumentElement.AppendChild($authClientNode)
}

function Set-UnattendedYastFirewallState
{
	param (
		[Parameter(Mandatory = $true)]
		[boolean]$State
		)

		$fwState = $script:un.SelectSingleNode('/un:profile/un:firewall/un:enable_firewall', $script:nsm)
		$fwState.InnerText = $State.ToString().ToLower()
}

function Set-UnattendedYastIpSettings
{
	param (
		[string]$IpAddress,

		[string]$Gateway,

		[String[]]$DnsServers,

        [string]$DnsDomain
    )

}

function Set-UnattendedYastLocalIntranetSites
{
	param (
		[Parameter(Mandatory = $true)]
		[string[]]$Values
	)
}

function Set-UnattendedYastPackage {
    param
    (
        [string[]]$Package
    )

    $packagesNode = $script:un.SelectSingleNode('/un:profile/un:software/un:packages', $script:nsm)
    $patternsNode = $script:un.SelectSingleNode('/un:profile/un:software/un:patterns', $script:nsm)

    foreach ($p in $Package) {
        if ($p -replace '^pattern_', '' -in $patternsNode.ChildNodes.InnerText -or $p -in $packagesNode.ChildNodes.InnerText) {
            Write-Verbose "Package or pattern '$p' already exists in the unattended Yast configuration."
            continue
        }

        if ($p -match "^pattern_") {
            $patternNode = $script:un.CreateElement('pattern', $script:nsm.LookupNamespace('un'))
            $patternNode.InnerText = $p -replace '^pattern_'
            $null = $patternsNode.AppendChild($patternNode)
        }
        else {
            $packageNode = $script:un.CreateElement('package', $script:nsm.LookupNamespace('un'))
            $packageNode.InnerText = $p
            $null = $packagesNode.AppendChild($packageNode)
        }
    }
}


function Set-UnattendedYastProductKey
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$ProductKey
    )
}

function Set-UnattendedYastTimeZone
{
    param
    (
        [Parameter(Mandatory = $true)]
        [string]$TimeZone
    )

    $tzInfo = Get-TimeZone -Id $TimeZone -ErrorAction SilentlyContinue

    if (-not $tzInfo) { Get-TimeZone }

    Write-PSFMessage -Message ('Since non-standard timezone names are used, we revert to Etc/GMT{0}' -f $tzInfo.BaseUtcOffset.TotalHours)

    $timeNode = $script:un.SelectSingleNode('/un:profile/un:timezone/un:timezone', $script:nsm)

    $timeNode.InnerText = if ($tzInfo.BaseUtcOffset.TotalHours -gt 0)
    {
        'Etc/GMT+{0}' -f $tzInfo.BaseUtcOffset.TotalHours
    }
    elseif ($tzInfo.BaseUtcOffset.TotalHours -eq 0)
    {
        'Etc/GMT'
    }
    else
    {
        'Etc/GMT{0}' -f $tzInfo.BaseUtcOffset.TotalHours
    }
}


function Set-UnattendedYastUserLocale
{
    param (
        [Parameter(Mandatory = $true)]
        [string]$UserLocale
    )

    $language = $script:un.SelectSingleNode('/un:profile/un:language', $script:nsm)
    $languageNode = $script:un.SelectSingleNode('/un:profile/un:language/un:language', $script:nsm)
    $keyboard = $script:un.SelectSingleNode('/un:profile/un:keyboard/un:keymap', $script:nsm)

    try
    {
        $ci = [cultureinfo]::new($UserLocale)
    }
    catch
    {
        $ci = [cultureinfo]::new('en-us')
    }

    # Primary language
    $languageNode.InnerText = $ci.IetfLanguageTag -replace '-', '_'

    # Secondary language
    if ($ci.Name -ne 'en-US')
    {
        $languagesNode = $script:un.CreateElement('languages', $script:nsm.LookupNamespace('un'))
        $languagesNode.InnerText = 'en-us'
        $null = $language.AppendChild($languagesNode)
    }

    $keyMapName = '{0}-{1}' -f ($ci.EnglishName -split " ")[0].Trim().ToLower(), ($ci.Name -split '-')[-1].ToLower()
    $keyboard.InnerText = $keyMapName
}

function Set-UnattendedYastWorkgroup
{
    param
    (
		[Parameter(Mandatory = $true)]
        [string]
        $WorkgroupName
    )

    $smbClientNode = $script:un.CreateElement('samba-client', $script:nsm.LookupNamespace('un'))
	$boolAttrib = $script:un.CreateAttribute('config','type', $script:nsm.LookupNamespace('config'))
    $boolAttrib.InnerText = 'boolean'
	$disableDhcp = $script:un.CreateElement('disable_dhcp_hostname', $script:nsm.LookupNamespace('un'))
	$globalNode = $script:un.CreateElement('global', $script:nsm.LookupNamespace('un'))
	$securityNode = $script:un.CreateElement('security', $script:nsm.LookupNamespace('un'))
	$shellNode = $script:un.CreateElement('template_shell', $script:nsm.LookupNamespace('un'))
	$guestNode = $script:un.CreateElement('usershare_allow_guests', $script:nsm.LookupNamespace('un'))
	$domainNode = $script:un.CreateElement('workgroup', $script:nsm.LookupNamespace('un'))
	$homedirNode = $script:un.CreateElement('mkhomedir', $script:nsm.LookupNamespace('un'))
	$winbindNode = $script:un.CreateElement('winbind', $script:nsm.LookupNamespace('un'))

	$null = $disableDhcp.Attributes.Append($boolAttrib)
	$null = $homedirNode.Attributes.Append($boolAttrib)
	$null = $winbindNode.Attributes.Append($boolAttrib)

	$disableDhcp.InnerText = 'true'
	$securityNode.InnerText = 'domain'
	$shellNode.InnerText = '/bin/bash'
	$guestNode.InnerText = 'no'
	$domainNode.InnerText = $DomainName
	$homedirNode.InnerText = 'true'
	$winbindNode.InnerText = 'true'

	$null = $globalNode.AppendChild($securityNode)
	$null = $globalNode.AppendChild($shellNode)
	$null = $globalNode.AppendChild($guestNode)
    $null = $globalNode.AppendChild($domainNode)
    $null = $smbClientNode.AppendChild($disableDhcp)
	$null = $smbClientNode.AppendChild($globalNode)
	$null = $smbClientNode.AppendChild($homedirNode)
	$null = $smbClientNode.AppendChild($winbindNode)

	$null = $script:un.DocumentElement.AppendChild($smbClientNode)
}

function Write-UnattendedYastFile
{
    param
    (
        [string]
        $Content,

        [string]
        $DestinationPath,

        [switch]
        $Append
    )
   Write-PSFMessage -Message 'Unattended SuSE File Not Implemented Yet'
}


function Add-UnattendedNetworkAdapter
{
	[CmdletBinding(DefaultParameterSetName = 'Windows')]
    param (
        [Parameter(ParameterSetName='Windows')]
        [Parameter(ParameterSetName='Kickstart')]
        [Parameter(ParameterSetName='Yast')]
        [Parameter(ParameterSetName='CloudInit')]
        [string]$Interfacename,

        [Parameter(ParameterSetName='Windows')]
        [Parameter(ParameterSetName='Kickstart')]
        [Parameter(ParameterSetName='Yast')]
        [Parameter(ParameterSetName='CloudInit')]
        [AutomatedLab.IPNetwork[]]$IpAddresses,

        [Parameter(ParameterSetName='Windows')]
        [Parameter(ParameterSetName='Kickstart')]
        [Parameter(ParameterSetName='Yast')]
        [Parameter(ParameterSetName='CloudInit')]
        [AutomatedLab.IPAddress[]]$Gateways,

        [Parameter(ParameterSetName='Windows')]
        [Parameter(ParameterSetName='Kickstart')]
        [Parameter(ParameterSetName='Yast')]
        [Parameter(ParameterSetName='CloudInit')]
        [AutomatedLab.IPAddress[]]$DnsServers,

        [Parameter(ParameterSetName='Windows')]
        [Parameter(ParameterSetName='Kickstart')]
        [Parameter(ParameterSetName='Yast')]
        [Parameter(ParameterSetName='CloudInit')]
        [string]$ConnectionSpecificDNSSuffix,

        [Parameter(ParameterSetName='Windows')]
        [Parameter(ParameterSetName='Kickstart')]
        [Parameter(ParameterSetName='Yast')]
        [Parameter(ParameterSetName='CloudInit')]
        [string]$DnsDomain,

        [Parameter(ParameterSetName='Windows')]
        [Parameter(ParameterSetName='Kickstart')]
        [Parameter(ParameterSetName='Yast')]
        [Parameter(ParameterSetName='CloudInit')]
        [string]$UseDomainNameDevolution,

        [Parameter(ParameterSetName='Windows')]
        [Parameter(ParameterSetName='Kickstart')]
        [Parameter(ParameterSetName='Yast')]
        [Parameter(ParameterSetName='CloudInit')]
        [string]$DNSSuffixSearchOrder,

        [Parameter(ParameterSetName='Windows')]
        [Parameter(ParameterSetName='Kickstart')]
        [Parameter(ParameterSetName='Yast')]
        [Parameter(ParameterSetName='CloudInit')]
        [string]$EnableAdapterDomainNameRegistration,

        [Parameter(ParameterSetName='Windows')]
        [Parameter(ParameterSetName='Kickstart')]
        [Parameter(ParameterSetName='Yast')]
        [Parameter(ParameterSetName='CloudInit')]
        [string]$DisableDynamicUpdate,

        [Parameter(ParameterSetName='Windows')]
        [Parameter(ParameterSetName='Kickstart')]
        [Parameter(ParameterSetName='Yast')]
        [Parameter(ParameterSetName='CloudInit')]
        [string]$NetbiosOptions,

        [Parameter(ParameterSetName='Kickstart')]
        [switch]
        $IsKickstart,

        [Parameter(ParameterSetName='Yast')]
        [switch]
        $IsAutoYast,

        [Parameter(ParameterSetName='CloudInit')]
        [switch]
        $IsCloudInit
    )

	if (-not $script:un)
	{
		Write-Error 'No unattended file imported. Please use Import-UnattendedFile first'
		return
	}

    $command = Get-Command -Name $PSCmdlet.MyInvocation.MyCommand.Name.Replace('Unattended', "Unattended$($PSCmdlet.ParameterSetName)")
    $parameters = Sync-Parameter $command -Parameters $PSBoundParameters
    & $command @parameters
}

function Add-UnattendedPreinstallationCommand
{
	[CmdletBinding(DefaultParameterSetName = 'Windows')]
    param (
        [Parameter(ParameterSetName='Windows', Mandatory = $true)]
        [Parameter(ParameterSetName='Kickstart', Mandatory = $true)]
        [Parameter(ParameterSetName='Yast', Mandatory = $true)]
        [Parameter(ParameterSetName='CloudInit', Mandatory = $true)]
        [string]$Command,

        [Parameter(ParameterSetName='Windows', Mandatory = $true)]
        [Parameter(ParameterSetName='Kickstart', Mandatory = $true)]
        [Parameter(ParameterSetName='Yast', Mandatory = $true)]
        [Parameter(ParameterSetName='CloudInit', Mandatory = $true)]
        [string]$Description,

        [Parameter(ParameterSetName='Kickstart')]
        [switch]
        $IsKickstart,

        [Parameter(ParameterSetName='Yast')]
        [switch]
        $IsAutoYast,

        [Parameter(ParameterSetName='CloudInit')]
        [switch]
        $IsCloudInit
    )

	if (-not $script:un)
	{
		Write-Error 'No unattended file imported. Please use Import-UnattendedFile first'
		return
	}

    $commandObject = Get-Command -Name $PSCmdlet.MyInvocation.MyCommand.Name.Replace('Unattended', "Unattended$($PSCmdlet.ParameterSetName)")
    $parameters = Sync-Parameter $commandObject -Parameters $PSBoundParameters
    & $commandObject @parameters
}


function Add-UnattendedRenameNetworkAdapters
{
	[CmdletBinding(DefaultParameterSetName = 'Windows')]
    param
    (
        [Parameter(ParameterSetName='Kickstart')]
        [switch]
        $IsKickstart,

        [Parameter(ParameterSetName='Yast')]
        [switch]
        $IsAutoYast,

        [Parameter(ParameterSetName='CloudInit')]
        [switch]
        $IsCloudInit
    )

	if (-not $script:un)
	{
		Write-Error 'No unattended file imported. Please use Import-UnattendedFile first'
		return
	}

    $command = Get-Command -Name $PSCmdlet.MyInvocation.MyCommand.Name.Replace('Unattended', "Unattended$($PSCmdlet.ParameterSetName)")
    & $command
}

function Add-UnattendedSshPublicKey
{
    [CmdletBinding(DefaultParameterSetName = 'Windows')]
    param (
        [Parameter(ParameterSetName = 'Windows', Mandatory = $true)]
        [Parameter(ParameterSetName = 'Kickstart', Mandatory = $true)]
        [Parameter(ParameterSetName = 'Yast', Mandatory = $true)]
        [Parameter(ParameterSetName = 'CloudInit', Mandatory = $true)]
        [string]$PublicKey,

        [Parameter(ParameterSetName = 'Kickstart')]
        [switch]
        $IsKickstart,

        [Parameter(ParameterSetName = 'Yast')]
        [switch]
        $IsAutoYast,

        [Parameter(ParameterSetName = 'CloudInit')]
        [switch]
        $IsCloudInit
    )

    if (-not $script:un)
    {
        Write-Error 'No unattended file imported. Please use Import-UnattendedFile first'
        return
    }

    $command = Get-Command -Name $PSCmdlet.MyInvocation.MyCommand.Name.Replace('Unattended', "Unattended$($PSCmdlet.ParameterSetName)")
    $parameters = Sync-Parameter $command -Parameters $PSBoundParameters
    & $command @parameters
}

function Add-UnattendedSynchronousCommand
{
	[CmdletBinding(DefaultParameterSetName = 'Windows')]
    param (
        [Parameter(ParameterSetName='Windows', Mandatory = $true)]
        [Parameter(ParameterSetName='Kickstart', Mandatory = $true)]
        [Parameter(ParameterSetName='Yast', Mandatory = $true)]
        [Parameter(ParameterSetName='CloudInit', Mandatory = $true)]
        [string]$Command,

        [Parameter(ParameterSetName='Windows', Mandatory = $true)]
        [Parameter(ParameterSetName='Kickstart', Mandatory = $true)]
        [Parameter(ParameterSetName='Yast', Mandatory = $true)]
        [Parameter(ParameterSetName='CloudInit', Mandatory = $true)]
        [string]$Description,

        [Parameter(ParameterSetName='Kickstart')]
        [switch]
        $IsKickstart,

        [Parameter(ParameterSetName='Yast')]
        [switch]
        $IsAutoYast,

        [Parameter(ParameterSetName='CloudInit')]
        [switch]
        $IsCloudInit
    )

	if (-not $script:un)
	{
		Write-Error 'No unattended file imported. Please use Import-UnattendedFile first'
		return
	}

    $commandObject = Get-Command -Name $PSCmdlet.MyInvocation.MyCommand.Name.Replace('Unattended', "Unattended$($PSCmdlet.ParameterSetName)")
    $parameters = Sync-Parameter $commandObject -Parameters $PSBoundParameters
    & $commandObject @parameters
}


function Export-UnattendedFile
{
	[CmdletBinding(DefaultParameterSetName = 'Windows')]
    param (
        [Parameter(ParameterSetName='Windows', Mandatory = $true)]
        [Parameter(ParameterSetName='Kickstart', Mandatory = $true)]
        [Parameter(ParameterSetName='Yast', Mandatory = $true)]
        [Parameter(ParameterSetName='CloudInit', Mandatory = $true)]
        [string]$Path,

        [Parameter(ParameterSetName='Kickstart')]
        [string]
        $Version,

        [Parameter(ParameterSetName='Kickstart')]
        [switch]
        $IsKickstart,

        [Parameter(ParameterSetName='Yast')]
        [switch]
        $IsAutoYast,

        [Parameter(ParameterSetName='CloudInit')]
        [switch]
        $IsCloudInit
    )

	if (-not $script:un)
	{
		Write-Error 'No unattended file imported. Please use Import-UnattendedFile first'
		return
	}

    $command = Get-Command -Name $PSCmdlet.MyInvocation.MyCommand.Name.Replace('Unattended', "Unattended$($PSCmdlet.ParameterSetName)")
    $parameters = Sync-Parameter $command -Parameters $PSBoundParameters
    & $command @parameters
}

function Get-UnattendedContent
{
	[CmdletBinding()]
	param ()

	return $script:un
}

function Import-UnattendedContent
{
    [CmdletBinding(DefaultParameterSetName = 'Windows')]
    param (
        [Parameter(ParameterSetName = 'Windows', Mandatory = $true)]
        [Parameter(ParameterSetName = 'Kickstart', Mandatory = $true)]
        [Parameter(ParameterSetName = 'Yast', Mandatory = $true)]
        [Parameter(ParameterSetName = 'CloudInit', Mandatory = $true)]
        [string[]]
        $Content,

        [Parameter(ParameterSetName = 'Kickstart')]
        [switch]
        $IsKickstart,

        [Parameter(ParameterSetName = 'Yast')]
        [switch]
        $IsAutoYast,

        [Parameter(ParameterSetName = 'CloudInit')]
        [switch]
        $IsCloudInit
    )

    $command = Get-Command -Name $PSCmdlet.MyInvocation.MyCommand.Name.Replace('Unattended', "Unattended$($PSCmdlet.ParameterSetName)")
    $parameters = Sync-Parameter $command -Parameters $PSBoundParameters
    & $command @parameters
}


function Import-UnattendedFile
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$Path
	)

	$script:un = [xml](Get-Content -Path $Path)
	$script:ns = @{ un = 'urn:schemas-microsoft-com:unattend' }
	$Script:wcmNamespaceUrl = 'http://schemas.microsoft.com/WMIConfig/2002/State'
}


function Set-UnattendedAdministratorName
{
    [CmdletBinding(DefaultParameterSetName = 'Windows')]
    param (
        [Parameter(ParameterSetName = 'Windows', Mandatory = $true)]
        [Parameter(ParameterSetName = 'Kickstart', Mandatory = $true)]
        [Parameter(ParameterSetName = 'Yast', Mandatory = $true)]
        [Parameter(ParameterSetName = 'CloudInit', Mandatory = $true)]
        [string]$Name,

        [Parameter(ParameterSetName = 'Kickstart')]
        [switch]
        $IsKickstart,

        [Parameter(ParameterSetName = 'Yast')]
        [switch]
        $IsAutoYast,

        [Parameter(ParameterSetName = 'CloudInit')]
        [switch]
        $IsCloudInit
    )

    if (-not $script:un)
    {
        Write-Error 'No unattended file imported. Please use Import-UnattendedFile first'
        return
    }

    $command = Get-Command -Name $PSCmdlet.MyInvocation.MyCommand.Name.Replace('Unattended', "Unattended$($PSCmdlet.ParameterSetName)")
    $parameters = Sync-Parameter $command -Parameters $PSBoundParameters
    & $command @parameters
}

function Set-UnattendedAdministratorPassword
{
	[CmdletBinding(DefaultParameterSetName = 'Windows')]
	param (
		[Parameter(ParameterSetName = 'Windows', Mandatory = $true)]
		[Parameter(ParameterSetName = 'Kickstart', Mandatory = $true)]
		[Parameter(ParameterSetName = 'Yast', Mandatory = $true)]
		[Parameter(ParameterSetName = 'CloudInit', Mandatory = $true)]
		[string]$Password,

		[Parameter(ParameterSetName = 'Kickstart')]
		[switch]
		$IsKickstart,

		[Parameter(ParameterSetName = 'Yast')]
		[switch]
		$IsAutoYast,

		[Parameter(ParameterSetName = 'CloudInit')]
		[switch]
		$IsCloudInit
	)

	if (-not $script:un)
	{
		Write-Error 'No unattended file imported. Please use Import-UnattendedFile first'
		return
	}

	$command = Get-Command -Name $PSCmdlet.MyInvocation.MyCommand.Name.Replace('Unattended', "Unattended$($PSCmdlet.ParameterSetName)")
	$parameters = Sync-Parameter $command -Parameters $PSBoundParameters
	& $command @parameters
}

function Set-UnattendedAntiMalware
{
    [CmdletBinding(DefaultParameterSetName = 'Windows')]
    param (
        [Parameter(ParameterSetName = 'Windows', Mandatory = $true)]
        [Parameter(ParameterSetName = 'Kickstart', Mandatory = $true)]
        [Parameter(ParameterSetName = 'Yast', Mandatory = $true)]
        [Parameter(ParameterSetName = 'CloudInit', Mandatory = $true)]
        [bool]$Enabled,

        [Parameter(ParameterSetName = 'Kickstart')]
        [switch]
        $IsKickstart,

        [Parameter(ParameterSetName = 'Yast')]
        [switch]
        $IsAutoYast,

        [Parameter(ParameterSetName = 'CloudInit')]
        [switch]
        $IsCloudInit
    )

    if (-not $script:un)
    {
        Write-Error 'No unattended file imported. Please use Import-UnattendedFile first'
        return
    }

    $command = Get-Command -Name $PSCmdlet.MyInvocation.MyCommand.Name.Replace('Unattended', "Unattended$($PSCmdlet.ParameterSetName)")
    $parameters = Sync-Parameter $command -Parameters $PSBoundParameters
    & $command @parameters
}

function Set-UnattendedAutoLogon
{
    [CmdletBinding(DefaultParameterSetName = 'Windows')]
    param (
        [Parameter(ParameterSetName = 'Windows', Mandatory = $true)]
        [Parameter(ParameterSetName = 'Kickstart', Mandatory = $true)]
        [Parameter(ParameterSetName = 'Yast', Mandatory = $true)]
        [Parameter(ParameterSetName = 'CloudInit', Mandatory = $true)]
        [string]$DomainName,

        [Parameter(ParameterSetName = 'Windows', Mandatory = $true)]
        [Parameter(ParameterSetName = 'Kickstart', Mandatory = $true)]
        [Parameter(ParameterSetName = 'Yast', Mandatory = $true)]
        [Parameter(ParameterSetName = 'CloudInit', Mandatory = $true)]
        [string]$Username,

        [Parameter(ParameterSetName = 'Windows', Mandatory = $true)]
        [Parameter(ParameterSetName = 'Kickstart', Mandatory = $true)]
        [Parameter(ParameterSetName = 'Yast', Mandatory = $true)]
        [Parameter(ParameterSetName = 'CloudInit', Mandatory = $true)]
        [string]$Password,

        [Parameter(ParameterSetName = 'Kickstart')]
        [switch]
        $IsKickstart,

        [Parameter(ParameterSetName = 'Yast')]
        [switch]
        $IsAutoYast,

        [Parameter(ParameterSetName = 'CloudInit')]
        [switch]
        $IsCloudInit
    )

    if (-not $script:un)
    {
        Write-Error 'No unattended file imported. Please use Import-UnattendedFile first'
        return
    }

    $command = Get-Command -Name $PSCmdlet.MyInvocation.MyCommand.Name.Replace('Unattended', "Unattended$($PSCmdlet.ParameterSetName)")
    $parameters = Sync-Parameter $command -Parameters $PSBoundParameters
    & $command @parameters
}

function Set-UnattendedComputerName
{
    [CmdletBinding(DefaultParameterSetName = 'Windows')]
    param (
        [Parameter(ParameterSetName = 'Windows', Mandatory = $true)]
        [Parameter(ParameterSetName = 'Kickstart', Mandatory = $true)]
        [Parameter(ParameterSetName = 'Yast', Mandatory = $true)]
        [Parameter(ParameterSetName = 'CloudInit', Mandatory = $true)]
        [string]$ComputerName,

        [Parameter(ParameterSetName = 'Kickstart')]
        [switch]
        $IsKickstart,

        [Parameter(ParameterSetName = 'Yast')]
        [switch]
        $IsAutoYast,

        [Parameter(ParameterSetName = 'CloudInit')]
        [switch]
        $IsCloudInit
    )

    if (-not $script:un)
    {
        Write-Error 'No unattended file imported. Please use Import-UnattendedFile first'
        return
    }

    $command = Get-Command -Name $PSCmdlet.MyInvocation.MyCommand.Name.Replace('Unattended', "Unattended$($PSCmdlet.ParameterSetName)")
    $parameters = Sync-Parameter $command -Parameters $PSBoundParameters
    & $command @parameters
}

function Set-UnattendedDomain
{
	[CmdletBinding(DefaultParameterSetName = 'Windows')]
	param (
		[Parameter(ParameterSetName = 'Windows', Mandatory = $true)]
		[Parameter(ParameterSetName = 'Kickstart', Mandatory = $true)]
		[Parameter(ParameterSetName = 'Yast', Mandatory = $true)]
		[Parameter(ParameterSetName = 'CloudInit', Mandatory = $true)]
		[string]$DomainName,

		[Parameter(ParameterSetName = 'Windows', Mandatory = $true)]
		[Parameter(ParameterSetName = 'Kickstart', Mandatory = $true)]
		[Parameter(ParameterSetName = 'Yast', Mandatory = $true)]
		[Parameter(ParameterSetName = 'CloudInit', Mandatory = $true)]
		[string]$Username,

		[Parameter(ParameterSetName = 'Windows', Mandatory = $true)]
		[Parameter(ParameterSetName = 'Kickstart', Mandatory = $true)]
		[Parameter(ParameterSetName = 'Yast', Mandatory = $true)]
		[Parameter(ParameterSetName = 'CloudInit', Mandatory = $true)]
		[string]$Password,

		[Parameter()]
		[string]$OrganizationalUnit,

		[Parameter(ParameterSetName = 'Kickstart')]
		[switch]
		$IsKickstart,

		[Parameter(ParameterSetName = 'Yast')]
		[switch]
		$IsAutoYast,

		[Parameter(ParameterSetName = 'CloudInit')]
		[switch]
		$IsCloudInit
	)

	if (-not $script:un)
	{
		Write-Error 'No unattended file imported. Please use Import-UnattendedFile first'
		return
	}

	$command = Get-Command -Name $PSCmdlet.MyInvocation.MyCommand.Name.Replace('Unattended', "Unattended$($PSCmdlet.ParameterSetName)")
	$parameters = Sync-Parameter $command -Parameters $PSBoundParameters
	& $command @parameters
}


function Set-UnattendedFirewallState
{
    [CmdletBinding(DefaultParameterSetName = 'Windows')]
    param (
        [Parameter(ParameterSetName = 'Windows', Mandatory = $true)]
        [Parameter(ParameterSetName = 'Kickstart', Mandatory = $true)]
        [Parameter(ParameterSetName = 'Yast', Mandatory = $true)]
        [Parameter(ParameterSetName = 'CloudInit', Mandatory = $true)]
        [boolean]$State,

        [Parameter(ParameterSetName = 'Kickstart')]
        [switch]
        $IsKickstart,

        [Parameter(ParameterSetName = 'Yast')]
        [switch]
        $IsAutoYast,

        [Parameter(ParameterSetName = 'CloudInit')]
        [switch]
        $IsCloudInit
    )

    if (-not $script:un)
    {
        Write-Error 'No unattended file imported. Please use Import-UnattendedFile first'
        return
    }

    $command = Get-Command -Name $PSCmdlet.MyInvocation.MyCommand.Name.Replace('Unattended', "Unattended$($PSCmdlet.ParameterSetName)")
    $parameters = Sync-Parameter $command -Parameters $PSBoundParameters
    & $command @parameters
}

function Set-UnattendedIpSettings
{
    [CmdletBinding(DefaultParameterSetName = 'Windows')]
    param (
        [Parameter(ParameterSetName = 'Windows')]
        [Parameter(ParameterSetName = 'Kickstart')]
        [Parameter(ParameterSetName = 'Yast')]
        [Parameter(ParameterSetName = 'CloudInit')]
        [string]$IpAddress,

        [Parameter(ParameterSetName = 'Windows')]
        [Parameter(ParameterSetName = 'Kickstart')]
        [Parameter(ParameterSetName = 'Yast')]
        [Parameter(ParameterSetName = 'CloudInit')]
        [string]$Gateway,

        [Parameter(ParameterSetName = 'Windows')]
        [Parameter(ParameterSetName = 'Kickstart')]
        [Parameter(ParameterSetName = 'Yast')]
        [Parameter(ParameterSetName = 'CloudInit')]
        [String[]]$DnsServers,

        [Parameter(ParameterSetName = 'Windows')]
        [Parameter(ParameterSetName = 'Kickstart')]
        [Parameter(ParameterSetName = 'Yast')]
        [Parameter(ParameterSetName = 'CloudInit')]
        [string]$DnsDomain,

        [Parameter(ParameterSetName = 'Kickstart')]
        [switch]
        $IsKickstart,

        [Parameter(ParameterSetName = 'Yast')]
        [switch]
        $IsAutoYast,

        [Parameter(ParameterSetName = 'CloudInit')]
        [switch]
        $IsCloudInit
    )

    if (-not $script:un)
    {
        Write-Error 'No unattended file imported. Please use Import-UnattendedFile first'
        return
    }

    $command = Get-Command -Name $PSCmdlet.MyInvocation.MyCommand.Name.Replace('Unattended', "Unattended$($PSCmdlet.ParameterSetName)")
    $parameters = Sync-Parameter $command -Parameters $PSBoundParameters
    & $command @parameters
}

function Set-UnattendedLocalIntranetSites
{
	[CmdletBinding(DefaultParameterSetName = 'Windows')]
	param (
        [Parameter(ParameterSetName = 'Windows', Mandatory = $true)]
        [Parameter(ParameterSetName = 'Kickstart', Mandatory = $true)]
        [Parameter(ParameterSetName = 'Yast', Mandatory = $true)]
        [Parameter(ParameterSetName = 'CloudInit', Mandatory = $true)]
		[string[]]$Values,

        [Parameter(ParameterSetName='Kickstart')]
        [switch]
        $IsKickstart,

        [Parameter(ParameterSetName='Yast')]
        [switch]
        $IsAutoYast,

        [Parameter(ParameterSetName='CloudInit')]
        [switch]
        $IsCloudInit
	)

	if (-not $script:un)
	{
		Write-Error 'No unattended file imported. Please use Import-UnattendedFile first'
		return
	}

    $command = Get-Command -Name $PSCmdlet.MyInvocation.MyCommand.Name.Replace('Unattended', "Unattended$($PSCmdlet.ParameterSetName)")
    $parameters = Sync-Parameter $command -Parameters $PSBoundParameters
    & $command @parameters
}

function Set-UnattendedPackage
{
    [CmdletBinding(DefaultParameterSetName = 'Windows')]
    param (
        [Parameter(ParameterSetName = 'Windows', Mandatory = $true)]
        [Parameter(ParameterSetName = 'Kickstart', Mandatory = $true)]
        [Parameter(ParameterSetName = 'Yast', Mandatory = $true)]
        [Parameter(ParameterSetName = 'CloudInit', Mandatory = $true)]
        [string[]]$Package,

        [bool]
        $IsSnap = $false,

        [Parameter(ParameterSetName = 'Kickstart')]
        [switch]
        $IsKickstart,

        [Parameter(ParameterSetName = 'Yast')]
        [switch]
        $IsAutoYast,

        [Parameter(ParameterSetName = 'CloudInit')]
        [switch]
        $IsCloudInit
    )

    if (-not $script:un)
    {
        Write-Error 'No unattended file imported. Please use Import-UnattendedFile first'
        return
    }

    $command = Get-Command -Name $PSCmdlet.MyInvocation.MyCommand.Name.Replace('Unattended', "Unattended$($PSCmdlet.ParameterSetName)")
    $parameters = Sync-Parameter $command -Parameters $PSBoundParameters
    & $command @parameters
}

function Set-UnattendedProductKey
{
    [CmdletBinding(DefaultParameterSetName = 'Windows')]
    param (
        [Parameter(ParameterSetName = 'Windows', Mandatory = $true)]
        [Parameter(ParameterSetName = 'Kickstart', Mandatory = $true)]
        [Parameter(ParameterSetName = 'Yast', Mandatory = $true)]
        [Parameter(ParameterSetName = 'CloudInit', Mandatory = $true)]
        [string]$ProductKey,

        [Parameter(ParameterSetName = 'Kickstart')]
        [switch]
        $IsKickstart,

        [Parameter(ParameterSetName = 'Yast')]
        [switch]
        $IsAutoYast,

        [Parameter(ParameterSetName = 'CloudInit')]
        [switch]
        $IsCloudInit
    )

    if (-not $script:un)
    {
        Write-Error 'No unattended file imported. Please use Import-UnattendedFile first'
        return
    }

    $command = Get-Command -Name $PSCmdlet.MyInvocation.MyCommand.Name.Replace('Unattended', "Unattended$($PSCmdlet.ParameterSetName)")
    $parameters = Sync-Parameter $command -Parameters $PSBoundParameters
    & $command @parameters
}

function Set-UnattendedTimeZone
{
    [CmdletBinding(DefaultParameterSetName = 'Windows')]
    param
    (
        [Parameter(ParameterSetName = 'Windows', Mandatory = $true)]
        [Parameter(ParameterSetName = 'Kickstart', Mandatory = $true)]
        [Parameter(ParameterSetName = 'Yast', Mandatory = $true)]
        [Parameter(ParameterSetName = 'CloudInit', Mandatory = $true)]
        [string]$TimeZone,

        [Parameter(ParameterSetName = 'Kickstart')]
        [switch]
        $IsKickstart,

        [Parameter(ParameterSetName = 'Yast')]
        [switch]
        $IsAutoYast,

        [Parameter(ParameterSetName = 'CloudInit')]
        [switch]
        $IsCloudInit
    )

    if (-not $script:un)
    {
        Write-Error 'No unattended file imported. Please use Import-UnattendedFile first'
        return
    }

    $command = Get-Command -Name $PSCmdlet.MyInvocation.MyCommand.Name.Replace('Unattended', "Unattended$($PSCmdlet.ParameterSetName)")
    $parameters = Sync-Parameter $command -Parameters $PSBoundParameters
    & $command @parameters
}

function Set-UnattendedUserLocale
{
    [CmdletBinding(DefaultParameterSetName = 'Windows')]
    param (
        [Parameter(ParameterSetName = 'Windows', Mandatory = $true)]
        [Parameter(ParameterSetName = 'Kickstart', Mandatory = $true)]
        [Parameter(ParameterSetName = 'Yast', Mandatory = $true)]
        [Parameter(ParameterSetName = 'CloudInit', Mandatory = $true)]
        [string]$UserLocale,

        [Parameter(ParameterSetName = 'Kickstart')]
        [switch]
        $IsKickstart,

        [Parameter(ParameterSetName = 'Yast')]
        [switch]
        $IsAutoYast,

        [Parameter(ParameterSetName = 'CloudInit')]
        [switch]
        $IsCloudInit
    )

    if (-not $script:un)
    {
        Write-Error 'No unattended file imported. Please use Import-UnattendedFile first'
        return
    }

    $command = Get-Command -Name $PSCmdlet.MyInvocation.MyCommand.Name.Replace('Unattended', "Unattended$($PSCmdlet.ParameterSetName)")
    $parameters = Sync-Parameter $command -Parameters $PSBoundParameters
    & $command @parameters
}

function Set-UnattendedWorkgroup
{
    [CmdletBinding(DefaultParameterSetName = 'Windows')]
    param (
        [Parameter(ParameterSetName = 'Windows', Mandatory = $true)]
        [Parameter(ParameterSetName = 'Kickstart', Mandatory = $true)]
        [Parameter(ParameterSetName = 'Yast', Mandatory = $true)]
        [Parameter(ParameterSetName = 'CloudInit', Mandatory = $true)]
        [string]$WorkgroupName,

        [Parameter(ParameterSetName = 'Kickstart')]
        [switch]
        $IsKickstart,

        [Parameter(ParameterSetName = 'Yast')]
        [switch]
        $IsAutoYast,

        [Parameter(ParameterSetName = 'CloudInit')]
        [switch]
        $IsCloudInit
    )

    if (-not $script:un)
    {
        Write-Error 'No unattended file imported. Please use Import-UnattendedFile first'
        return
    }

    $command = Get-Command -Name $PSCmdlet.MyInvocation.MyCommand.Name.Replace('Unattended', "Unattended$($PSCmdlet.ParameterSetName)")
    $parameters = Sync-Parameter $command -Parameters $PSBoundParameters
    & $command @parameters
}


function Write-UnattendedFile
{
    param
    (
        [Parameter(Mandatory = $true)]
        [string]
        $Content,

        [Parameter(Mandatory = $true)]
        [string]
        $DestinationPath,

        [switch]
        $Append,

        [Parameter(ParameterSetName = 'Kickstart')]
        [switch]
        $IsKickstart,

        [Parameter(ParameterSetName = 'Yast')]
        [switch]
        $IsAutoYast,

        [Parameter(ParameterSetName = 'CloudInit')]
        [switch]
        $IsCloudInit
    )

    if (-not $script:un)
    {
        Write-Error 'No unattended file imported. Please use Import-UnattendedFile first'
        return
    }

    $command = Get-Command -Name $PSCmdlet.MyInvocation.MyCommand.Name.Replace('Unattended', "Unattended$($PSCmdlet.ParameterSetName)")
    $parameters = Sync-Parameter $command -Parameters $PSBoundParameters
    & $command @parameters
}


<#
$cultures = foreach ($culture in [cultureinfo]::GetCultures('AllCultures')) { 
    $list = New-WinUserLanguageList -Language $culture.Name -ea SilentlyContinue
    if (-not $list.InputMethodTips) { continue }
    '"{0}" = "{1}"' -f $list.LanguageTag, $list.InputMethodTips
}
#>
$languageList = @{
    "af"             = "0436:00000409"
    "af-ZA"          = "0436:00000409"
    "am"             = "045E:{7C472071-36A7-4709-88CC-859513E583A9}{9A4E8FC7-76BF-4A63-980D-FADDADF7E987}"
    "am-ET"          = "045E:{7C472071-36A7-4709-88CC-859513E583A9}{9A4E8FC7-76BF-4A63-980D-FADDADF7E987}"
    "ar-AE"          = "3801:00000401"
    "ar-BH"          = "3C01:00000401"
    "ar-DZ"          = "1401:00020401"
    "ar-EG"          = "0C01:00000401"
    "ar-IQ"          = "0801:00000401"
    "ar-JO"          = "2C01:00000401"
    "ar-KW"          = "3401:00000401"
    "ar-LB"          = "3001:00000401"
    "ar-LY"          = "1001:00000401"
    "ar-MA"          = "1801:00020401"
    "ar-OM"          = "2001:00000401"
    "ar-QA"          = "4001:00000401"
    "ar-SA"          = "0401:00000401"
    "ar-SY"          = "2801:00000401"
    "ar-TN"          = "1C01:00020401"
    "ar-YE"          = "2401:00000401"
    "arn-CL"         = "047A:0000080A"
    "as"             = "044D:0000044D"
    "as-IN"          = "044D:0000044D"
    "az-Cyrl"        = "082C:0000082C"
    "az-Cyrl-AZ"     = "082C:0000082C"
    "az-Latn"        = "042C:0000042C"
    "az-Latn-AZ"     = "042C:0000042C"
    "ba-RU"          = "046D:0000046D"
    "be"             = "0423:00000423"
    "be-BY"          = "0423:00000423"
    "bg"             = "0402:00030402"
    "bg-BG"          = "0402:00030402"
    "bn-BD"          = "0845:00000445"
    "bn-IN"          = "0445:00020445"
    "bo-CN"          = "0451:00010451"
    "br-FR"          = "047E:0000040C"
    "bs"             = "141A:0000041A"
    "bs-Cyrl"        = "201A:0000201A"
    "bs-Cyrl-BA"     = "201A:0000201A"
    "bs-Latn"        = "141A:0000041A"
    "bs-Latn-BA"     = "141A:0000041A"
    "ca"             = "0403:0000040A"
    "ca-ES"          = "0403:0000040A"
    "ca-ES-valencia" = "0803:0000040A"
    "chr-Cher"       = "045C:0000045C"
    "chr-Cher-US"    = "045C:0000045C"
    "co-FR"          = "0483:0000040C"
    "cs"             = "0405:00000405"
    "cs-CZ"          = "0405:00000405"
    "cy"             = "0452:00000452"
    "cy-GB"          = "0452:00000452"
    "da"             = "0406:00000406"
    "da-DK"          = "0406:00000406"
    "de-AT"          = "0C07:00000407"
    "de-CH"          = "0807:00000807"
    "de-DE"          = "0407:00000407"
    "de-LI"          = "1407:00000807"
    "de-LU"          = "1007:00000407"
    "dsb"            = "082E:0002042E"
    "dsb-DE"         = "082E:0002042E"
    "dv"             = "0465:00000465"
    "dv-MV"          = "0465:00000465"
    "dz"             = "0C51:00000C51"
    "dz-BT"          = "0C51:00000C51"
    "el"             = "0408:00000408"
    "el-GR"          = "0408:00000408"
    "en-029"         = "2409:00000409"
    "en-AE"          = "4C09:00000409"
    "en-AU"          = "0C09:00000409"
    "en-BZ"          = "2809:00000409"
    "en-CA"          = "1009:00000409"
    "en-GB"          = "0809:00000809"
    "en-HK"          = "3C09:00000409"
    "en-ID"          = "3809:00000409"
    "en-IE"          = "1809:00001809"
    "en-IN"          = "4009:00004009"
    "en-JM"          = "2009:00000409"
    "en-MY"          = "4409:00000409"
    "en-NZ"          = "1409:00001409"
    "en-PH"          = "3409:00000409"
    "en-SG"          = "4809:00000409"
    "en-TT"          = "2C09:00000409"
    "en-US"          = "0409:00000409"
    "en-ZA"          = "1C09:00000409"
    "en-ZW"          = "3009:00000409"
    "es-419"         = "580A:0000080A"
    "es-AR"          = "2C0A:0000080A"
    "es-BO"          = "400A:0000080A"
    "es-CL"          = "340A:0000080A"
    "es-CO"          = "240A:0000080A"
    "es-CR"          = "140A:0000080A"
    "es-CU"          = "5C0A:0000080A"
    "es-DO"          = "1C0A:0000080A"
    "es-EC"          = "300A:0000080A"
    "es-ES"          = "0C0A:0000040A"
    "es-GT"          = "100A:0000080A"
    "es-HN"          = "480A:0000080A"
    "es-MX"          = "080A:0000080A"
    "es-NI"          = "4C0A:0000080A"
    "es-PA"          = "180A:0000080A"
    "es-PE"          = "280A:0000080A"
    "es-PR"          = "500A:0000080A"
    "es-PY"          = "3C0A:0000080A"
    "es-SV"          = "440A:0000080A"
    "es-US"          = "540A:0000080A"
    "es-UY"          = "380A:0000080A"
    "es-VE"          = "200A:0000080A"
    "et"             = "0425:00000425"
    "et-EE"          = "0425:00000425"
    "eu"             = "042D:0000040A"
    "eu-ES"          = "042D:0000040A"
    "fa"             = "0429:00000429"
    "fa-AF"          = "048C:00050429"
    "fa-IR"          = "0429:00000429"
    "ff-Latn-NG"     = "0467:00000488"
    "ff-Latn-SN"     = "0867:00000488"
    "fi"             = "040B:0000040B"
    "fi-FI"          = "040B:0000040B"
    "fil"            = "0464:00000409"
    "fil-PH"         = "0464:00000409"
    "fo"             = "0438:00000406"
    "fo-FO"          = "0438:00000406"
    "fr-029"         = "1C0C:0000040C"
    "fr-BE"          = "080C:0000080C"
    "fr-CA"          = "0C0C:00001009"
    "fr-CD"          = "240C:0000040C"
    "fr-CH"          = "100C:0000100C"
    "fr-CI"          = "300C:0000040C"
    "fr-CM"          = "2C0C:0000040C"
    "fr-FR"          = "040C:0000040C"
    "fr-HT"          = "3C0C:0000040C"
    "fr-LU"          = "140C:0000100C"
    "fr-MA"          = "380C:0000040C"
    "fr-MC"          = "180C:0000040C"
    "fr-ML"          = "340C:0000040C"
    "fr-RE"          = "200C:0000040C"
    "fr-SN"          = "280C:0000040C"
    "fy"             = "0462:00020409"
    "fy-NL"          = "0462:00020409"
    "gd-GB"          = "0491:00011809"
    "gl"             = "0456:0000040A"
    "gl-ES"          = "0456:0000040A"
    "gn"             = "0474:00000474"
    "gn-PY"          = "0474:00000474"
    "gsw-FR"         = "0484:0000040C"
    "gu"             = "0447:00000447"
    "gu-IN"          = "0447:00000447"
    "ha-Latn"        = "0468:00000468"
    "ha-Latn-NG"     = "0468:00000468"
    "haw"            = "0475:00000475"
    "haw-US"         = "0475:00000475"
    "he"             = "040D:0002040D"
    "he-IL"          = "040D:0002040D"
    "hi"             = "0439:00010439"
    "hi-IN"          = "0439:00010439"
    "hr-BA"          = "101A:0000041A"
    "hr-HR"          = "041A:0000041A"
    "hsb"            = "042E:0002042E"
    "hsb-DE"         = "042E:0002042E"
    "hu"             = "040E:0000040E"
    "hu-HU"          = "040E:0000040E"
    "hy"             = "042B:0002042B"
    "hy-AM"          = "042B:0002042B"
    "id"             = "0421:00000409"
    "id-ID"          = "0421:00000409"
    "ig-NG"          = "0470:00000470"
    "ii-CN"          = "0478:{E429B25A-E5D3-4D1F-9BE3-0C608477E3A1}{409C8376-007B-4357-AE8E-26316EE3FB0D}"
    "is"             = "040F:0000040F"
    "is-IS"          = "040F:0000040F"
    "it-CH"          = "0810:0000100C"
    "it-IT"          = "0410:00000410"
    "iu-Cans"        = "045D:0001045D"
    "iu-Cans-CA"     = "045D:0001045D"
    "iu-Latn"        = "085D:0000085D"
    "iu-Latn-CA"     = "085D:0000085D"
    "ja"             = "0411:{03B5835F-F03C-411B-9CE2-AA23E1171E36}{A76C93D9-5523-4E90-AAFA-4DB112F9AC76}"
    "ja-JP"          = "0411:{03B5835F-F03C-411B-9CE2-AA23E1171E36}{A76C93D9-5523-4E90-AAFA-4DB112F9AC76}"
    "ka"             = "0437:00010437"
    "ka-GE"          = "0437:00010437"
    "kk"             = "043F:0000043F"
    "kk-KZ"          = "043F:0000043F"
    "kl"             = "046F:00000406"
    "kl-GL"          = "046F:00000406"
    "km"             = "0453:00000453"
    "km-KH"          = "0453:00000453"
    "kn"             = "044B:0000044B"
    "kn-IN"          = "044B:0000044B"
    "ko"             = "0412:{A028AE76-01B1-46C2-99C4-ACD9858AE02F}{B5FE1F02-D5F2-4445-9C03-C568F23C99A1}"
    "ko-KR"          = "0412:{A028AE76-01B1-46C2-99C4-ACD9858AE02F}{B5FE1F02-D5F2-4445-9C03-C568F23C99A1}"
    "kok"            = "0457:00000439"
    "kok-IN"         = "0457:00000439"
    "kr-Latn"        = "0471:00000409"
    "kr-Latn-NG"     = "0471:00000409"
    "ks-Deva"        = "0860:00010439"
    "ks-Deva-IN"     = "0860:00010439"
    "ku-Arab"        = "0492:00000492"
    "ku-Arab-IQ"     = "0492:00000492"
    "ky-KG"          = "0440:00000440"
    "la"             = "0476:00000409"
    "la-VA"          = "0476:00000409"
    "lb"             = "046E:0000046E"
    "lb-LU"          = "046E:0000046E"
    "lo"             = "0454:00000454"
    "lo-LA"          = "0454:00000454"
    "lt"             = "0427:00010427"
    "lt-LT"          = "0427:00010427"
    "lv"             = "0426:00020426"
    "lv-LV"          = "0426:00020426"
    "mi"             = "0481:00000481"
    "mi-NZ"          = "0481:00000481"
    "mk"             = "042F:0001042F"
    "mk-MK"          = "042F:0001042F"
    "ml"             = "044C:0000044C"
    "ml-IN"          = "044C:0000044C"
    "mn-Cyrl"        = "0450:00000450"
    "mn-MN"          = "0450:00000450"
    "mn-Mong"        = "0850:00010850"
    "mn-Mong-CN"     = "0850:00010850"
    "mn-Mong-MN"     = "0C50:00010850"
    "mni-Beng"       = "0458:00000445"
    "mni-IN"         = "0458:00000445"
    "moh-CA"         = "047C:00000409"
    "mr"             = "044E:0000044E"
    "mr-IN"          = "044E:0000044E"
    "ms-BN"          = "083E:00000409"
    "ms-MY"          = "043E:00000409"
    "mt"             = "043A:0000043A"
    "mt-MT"          = "043A:0000043A"
    "my"             = "0455:00130C00"
    "my-MM"          = "0455:00130C00"
    "nb"             = "0414:00000414"
    "nb-NO"          = "0414:00000414"
    "ne-IN"          = "0861:00000461"
    "ne-NP"          = "0461:00000461"
    "nl-BE"          = "0813:00000813"
    "nl-NL"          = "0413:00020409"
    "nn"             = "0814:00000414"
    "nn-NO"          = "0814:00000414"
    "no"             = "0414:00000414"
    "nso"            = "046C:0000046C"
    "nso-ZA"         = "046C:0000046C"
    "oc-FR"          = "0482:0000040C"
    "om"             = "0472:00000409"
    "om-ET"          = "0472:00000409"
    "or"             = "0448:00000448"
    "or-IN"          = "0448:00000448"
    "pa"             = "0446:00000446"
    "pa-Arab"        = "0846:00000420"
    "pa-Arab-PK"     = "0846:00000420"
    "pa-Guru"        = "0446:00000446"
    "pa-IN"          = "0446:00000446"
    "pl"             = "0415:00000415"
    "pl-PL"          = "0415:00000415"
    "pt-BR"          = "0416:00000416"
    "pt-PT"          = "0816:00000816"
    "quc-Latn"       = "0486:0000080A"
    "quc-Latn-GT"    = "0486:0000080A"
    "quz-BO"         = "046B:0000080A"
    "quz-EC"         = "086B:0000080A"
    "quz-PE"         = "0C6B:0000080A"
    "rm"             = "0417:00000807"
    "rm-CH"          = "0417:00000807"
    "ro-MD"          = "0818:00010418"
    "ro-RO"          = "0418:00010418"
    "ru"             = "0419:00000419"
    "ru-MD"          = "0819:00000419"
    "ru-RU"          = "0419:00000419"
    "rw"             = "0487:00000409"
    "rw-RW"          = "0487:00000409"
    "sa-IN"          = "044F:00000439"
    "sah-RU"         = "0485:00000485"
    "sd-Arab"        = "0859:00000420"
    "sd-Arab-PK"     = "0859:00000420"
    "sd-Deva"        = "0459:00010439"
    "sd-Deva-IN"     = "0459:00010439"
    "se-FI"          = "0C3B:0001083B"
    "se-NO"          = "043B:0000043B"
    "se-SE"          = "083B:0000083B"
    "si"             = "045B:0000045B"
    "si-LK"          = "045B:0000045B"
    "sk"             = "041B:0000041B"
    "sk-SK"          = "041B:0000041B"
    "sl"             = "0424:00000424"
    "sl-SI"          = "0424:00000424"
    "sma-NO"         = "183B:0000043B"
    "sma-SE"         = "1C3B:0000083B"
    "smj-NO"         = "103B:0000043B"
    "smj-SE"         = "143B:0000083B"
    "smn-FI"         = "243B:0001083B"
    "sms-FI"         = "203B:0001083B"
    "so"             = "0477:00000409"
    "so-SO"          = "0477:00000409"
    "sq"             = "041C:0000041C"
    "sq-AL"          = "041C:0000041C"
    "sr-Cyrl-BA"     = "1C1A:00000C1A"
    "sr-Cyrl-ME"     = "301A:00000C1A"
    "sr-Cyrl-RS"     = "281A:00000C1A"
    "sr-Latn-BA"     = "181A:0000081A"
    "sr-Latn-ME"     = "2C1A:0000081A"
    "sr-Latn-RS"     = "241A:0000081A"
    "st"             = "0430:00000409"
    "st-ZA"          = "0430:00000409"
    "sv-FI"          = "081D:0000041D"
    "sv-SE"          = "041D:0000041D"
    "sw"             = "0441:00000409"
    "sw-KE"          = "0441:00000409"
    "syr-SY"         = "045A:0000045A"
    "ta-IN"          = "0449:00020449"
    "ta-LK"          = "0849:00020449"
    "te"             = "044A:0000044A"
    "te-IN"          = "044A:0000044A"
    "tg-Cyrl"        = "0428:00000428"
    "tg-Cyrl-TJ"     = "0428:00000428"
    "th"             = "041E:0000041E"
    "th-TH"          = "041E:0000041E"
    "ti-ER"          = "0473:{E429B25A-E5D3-4D1F-9BE3-0C608477E3A1}{3CAB88B7-CC3E-46A6-9765-B772AD7761FF}"
    "ti-ET"          = "0473:{E429B25A-E5D3-4D1F-9BE3-0C608477E3A1}{3CAB88B7-CC3E-46A6-9765-B772AD7761FF}"
    "tk-TM"          = "0442:00000442"
    "tn-BW"          = "0832:00000432"
    "tn-ZA"          = "0432:00000432"
    "tr"             = "041F:0000041F"
    "tr-TR"          = "041F:0000041F"
    "ts"             = "0431:00000409"
    "ts-ZA"          = "0431:00000409"
    "tt-RU"          = "0444:00010444"
    "tzm-Arab"       = "045F:00020401"
    "tzm-Arab-MA"    = "045F:00020401"
    "tzm-Latn"       = "085F:0000085F"
    "tzm-Latn-DZ"    = "085F:0000085F"
    "tzm-Tfng"       = "105F:0000105F"
    "tzm-Tfng-MA"    = "105F:0000105F"
    "ug-CN"          = "0480:00010480"
    "uk"             = "0422:00020422"
    "uk-UA"          = "0422:00020422"
    "ur-IN"          = "0820:00000420"
    "ur-PK"          = "0420:00000420"
    "uz-Cyrl"        = "0843:00000843"
    "uz-Cyrl-UZ"     = "0843:00000843"
    "uz-Latn"        = "0443:00000409"
    "uz-Latn-UZ"     = "0443:00000409"
    "ve"             = "0433:00000409"
    "ve-ZA"          = "0433:00000409"
    "vi"             = "042A:{C2CB2CF0-AF47-413E-9780-8BC3A3C16068}{5FB02EC5-0A77-4684-B4FA-DEF8A2195628}"
    "vi-VN"          = "042A:{C2CB2CF0-AF47-413E-9780-8BC3A3C16068}{5FB02EC5-0A77-4684-B4FA-DEF8A2195628}"
    "wo"             = "0488:00000488"
    "wo-SN"          = "0488:00000488"
    "xh"             = "0434:00000409"
    "xh-ZA"          = "0434:00000409"
    "yi"             = "043D:0002040D"
    "yi-001"         = "043D:0002040D"
    "yo-NG"          = "046A:0000046A"
    "zh-CN"          = "0804:{81D4E9C9-1D3B-41BC-9E6C-4B40BF79E35E}{FA550B04-5AD7-411F-A5AC-CA038EC515D7}"
    "zh-Hans-HK"     = "0804:{81D4E9C9-1D3B-41BC-9E6C-4B40BF79E35E}{FA550B04-5AD7-411F-A5AC-CA038EC515D7}"
    "zh-Hans-MO"     = "0804:{81D4E9C9-1D3B-41BC-9E6C-4B40BF79E35E}{FA550B04-5AD7-411F-A5AC-CA038EC515D7}"
    "zh-HK"          = "0404:{531FDEBF-9B4C-4A43-A2AA-960E8FCDC732}{6024B45F-5C54-11D4-B921-0080C882687E}"
    "zh-MO"          = "0404:{531FDEBF-9B4C-4A43-A2AA-960E8FCDC732}{6024B45F-5C54-11D4-B921-0080C882687E}"
    "zh-SG"          = "0804:{81D4E9C9-1D3B-41BC-9E6C-4B40BF79E35E}{FA550B04-5AD7-411F-A5AC-CA038EC515D7}"
    "zh-TW"          = "0404:{B115690A-EA02-48D5-A231-E3578D2FDF80}{B2F9C502-1742-11D4-9790-0080C882687E}"
    "zu"             = "0435:00000409"
    "zu-ZA"          = "0435:00000409"

}
