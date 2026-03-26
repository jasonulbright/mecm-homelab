function Send-ALIftttNotification
{
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]
        $Activity,

        [Parameter(Mandatory = $true)]
        [System.String]
        $Message
    )

    $lab = Get-Lab -ErrorAction SilentlyContinue

    $key = Get-LabConfigurationItem -Name Notifications.NotificationProviders.Ifttt.Key
    $eventName = Get-LabConfigurationItem -Name Notifications.NotificationProviders.Ifttt.EventName

    $messageBody = @{
        value1 = $lab.Name + " on " + $lab.DefaultVirtualizationEngine
        value2 = $Activity
        value3 = $Message
    }

    try
    {
        $request = Invoke-WebRequest -Method Post -Uri https://maker.ifttt.com/trigger/$($eventName)/with/key/$($key) -ContentType "application/json" -Body ($messageBody | ConvertTo-Json -Compress) -ErrorAction Stop

        if (-not $request.StatusCode -eq 200)
        {
            Write-PSFMessage -Message "Notification to IFTTT could not be sent with event $eventName. Status code was $($request.StatusCode)"
        }
    }
    catch
    {
        Write-PSFMessage -Message "Notification to IFTTT could not be sent with event $eventName."
    }
}


function Send-ALMailNotification
{
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]
        $Activity,

        [Parameter(Mandatory = $true)]
        [System.String]
        $Message
    )

    $lab = Get-Lab

    $body = @"
    Dear recipient,

    Lab $($lab.Name) on $($Lab.DefaultVirtualizationEngine)logged activity "$Activity" with the following message:

    $Message
"@

    $mailParameters = @{
        SmtpServer =  Get-LabConfigurationItem -Name Notifications.NotificationProviders.Mail.SmtpServer
        From = Get-LabConfigurationItem -Name Notifications.NotificationProviders.Mail.From
        CC = Get-LabConfigurationItem -Name Notifications.NotificationProviders.Mail.CC
        To = Get-LabConfigurationItem -Name Notifications.NotificationProviders.Mail.To
        Priority = Get-LabConfigurationItem -Name Notifications.NotificationProviders.Mail.Priority
        Port = Get-LabConfigurationItem -Name Notifications.NotificationProviders.Mail.Port
        Body = $body
        Subject = "AutomatedLab notification: $($lab.Name) $Activity"
    }


    Send-MailMessage @mailParameters
}


function Send-ALToastNotification
{
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]
        $Activity,

        [Parameter(Mandatory = $true)]
        [System.String]
        $Message
    )

    $isFullGui = $true # Client

    if (Get-Item 'HKLM:\software\Microsoft\Windows NT\CurrentVersion\Server\ServerLevels' -ErrorAction SilentlyContinue)
    {
        [bool]$core = [int](Get-ItemProperty 'HKLM:\software\Microsoft\Windows NT\CurrentVersion\Server\ServerLevels' -Name ServerCore -ErrorAction SilentlyContinue).ServerCore
        [bool]$guimgmt = [int](Get-ItemProperty 'HKLM:\software\Microsoft\Windows NT\CurrentVersion\Server\ServerLevels' -Name Server-Gui-Mgmt -ErrorAction SilentlyContinue).'Server-Gui-Mgmt'
        [bool]$guimgmtshell = [int](Get-ItemProperty 'HKLM:\software\Microsoft\Windows NT\CurrentVersion\Server\ServerLevels' -Name Server-Gui-Shell -ErrorAction SilentlyContinue).'Server-Gui-Shell'

        $isFullGui = $core -and $guimgmt -and $guimgmtshell
    }

    if ($PSVersionTable.BuildVersion -lt 6.3 -or -not $isFullGui)
    {
        Write-PSFMessage -Message 'No toasts for OS version < 6.3 or Server Core'
        return
    }

    # Hardcoded toaster from PowerShell - no custom Toast providers after 1709
    $toastProvider = "{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\WindowsPowerShell\v1.0\powershell.exe"
    $imageLocation = Get-LabConfigurationItem -Name Notifications.NotificationProviders.Toast.Image
    $imagePath = "$((Get-LabConfigurationItem -Name LabAppDataRoot))\Assets"
    $imageFilePath = Join-Path $imagePath -ChildPath (Split-Path $imageLocation -Leaf)

    if (-not (Test-Path -Path $imagePath))
    {
        [void](New-Item -ItemType Directory -Path $imagePath)
    }

    if (-not (Test-Path -Path $imageFilePath))
    {
        $file = Get-LabInternetFile -Uri $imageLocation -Path $imagePath -PassThru
    }

    $lab = Get-Lab

    $template = "<?xml version=`"1.0`" encoding=`"utf-8`"?><toast><visual><binding template=`"ToastGeneric`"><text>{2}</text><text>Deployment of {0} on {1}, current status '{2}'. Message {3}.</text><image src=`"{4}`" placement=`"appLogoOverride`" hint-crop=`"circle`" /></binding></visual></toast>" -f `
        $lab.Name, $lab.DefaultVirtualizationEngine, $Activity, $Message, $imageFilePath

    try
    {
        [void]([Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime])
        [void]([Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime])
        [void]([Windows.UI.Notifications.ToastNotification, Windows.UI.Notifications, ContentType = WindowsRuntime])
        $xml = New-Object Windows.Data.Xml.Dom.XmlDocument

        $xml.LoadXml($template)
        $toast = New-Object Windows.UI.Notifications.ToastNotification $xml
        [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($toastProvider).Show($toast)
    }
    catch
    {
        Write-PSFMessage "Error sending toast notification: $($_.Exception.Message)"
    }
}


function Send-ALVoiceNotification
{
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]
        $Activity,

        [Parameter(Mandatory = $true)]
        [System.String]
        $Message
    )

    $lab = Get-Lab
    $culture = Get-LabConfigurationItem -Name Notifications.NotificationProviders.Voice.Culture
    $gender = Get-LabConfigurationItem -Name Notifications.NotificationProviders.Voice.Gender

    try
    {
        Add-Type -AssemblyName System.Speech -ErrorAction Stop
    }
    catch
    {
        return
    }

    $synth = New-Object System.Speech.Synthesis.SpeechSynthesizer
    try
    {
        $synth.SelectVoiceByHints($gender, 30, $null, $culture)
    }
    catch {return}

    if (-not $synth.Voice)
    {
        Write-PSFMessage -Level Warning -Message ('No voice installed for culture {0} and gender {1}' -f $culture, $gender)
        return;
    }
    $synth.SetOutputToDefaultAudioDevice()

    $text = "
        Hi {4}!
        AutomatedLab has a new message for you!
        Deployment of {0} on {1} entered status {2}. Message {3}.
        Live long and prosper.
        " -f $lab.Name, $lab.DefaultVirtualizationEngine, $Activity, $Message, $env:USERNAME
    $synth.Speak($Text)
    $synth.Dispose()
}

function Send-ALNotification
{
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]
        $Activity,

        [Parameter(Mandatory = $true)]
        [System.String]
        $Message,

        [ValidateSet('Toast','Ifttt','Mail','Voice')]
        [string[]]
        $Provider
    )

    begin
    {
        $lab = Get-Lab -ErrorAction SilentlyContinue
        if (-not $lab)
        {
            Write-PSFMessage -Message "No lab data available. Skipping notification."
        }
    }

    process
    {
        if (-not $lab)
        {
            return
        }

        foreach ($selectedProvider in $Provider)
        {
            $functionName = "Send-AL$($selectedProvider)Notification"
            Write-PSFMessage $functionName

            &$functionName -Activity $Activity -Message $Message
        }
    }

}
