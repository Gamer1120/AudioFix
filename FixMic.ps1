# Because I was too lazy to switch microphone in Microsoft Teams every time I switched from headset to external mic to speaker mic,
# I wrote this script to do it for me. Depends on https://github.com/frgnca/AudioDeviceCmdlets
# Please note that I rarely code in Powershell. It's probably not pretty code, but it gets the job done.
# Requires local admin, because it needs to tweak the audio devices in registry.

# Source: https://den.dev/blog/powershell-windows-notification/
function Show-Notification {
    [cmdletbinding()]
    Param (
        [string]
        $ToastTitle,
        [string]
        $ToastText
    )

    [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] > $null
    $Template = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent([Windows.UI.Notifications.ToastTemplateType]::ToastText02)

    $RawXml = [xml] $Template.GetXml()
    ($RawXml.toast.visual.binding.text|where {$_.id -eq "1"}).AppendChild($RawXml.CreateTextNode($ToastTitle)) > $null
    ($RawXml.toast.visual.binding.text|where {$_.id -eq "2"}).AppendChild($RawXml.CreateTextNode($ToastText)) > $null

    $SerializedXml = New-Object Windows.Data.Xml.Dom.XmlDocument
    $SerializedXml.LoadXml($RawXml.OuterXml)

    $Toast = [Windows.UI.Notifications.ToastNotification]::new($SerializedXml)
    $Toast.Tag = "PowerShell"
    $Toast.Group = "PowerShell"
    $Toast.ExpirationTime = [DateTimeOffset]::Now.AddMinutes(1)

    $Notifier = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier("PowerShell")
    $Notifier.Show($Toast);
}

$NotificationText = ""

# For some reason, the name of my headset has a number in front of it that increases every
# now and then, and sometimes isn't even there, so I'm using regex matching instead of string compare.

# Priority of microphones. The higher in the list, the more priority the microphone has.
$MicrophonePriorityList = "Microphone \(Yeti Stereo Microphone\)",
                          "Microphone \((\d)*(- )?(Logitech )?G935(\/G933s)? Gaming Headset\)",
                          "Microphone Array \(Realtek\(R\) Audio\)"

# Priority of speakers. The higher in the list, the more priority the speaker has.
$SpeakersPriorityList = "Speakers \((\d)*(- )?(Logitech )?G935(\/G933s)? Gaming Headset\)",
                       "Speakers \(Realtek\(R\) Audio\)"

$CorrectMicrophoneID = ""
:micloop foreach ($mic in $MicrophonePriorityList) {
    if ($CorrectMicrophoneID -eq "") {
        #Write-Output "Checking if the following mic is plugged in: " $mic
        #Write-Output "cmi = " $CurrentMicrophoneIndex
        foreach ($ad in Get-AudioDevice -List) {
            if ($ad.Name -match $mic -and $ad.Type.Equals("Recording")) {
                #We already have the best mic, so exit
                if ($ad.Default){
                    $message += "Best mic is already default: " + $ad.Name + "`n"
                    Write-Output $message
                    $NotificationText += $message
                    break micloop
                } else {
                    $CorrectMicrophoneID = $ad.ID
                    break
                }
                
            }
        }
    }
}
if ($CorrectMicrophoneID -ne "") {
    $message += "Switching to better microphone: " + $ad.Name
    Write-Output $message
    $NotificationText += $message
    Set-AudioDevice -ID $CorrectMicrophoneID >$null 2>&1
}

#Yeah, repeat code. I don't really care
$CorrectSpeakersID = ""
:speakloop foreach ($speak in $SpeakersPriorityList) {
    if ($CorrectSpeakersID -eq "") {
        foreach ($ad in Get-AudioDevice -List) {
            if ($ad.Name -match $speak -and $ad.Type.Equals("Playback")) {
                #We already have the best speaker, so exit
                if ($ad.Default){
                    $message = "Best speaker is already default: " + $ad.Name
                    Write-Output $message
                    $NotificationText += $message
                    break speakloop
                } else {
                    $CorrectSpeakersID = $ad.ID
                    break
                }
                
            }
        }
    }
}
if ($CorrectSpeakersID -ne "") {
    $message = "Switching to better speaker: " + $ad.Name
    Write-Output $message
    $NotificationText += $message
    Set-AudioDevice -ID $CorrectSpeakersID >$null 2>&1
}

Show-Notification -ToastTitle "AudioFix" -ToastText $NotificationText

'lalala'
