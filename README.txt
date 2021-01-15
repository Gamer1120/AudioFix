Automatically switch to the best audio device by pressing CTRL+SHIFT+ALT+\.

Requires https://github.com/frgnca/AudioDeviceCmdlets
Put the files in C:\scripts\
Edit the regexes in FixMic.ps1 to match your audio devices
Install AutoHotkey
Set up a scheduled task to run the AutoHotkey script as admin when you login: schtasks /Create /RU "[Domain][Username]" /SC ONLOGON /TN "Run AudioFix" /TR "C:\scripts\FixMic.ahk" /IT /V1