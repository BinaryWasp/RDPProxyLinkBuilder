import-Module "${env:ProgramFiles(x86)}\Devolutions\Remote Desktop Manager Free\RemoteDesktopManager.PowerShellModule.dll"

$computerName = "components";
$theusername = "mike";
$thedomain = "cyberark.local";
$thepassword = "Cyberark1";
$session = New-RDMSession -Host $computerName -Type "RDPConfigured" -Name $computerName;
Set-RDMSession -Session $session -Refresh;
Update-RDMUI;
Set-RDMSessionUsername -ID $session.ID $theusername;
Set-RDMSessionDomain -ID $session.ID $thedomain;
$pass = ConvertTo-SecureString $thepassword -asplaintext -force;
Set-RDMSessionPassword -ID $session.ID -Password $pass;