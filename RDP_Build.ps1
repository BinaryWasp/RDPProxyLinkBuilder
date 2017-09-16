#Define Parameters
Param(
  [string]$filePath, [string]$Format
)

write-host $Format

#Load CSV
Write-Host "Onboarding file $FilePath. Please Wait" -ForegroundColor Yellow
$csv = Import-Csv $FilePath
	if ($format -eq "Devolutions")
	{import-Module "${env:ProgramFiles(x86)}\Devolutions\Remote Desktop Manager\RemoteDesktopManager.PowerShellModule.dll"}

foreach ($line in $csv) {

	#Define Variables
	   $RDPname=$line.filename + ".rdp"
	   $PSMserver=$line.PSM_Server
	   $Username=$line.username
	   $Server=$line.server
	   $EPVuser=$line.epv_username

		if ($Format -eq "RDP") { 
		#Building of RDP File
		   add-content $RDPname "screen mode id:i:2"
		   add-content $RDPname "use multimon:i:0"
		   add-content $RDPname "desktopwidth:i:1920"
		   add-content $RDPname "desktopheight:i:1080"
		   add-content $RDPname "session bpp:i:32"
		   add-content $RDPname "winposstr:s:0,3,0,0,800,600"
		   add-content $RDPname "compression:i:1"
 		   add-content $RDPname "keyboardhook:i:2"
		   add-content $RDPname "audiocapturemode:i:0"
		   add-content $RDPname "videoplaybackmode:i:1"
		   add-content $RDPname "connection type:i:7"
    	      	   add-content $RDPname "networkautodetect:i:1"
		   add-content $RDPname "bandwidthautodetect:i:1"
   		   add-content $RDPname "displayconnectionbar:i:1"
		   add-content $RDPname "importenableworkspacereconnect:i:0"
		   add-content $RDPname "disable wallpaper:i:0"
		   add-content $RDPname "allow font smoothing:i:0"
		   add-content $RDPname "allow desktop composition:i:0"
		   add-content $RDPname "disable full window drag:i:1"
		   add-content $RDPname "disable menu anims:i:1"
		   add-content $RDPname "disable themes:i:0"
		   add-content $RDPname "disable cursor setting:i:0"
		   add-content $RDPname "bitmapcachepersistenable:i:1"
		   add-content $RDPname "full address:s:$PSMserver"
		   add-content $RDPname "audiomode:i:0"
		   add-content $RDPname "redirectprinters:i:0"
		   add-content $RDPname "redirectcomports:i:0"
		   add-content $RDPname "redirectsmartcards:i:1"
		   add-content $RDPname "redirectclipboard:i:1"
		   add-content $RDPname "redirectposdevices:i:0"
   		   add-content $RDPname "autoreconnection enabled:i:1"
		   add-content $RDPname "authentication level:i:2"
		   add-content $RDPname "prompt for credentials:i:0"
		   add-content $RDPname "negotiate security layer:i:1"
		   add-content $RDPname "remoteapplicationmode:i:0"
		   add-content $RDPname "alternate shell:s:psm /u $Username /a $Server /c PSM-RDP"
 		   add-content $RDPname "shell working directory:s:"
		   add-content $RDPname "gatewayhostname:s:"
		   add-content $RDPname "gatewayusagemethod:i:4"
		   add-content $RDPname "gatewaycredentialssource:i:4"
		   add-content $RDPname "gatewayprofileusagemethod:i:0"
		   add-content $RDPname "promptcredentialonce:i:0"
		   add-content $RDPname "gatewaybrokeringtype:i:0"
		   add-content $RDPname "use redirection server name:i:0"
		   add-content $RDPname "rdgiskdcproxy:i:0"
		   add-content $RDPname "kdcproxyname:s:"
		   add-content $RDPname "username:s:$EPVuser"
		   add-content $RDPname "drivestoredirect:s:"
		Write-Host "$RDPname Created successfully." -ForegroundColor Green
		}
	
		if ($Format -eq "Devolutions") { 
		#Building of Devolutions File
		Get-RDMInstance
		$RDPname=$line.filename
		$computerName = $PSMserver;
		$theusername = $EPVuser;
		$session = New-RDMSession -Host $computerName -Type "RDPConfigured" -Name $RDPname;
		Set-RDMSession -Session $session -Refresh;
		Update-RDMUI #dont know why this is still necessary, it slows down the process significantly
		Set-RDMSessionProperty -ID $session.ID -Property "AlternateShell" -Value "psm /u $Username /a $Server /c PSM-RDP"
		#Write-Host "$RDPname Created successfully." -ForegroundColor Green
		}




if ($Format -eq "PSMP") { 
	#Building of REG File
	$REGname=$line.filename + ".reg"
	
add-content $REGname 'Windows Registry Editor Version 5.00'
add-content $REGname ''
$KeyName="[HKEY_CURRENT_USER\Software\SimonTatham\PuTTY\Sessions\"+ $line.filename +"]"
add-content $REGname $KeyName
add-content $REGname '"Present"=dword:00000001'
$KeyName='"HostName"="' +$PSMserver +'"'
add-content $REGname $KeyName
add-content $REGname '"LogFileName"="putty.log"'
add-content $REGname '"LogType"=dword:00000000'
add-content $REGname '"LogFileClash"=dword:ffffffff'
add-content $REGname '"LogFlush"=dword:00000001'
add-content $REGname '"SSHLogOmitPasswords"=dword:00000001'
add-content $REGname '"SSHLogOmitData"=dword:00000000'
add-content $REGname '"Protocol"="ssh"'
add-content $REGname '"PortNumber"=dword:00000016'
add-content $REGname '"CloseOnExit"=dword:00000001'
add-content $REGname '"WarnOnClose"=dword:00000001'
add-content $REGname '"PingInterval"=dword:00000000'
add-content $REGname '"PingIntervalSecs"=dword:00000000'
add-content $REGname '"TCPNoDelay"=dword:00000001'
add-content $REGname '"TCPKeepalives"=dword:00000000'
add-content $REGname '"TerminalType"="xterm"'
add-content $REGname '"TerminalSpeed"="38400,38400"'
add-content $REGname '"TerminalModes"="CS7=A,CS8=A,DISCARD=A,DSUSP=A,ECHO=A,ECHOCTL=A,ECHOE=A,ECHOK=A,ECHOKE=A,ECHONL=A,EOF=A,EOL=A,EOL2=A,ERASE=A,FLUSH=A,ICANON=A,ICRNL=A,IEXTEN=A,IGNCR=A,IGNPAR=A,IMAXBEL=A,INLCR=A,INPCK=A,INTR=A,ISIG=A,ISTRIP=A,IUCLC=A,IXANY=A,IXOFF=A,IXON=A,KILL=A,LNEXT=A,NOFLSH=A,OCRNL=A,OLCUC=A,ONLCR=A,ONLRET=A,ONOCR=A,OPOST=A,PARENB=A,PARMRK=A,PARODD=A,PENDIN=A,QUIT=A,REPRINT=A,START=A,STATUS=A,STOP=A,SUSP=A,SWTCH=A,TOSTOP=A,WERASE=A,XCASE=A"'
add-content $REGname '"AddressFamily"=dword:00000000'
add-content $REGname '"ProxyExcludeList"=""'
add-content $REGname '"ProxyDNS"=dword:00000001'
add-content $REGname '"ProxyLocalhost"=dword:00000000'
add-content $REGname '"ProxyMethod"=dword:00000000'
add-content $REGname '"ProxyHost"="proxy"'
add-content $REGname '"ProxyPort"=dword:00000050'
add-content $REGname '"ProxyUsername"=""'
add-content $REGname '"ProxyPassword"=""'
add-content $REGname '"ProxyTelnetCommand"="connect %host %port\\n"'
add-content $REGname '"Environment"=""'
add-content $REGname '"UserName"=""'
add-content $REGname '"UserNameFromEnvironment"=dword:00000000'
add-content $REGname '"LocalUserName"=""'
add-content $REGname '"NoPTY"=dword:00000000'
add-content $REGname '"Compression"=dword:00000000'
add-content $REGname '"TryAgent"=dword:00000001'
add-content $REGname '"AgentFwd"=dword:00000000'
add-content $REGname '"GssapiFwd"=dword:00000000'
add-content $REGname '"ChangeUsername"=dword:00000000'
add-content $REGname '"Cipher"="aes,blowfish,3des,WARN,arcfour,des"'
add-content $REGname '"KEX"="dh-gex-sha1,dh-group14-sha1,dh-group1-sha1,rsa,WARN"'
add-content $REGname '"RekeyTime"=dword:0000003c'
add-content $REGname '"RekeyBytes"="1G"'
add-content $REGname '"SshNoAuth"=dword:00000000'
add-content $REGname '"SshBanner"=dword:00000001'
add-content $REGname '"AuthTIS"=dword:00000000'
add-content $REGname '"AuthKI"=dword:00000001'
add-content $REGname '"AuthGSSAPI"=dword:00000001'
add-content $REGname '"GSSLibs"="gssapi32,sspi,custom"'
add-content $REGname '"GSSCustom"=""'
add-content $REGname '"SshNoShell"=dword:00000000'
add-content $REGname '"SshProt"=dword:00000003'
add-content $REGname '"LogHost"=""'
add-content $REGname '"SSH2DES"=dword:00000000'
add-content $REGname '"PublicKeyFile"=""'
$KeyName='"RemoteCommand"="'+ $EPVuser + " " + $Username + " " + $Username + '"'
add-content $REGname $KeyName
add-content $REGname '"RFCEnviron"=dword:00000000'
add-content $REGname '"PassiveTelnet"=dword:00000000'
add-content $REGname '"BackspaceIsDelete"=dword:00000001'
add-content $REGname '"RXVTHomeEnd"=dword:00000000'
add-content $REGname '"LinuxFunctionKeys"=dword:00000000'
add-content $REGname '"NoApplicationKeys"=dword:00000000'
add-content $REGname '"NoApplicationCursors"=dword:00000000'
add-content $REGname '"NoMouseReporting"=dword:00000000'
add-content $REGname '"NoRemoteResize"=dword:00000000'
add-content $REGname '"NoAltScreen"=dword:00000000'
add-content $REGname '"NoRemoteWinTitle"=dword:00000000'
add-content $REGname '"RemoteQTitleAction"=dword:00000001'
add-content $REGname '"NoDBackspace"=dword:00000000'
add-content $REGname '"NoRemoteCharset"=dword:00000000'
add-content $REGname '"ApplicationCursorKeys"=dword:00000000'
add-content $REGname '"ApplicationKeypad"=dword:00000000'
add-content $REGname '"NetHackKeypad"=dword:00000000'
add-content $REGname '"AltF4"=dword:00000001'
add-content $REGname '"AltSpace"=dword:00000000'
add-content $REGname '"AltOnly"=dword:00000000'
add-content $REGname '"ComposeKey"=dword:00000000'
add-content $REGname '"CtrlAltKeys"=dword:00000001'
add-content $REGname '"TelnetKey"=dword:00000000'
add-content $REGname '"TelnetRet"=dword:00000001'
add-content $REGname '"LocalEcho"=dword:00000002'
add-content $REGname '"LocalEdit"=dword:00000002'
add-content $REGname '"Answerback"="PuTTY"'
add-content $REGname '"AlwaysOnTop"=dword:00000000'
add-content $REGname '"FullScreenOnAltEnter"=dword:00000000'
add-content $REGname '"HideMousePtr"=dword:00000000'
add-content $REGname '"SunkenEdge"=dword:00000000'
add-content $REGname '"WindowBorder"=dword:00000001'
add-content $REGname '"CurType"=dword:00000000'
add-content $REGname '"BlinkCur"=dword:00000000'
add-content $REGname '"Beep"=dword:00000001'
add-content $REGname '"BeepInd"=dword:00000000'
add-content $REGname '"BellWaveFile"=""'
add-content $REGname '"BellOverload"=dword:00000001'
add-content $REGname '"BellOverloadN"=dword:00000005'
add-content $REGname '"BellOverloadT"=dword:000007d0'
add-content $REGname '"BellOverloadS"=dword:00001388'
add-content $REGname '"ScrollbackLines"=dword:000007d0'
add-content $REGname '"DECOriginMode"=dword:00000000'
add-content $REGname '"AutoWrapMode"=dword:00000001'
add-content $REGname '"LFImpliesCR"=dword:00000000'
add-content $REGname '"CRImpliesLF"=dword:00000000'
add-content $REGname '"DisableArabicShaping"=dword:00000000'
add-content $REGname '"DisableBidi"=dword:00000000'
add-content $REGname '"WinNameAlways"=dword:00000001'
add-content $REGname '"WinTitle"=""'
add-content $REGname '"TermWidth"=dword:00000050'
add-content $REGname '"TermHeight"=dword:00000018'
add-content $REGname '"Font"="Courier New"'
add-content $REGname '"FontIsBold"=dword:00000000'
add-content $REGname '"FontCharSet"=dword:00000000'
add-content $REGname '"FontHeight"=dword:0000000a'
add-content $REGname '"FontQuality"=dword:00000000'
add-content $REGname '"FontVTMode"=dword:00000004'
add-content $REGname '"UseSystemColours"=dword:00000000'
add-content $REGname '"TryPalette"=dword:00000000'
add-content $REGname '"ANSIColour"=dword:00000001'
add-content $REGname '"Xterm256Colour"=dword:00000001'
add-content $REGname '"BoldAsColour"=dword:00000001'
add-content $REGname '"Colour0"="187,187,187"'
add-content $REGname '"Colour1"="255,255,255"'
add-content $REGname '"Colour2"="0,0,0"'
add-content $REGname '"Colour3"="85,85,85"'
add-content $REGname '"Colour4"="0,0,0"'
add-content $REGname '"Colour5"="0,255,0"'
add-content $REGname '"Colour6"="0,0,0"'
add-content $REGname '"Colour7"="85,85,85"'
add-content $REGname '"Colour8"="187,0,0"'
add-content $REGname '"Colour9"="255,85,85"'
add-content $REGname '"Colour10"="0,187,0"'
add-content $REGname '"Colour11"="85,255,85"'
add-content $REGname '"Colour12"="187,187,0"'
add-content $REGname '"Colour13"="255,255,85"'
add-content $REGname '"Colour14"="0,0,187"'
add-content $REGname '"Colour15"="85,85,255"'
add-content $REGname '"Colour16"="187,0,187"'
add-content $REGname '"Colour17"="255,85,255"'
add-content $REGname '"Colour18"="0,187,187"'
add-content $REGname '"Colour19"="85,255,255"'
add-content $REGname '"Colour20"="187,187,187"'
add-content $REGname '"Colour21"="255,255,255"'
add-content $REGname '"RawCNP"=dword:00000000'
add-content $REGname '"PasteRTF"=dword:00000000'
add-content $REGname '"MouseIsXterm"=dword:00000000'
add-content $REGname '"RectSelect"=dword:00000000'
add-content $REGname '"MouseOverride"=dword:00000001'
add-content $REGname '"Wordness0"="0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"'
add-content $REGname '"Wordness32"="0,1,2,1,1,1,1,1,1,1,1,1,1,2,2,2,2,2,2,2,2,2,2,2,2,2,1,1,1,1,1,1"'
add-content $REGname '"Wordness64"="1,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,1,1,1,1,2"'
add-content $REGname '"Wordness96"="1,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,1,1,1,1,1"'
add-content $REGname '"Wordness128"="1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1"'
add-content $REGname '"Wordness160"="1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1"'
add-content $REGname '"Wordness192"="2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,1,2,2,2,2,2,2,2,2"'
add-content $REGname '"Wordness224"="2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,1,2,2,2,2,2,2,2,2"'
add-content $REGname '"LineCodePage"=""'
add-content $REGname '"CJKAmbigWide"=dword:00000000'
add-content $REGname '"UTF8Override"=dword:00000001'
add-content $REGname '"Printer"=""'
add-content $REGname '"CapsLockCyr"=dword:00000000'
add-content $REGname '"ScrollBar"=dword:00000001'
add-content $REGname '"ScrollBarFullScreen"=dword:00000000'
add-content $REGname '"ScrollOnKey"=dword:00000000'
add-content $REGname '"ScrollOnDisp"=dword:00000001'
add-content $REGname '"EraseToScrollback"=dword:00000001'
add-content $REGname '"LockSize"=dword:00000000'
add-content $REGname '"BCE"=dword:00000001'
add-content $REGname '"BlinkText"=dword:00000000'
add-content $REGname '"X11Forward"=dword:00000000'
add-content $REGname '"X11Display"=""'
add-content $REGname '"X11AuthType"=dword:00000001'
add-content $REGname '"X11AuthFile"=""'
add-content $REGname '"LocalPortAcceptAll"=dword:00000000'
add-content $REGname '"RemotePortAcceptAll"=dword:00000000'
add-content $REGname '"PortForwardings"=""'
add-content $REGname '"BugIgnore1"=dword:00000000'
add-content $REGname '"BugPlainPW1"=dword:00000000'
add-content $REGname '"BugRSA1"=dword:00000000'
add-content $REGname '"BugIgnore2"=dword:00000000'
add-content $REGname '"BugHMAC2"=dword:00000000'
add-content $REGname '"BugDeriveKey2"=dword:00000000'
add-content $REGname '"BugRSAPad2"=dword:00000000'
add-content $REGname '"BugPKSessID2"=dword:00000000'
add-content $REGname '"BugRekey2"=dword:00000000'
add-content $REGname '"BugMaxPkt2"=dword:00000000'
add-content $REGname '"BugOldGex2"=dword:00000000'
add-content $REGname '"BugWinadj"=dword:00000000'
add-content $REGname '"BugChanReq"=dword:00000000'
add-content $REGname '"StampUtmp"=dword:00000001'
add-content $REGname '"LoginShell"=dword:00000001'
add-content $REGname '"ScrollbarOnLeft"=dword:00000000'
add-content $REGname '"BoldFont"=""'
add-content $REGname '"BoldFontIsBold"=dword:00000000'
add-content $REGname '"BoldFontCharSet"=dword:00000000'
add-content $REGname '"BoldFontHeight"=dword:00000000'
add-content $REGname '"WideFont"=""'
add-content $REGname '"WideFontIsBold"=dword:00000000'
add-content $REGname '"WideFontCharSet"=dword:00000000'
add-content $REGname '"WideFontHeight"=dword:00000000'
add-content $REGname '"WideBoldFont"=""'
add-content $REGname '"WideBoldFontIsBold"=dword:00000000'
add-content $REGname '"WideBoldFontCharSet"=dword:00000000'
add-content $REGname '"WideBoldFontHeight"=dword:00000000'
add-content $REGname '"ShadowBold"=dword:00000000'
add-content $REGname '"ShadowBoldOffset"=dword:00000001'
add-content $REGname '"SerialLine"="COM1"'
add-content $REGname '"SerialSpeed"=dword:00002580'
add-content $REGname '"SerialDataBits"=dword:00000008'
add-content $REGname '"SerialStopHalfbits"=dword:00000002'
add-content $REGname '"SerialParity"=dword:00000000'
add-content $REGname '"SerialFlowControl"=dword:00000001'
add-content $REGname '"WindowClass"=""'
add-content $REGname '"ConnectionSharing"=dword:00000000'
add-content $REGname '"ConnectionSharingUpstream"=dword:00000001'
add-content $REGname '"ConnectionSharingDownstream"=dword:00000001'
add-content $REGname '"SSHManualHostKeys"=""'
add-content $REGname ''
add-content $REGname ''
reg import $REGname
remove-item $REGname
write-host "Session Loaded to PuTTY" -ForegroundColor Green
		}
}
