#Define Parameters
Param(
	[string]$filePath, [string]$Format
)

Write-Host $Format

# Load CSV
Write-Host "Onboarding file $FilePath. Please Wait" -ForegroundColor Yellow
$csv = Import-Csv $FilePath
If ($format -eq "Devolutions") {
	Import-Module "${env:ProgramFiles(x86)}\Devolutions\Remote Desktop Manager\RemoteDesktopManager.PowerShellModule.dll"
}

ForEach ($row in $csv) {

	# Define Variables
	$RDPname = $row.filename + ".rdp"
	$PSMserver = $row.PSM_Server
	$Username = $row.username
	$Server = $row.server
	$EPVuser = $row.epv_username

	$connectionString = "psm /u $Username /a $Server /c PSM-RDP"

	If ($Format.ToLower() -eq "rdp") { 
		#Building of RDP File
		Add-Content $RDPname "screen mode id:i:2"
		Add-Content $RDPname "use multimon:i:0"
		Add-Content $RDPname "desktopwidth:i:1920"
		Add-Content $RDPname "desktopheight:i:1080"
		Add-Content $RDPname "session bpp:i:32"
		Add-Content $RDPname "winposstr:s:0,3,0,0,800,600"
		Add-Content $RDPname "compression:i:1"
		Add-Content $RDPname "keyboardhook:i:2"
		Add-Content $RDPname "audiocapturemode:i:0"
		Add-Content $RDPname "videoplaybackmode:i:1"
		Add-Content $RDPname "connection type:i:7"
		Add-Content $RDPname "networkautodetect:i:1"
		Add-Content $RDPname "bandwidthautodetect:i:1"
		Add-Content $RDPname "displayconnectionbar:i:1"
		Add-Content $RDPname "importenableworkspacereconnect:i:0"
		Add-Content $RDPname "disable wallpaper:i:0"
		Add-Content $RDPname "allow font smoothing:i:0"
		Add-Content $RDPname "allow desktop composition:i:0"
		Add-Content $RDPname "disable full window drag:i:1"
		Add-Content $RDPname "disable menu anims:i:1"
		Add-Content $RDPname "disable themes:i:0"
		Add-Content $RDPname "disable cursor setting:i:0"
		Add-Content $RDPname "bitmapcachepersistenable:i:1"
		Add-Content $RDPname "full address:s:$PSMserver"
		Add-Content $RDPname "audiomode:i:0"
		Add-Content $RDPname "redirectprinters:i:0"
		Add-Content $RDPname "redirectcomports:i:0"
		Add-Content $RDPname "redirectsmartcards:i:1"
		Add-Content $RDPname "redirectclipboard:i:1"
		Add-Content $RDPname "redirectposdevices:i:0"
		Add-Content $RDPname "autoreconnection enabled:i:1"
		Add-Content $RDPname "authentication level:i:2"
		Add-Content $RDPname "prompt for credentials:i:0"
		Add-Content $RDPname "negotiate security layer:i:1"
		Add-Content $RDPname "remoteapplicationmode:i:0"
		Add-Content $RDPname "alternate shell:s:$connectionString"
		Add-Content $RDPname "shell working directory:s:"
		Add-Content $RDPname "gatewayhostname:s:"
		Add-Content $RDPname "gatewayusagemethod:i:4"
		Add-Content $RDPname "gatewaycredentialssource:i:4"
		Add-Content $RDPname "gatewayprofileusagemethod:i:0"
		Add-Content $RDPname "promptcredentialonce:i:0"
		Add-Content $RDPname "gatewaybrokeringtype:i:0"
		Add-Content $RDPname "use redirection server name:i:0"
		Add-Content $RDPname "rdgiskdcproxy:i:0"
		Add-Content $RDPname "kdcproxyname:s:"
		Add-Content $RDPname "username:s:$EPVuser"
		Add-Content $RDPname "drivestoredirect:s:"
		Write-Host "$RDPname Created successfully." -ForegroundColor Green
	}

	if ($Format.ToLower() -eq "devolutions") { 
		#Building of Devolutions File
		Get-RDMInstance
		$RDPname=$line.filename
		$computerName = $PSMserver;
		$session = New-RDMSession -Host $computerName -Type "RDPConfigured" -Name $RDPname;
		Set-RDMSession -Session $session -Refresh;
		Update-RDMUI #dont know why this is still necessary, it slows down the process significantly
		Set-RDMSessionProperty -ID $session.ID -Property "AlternateShell" -Value $connectionString
		#Write-Host "$RDPname Created successfully." -ForegroundColor Green
	}

	If ($Format.ToLower() -eq "putty") { 
		#Building of REG File
		$REGname=$line.filename + ".reg"

		Add-Content $REGname 'Windows Registry Editor Version 5.00'
		Add-Content $REGname ''
		$KeyName="[HKEY_CURRENT_USER\Software\SimonTatham\PuTTY\Sessions\"+ $line.filename +"]"
		Add-Content $REGname $KeyName
		Add-Content $REGname '"Present"=dword:00000001'
		$KeyName='"HostName"="' +$PSMserver +'"'
		Add-Content $REGname $KeyName
		Add-Content $REGname '"LogFileName"="putty.log"'
		Add-Content $REGname '"LogType"=dword:00000000'
		Add-Content $REGname '"LogFileClash"=dword:ffffffff'
		Add-Content $REGname '"LogFlush"=dword:00000001'
		Add-Content $REGname '"SSHLogOmitPasswords"=dword:00000001'
		Add-Content $REGname '"SSHLogOmitData"=dword:00000000'
		Add-Content $REGname '"Protocol"="ssh"'
		Add-Content $REGname '"PortNumber"=dword:00000016'
		Add-Content $REGname '"CloseOnExit"=dword:00000001'
		Add-Content $REGname '"WarnOnClose"=dword:00000001'
		Add-Content $REGname '"PingInterval"=dword:00000000'
		Add-Content $REGname '"PingIntervalSecs"=dword:00000000'
		Add-Content $REGname '"TCPNoDelay"=dword:00000001'
		Add-Content $REGname '"TCPKeepalives"=dword:00000000'
		Add-Content $REGname '"TerminalType"="xterm"'
		Add-Content $REGname '"TerminalSpeed"="38400,38400"'
		Add-Content $REGname '"TerminalModes"="CS7=A,CS8=A,DISCARD=A,DSUSP=A,ECHO=A,ECHOCTL=A,ECHOE=A,ECHOK=A,ECHOKE=A,ECHONL=A,EOF=A,EOL=A,EOL2=A,ERASE=A,FLUSH=A,ICANON=A,ICRNL=A,IEXTEN=A,IGNCR=A,IGNPAR=A,IMAXBEL=A,INLCR=A,INPCK=A,INTR=A,ISIG=A,ISTRIP=A,IUCLC=A,IXANY=A,IXOFF=A,IXON=A,KILL=A,LNEXT=A,NOFLSH=A,OCRNL=A,OLCUC=A,ONLCR=A,ONLRET=A,ONOCR=A,OPOST=A,PARENB=A,PARMRK=A,PARODD=A,PENDIN=A,QUIT=A,REPRINT=A,START=A,STATUS=A,STOP=A,SUSP=A,SWTCH=A,TOSTOP=A,WERASE=A,XCASE=A"'
		Add-Content $REGname '"AddressFamily"=dword:00000000'
		Add-Content $REGname '"ProxyExcludeList"=""'
		Add-Content $REGname '"ProxyDNS"=dword:00000001'
		Add-Content $REGname '"ProxyLocalhost"=dword:00000000'
		Add-Content $REGname '"ProxyMethod"=dword:00000000'
		Add-Content $REGname '"ProxyHost"="proxy"'
		Add-Content $REGname '"ProxyPort"=dword:00000050'
		Add-Content $REGname '"ProxyUsername"=""'
		Add-Content $REGname '"ProxyPassword"=""'
		Add-Content $REGname '"ProxyTelnetCommand"="connect %host %port\\n"'
		Add-Content $REGname '"Environment"=""'
		Add-Content $REGname '"UserName"="PSMConnect"'
		Add-Content $REGname '"UserNameFromEnvironment"=dword:00000000'
		Add-Content $REGname '"LocalUserName"=""'
		Add-Content $REGname '"NoPTY"=dword:00000000'
		Add-Content $REGname '"Compression"=dword:00000000'
		Add-Content $REGname '"TryAgent"=dword:00000001'
		Add-Content $REGname '"AgentFwd"=dword:00000000'
		Add-Content $REGname '"GssapiFwd"=dword:00000000'
		Add-Content $REGname '"ChangeUsername"=dword:00000000'
		Add-Content $REGname '"Cipher"="aes,blowfish,3des,WARN,arcfour,des"'
		Add-Content $REGname '"KEX"="dh-gex-sha1,dh-group14-sha1,dh-group1-sha1,rsa,WARN"'
		Add-Content $REGname '"RekeyTime"=dword:0000003c'
		Add-Content $REGname '"RekeyBytes"="1G"'
		Add-Content $REGname '"SshNoAuth"=dword:00000000'
		Add-Content $REGname '"SshBanner"=dword:00000001'
		Add-Content $REGname '"AuthTIS"=dword:00000000'
		Add-Content $REGname '"AuthKI"=dword:00000001'
		Add-Content $REGname '"AuthGSSAPI"=dword:00000001'
		Add-Content $REGname '"GSSLibs"="gssapi32,sspi,custom"'
		Add-Content $REGname '"GSSCustom"=""'
		Add-Content $REGname '"SshNoShell"=dword:00000000'
		Add-Content $REGname '"SshProt"=dword:00000003'
		Add-Content $REGname '"LogHost"=""'
		Add-Content $REGname '"SSH2DES"=dword:00000000'
		Add-Content $REGname '"PublicKeyFile"=""'
		$KeyName='"RemoteCommand"="'+ $EPVuser + " " + $Username + " " + $Server + '"'
		Add-Content $REGname $KeyName
		Add-Content $REGname '"RFCEnviron"=dword:00000000'
		Add-Content $REGname '"PassiveTelnet"=dword:00000000'
		Add-Content $REGname '"BackspaceIsDelete"=dword:00000001'
		Add-Content $REGname '"RXVTHomeEnd"=dword:00000000'
		Add-Content $REGname '"LinuxFunctionKeys"=dword:00000000'
		Add-Content $REGname '"NoApplicationKeys"=dword:00000000'
		Add-Content $REGname '"NoApplicationCursors"=dword:00000000'
		Add-Content $REGname '"NoMouseReporting"=dword:00000000'
		Add-Content $REGname '"NoRemoteResize"=dword:00000000'
		Add-Content $REGname '"NoAltScreen"=dword:00000000'
		Add-Content $REGname '"NoRemoteWinTitle"=dword:00000000'
		Add-Content $REGname '"RemoteQTitleAction"=dword:00000001'
		Add-Content $REGname '"NoDBackspace"=dword:00000000'
		Add-Content $REGname '"NoRemoteCharset"=dword:00000000'
		Add-Content $REGname '"ApplicationCursorKeys"=dword:00000000'
		Add-Content $REGname '"ApplicationKeypad"=dword:00000000'
		Add-Content $REGname '"NetHackKeypad"=dword:00000000'
		Add-Content $REGname '"AltF4"=dword:00000001'
		Add-Content $REGname '"AltSpace"=dword:00000000'
		Add-Content $REGname '"AltOnly"=dword:00000000'
		Add-Content $REGname '"ComposeKey"=dword:00000000'
		Add-Content $REGname '"CtrlAltKeys"=dword:00000001'
		Add-Content $REGname '"TelnetKey"=dword:00000000'
		Add-Content $REGname '"TelnetRet"=dword:00000001'
		Add-Content $REGname '"LocalEcho"=dword:00000002'
		Add-Content $REGname '"LocalEdit"=dword:00000002'
		Add-Content $REGname '"Answerback"="PuTTY"'
		Add-Content $REGname '"AlwaysOnTop"=dword:00000000'
		Add-Content $REGname '"FullScreenOnAltEnter"=dword:00000000'
		Add-Content $REGname '"HideMousePtr"=dword:00000000'
		Add-Content $REGname '"SunkenEdge"=dword:00000000'
		Add-Content $REGname '"WindowBorder"=dword:00000001'
		Add-Content $REGname '"CurType"=dword:00000000'
		Add-Content $REGname '"BlinkCur"=dword:00000000'
		Add-Content $REGname '"Beep"=dword:00000001'
		Add-Content $REGname '"BeepInd"=dword:00000000'
		Add-Content $REGname '"BellWaveFile"=""'
		Add-Content $REGname '"BellOverload"=dword:00000001'
		Add-Content $REGname '"BellOverloadN"=dword:00000005'
		Add-Content $REGname '"BellOverloadT"=dword:000007d0'
		Add-Content $REGname '"BellOverloadS"=dword:00001388'
		Add-Content $REGname '"ScrollbackLines"=dword:000007d0'
		Add-Content $REGname '"DECOriginMode"=dword:00000000'
		Add-Content $REGname '"AutoWrapMode"=dword:00000001'
		Add-Content $REGname '"LFImpliesCR"=dword:00000000'
		Add-Content $REGname '"CRImpliesLF"=dword:00000000'
		Add-Content $REGname '"DisableArabicShaping"=dword:00000000'
		Add-Content $REGname '"DisableBidi"=dword:00000000'
		Add-Content $REGname '"WinNameAlways"=dword:00000001'
		Add-Content $REGname '"WinTitle"=""'
		Add-Content $REGname '"TermWidth"=dword:00000050'
		Add-Content $REGname '"TermHeight"=dword:00000018'
		Add-Content $REGname '"Font"="Courier New"'
		Add-Content $REGname '"FontIsBold"=dword:00000000'
		Add-Content $REGname '"FontCharSet"=dword:00000000'
		Add-Content $REGname '"FontHeight"=dword:0000000a'
		Add-Content $REGname '"FontQuality"=dword:00000000'
		Add-Content $REGname '"FontVTMode"=dword:00000004'
		Add-Content $REGname '"UseSystemColours"=dword:00000000'
		Add-Content $REGname '"TryPalette"=dword:00000000'
		Add-Content $REGname '"ANSIColour"=dword:00000001'
		Add-Content $REGname '"Xterm256Colour"=dword:00000001'
		Add-Content $REGname '"BoldAsColour"=dword:00000001'
		Add-Content $REGname '"Colour0"="187,187,187"'
		Add-Content $REGname '"Colour1"="255,255,255"'
		Add-Content $REGname '"Colour2"="0,0,0"'
		Add-Content $REGname '"Colour3"="85,85,85"'
		Add-Content $REGname '"Colour4"="0,0,0"'
		Add-Content $REGname '"Colour5"="0,255,0"'
		Add-Content $REGname '"Colour6"="0,0,0"'
		Add-Content $REGname '"Colour7"="85,85,85"'
		Add-Content $REGname '"Colour8"="187,0,0"'
		Add-Content $REGname '"Colour9"="255,85,85"'
		Add-Content $REGname '"Colour10"="0,187,0"'
		Add-Content $REGname '"Colour11"="85,255,85"'
		Add-Content $REGname '"Colour12"="187,187,0"'
		Add-Content $REGname '"Colour13"="255,255,85"'
		Add-Content $REGname '"Colour14"="0,0,187"'
		Add-Content $REGname '"Colour15"="85,85,255"'
		Add-Content $REGname '"Colour16"="187,0,187"'
		Add-Content $REGname '"Colour17"="255,85,255"'
		Add-Content $REGname '"Colour18"="0,187,187"'
		Add-Content $REGname '"Colour19"="85,255,255"'
		Add-Content $REGname '"Colour20"="187,187,187"'
		Add-Content $REGname '"Colour21"="255,255,255"'
		Add-Content $REGname '"RawCNP"=dword:00000000'
		Add-Content $REGname '"PasteRTF"=dword:00000000'
		Add-Content $REGname '"MouseIsXterm"=dword:00000000'
		Add-Content $REGname '"RectSelect"=dword:00000000'
		Add-Content $REGname '"MouseOverride"=dword:00000001'
		Add-Content $REGname '"Wordness0"="0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"'
		Add-Content $REGname '"Wordness32"="0,1,2,1,1,1,1,1,1,1,1,1,1,2,2,2,2,2,2,2,2,2,2,2,2,2,1,1,1,1,1,1"'
		Add-Content $REGname '"Wordness64"="1,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,1,1,1,1,2"'
		Add-Content $REGname '"Wordness96"="1,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,1,1,1,1,1"'
		Add-Content $REGname '"Wordness128"="1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1"'
		Add-Content $REGname '"Wordness160"="1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1"'
		Add-Content $REGname '"Wordness192"="2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,1,2,2,2,2,2,2,2,2"'
		Add-Content $REGname '"Wordness224"="2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,1,2,2,2,2,2,2,2,2"'
		Add-Content $REGname '"LineCodePage"=""'
		Add-Content $REGname '"CJKAmbigWide"=dword:00000000'
		Add-Content $REGname '"UTF8Override"=dword:00000001'
		Add-Content $REGname '"Printer"=""'
		Add-Content $REGname '"CapsLockCyr"=dword:00000000'
		Add-Content $REGname '"ScrollBar"=dword:00000001'
		Add-Content $REGname '"ScrollBarFullScreen"=dword:00000000'
		Add-Content $REGname '"ScrollOnKey"=dword:00000000'
		Add-Content $REGname '"ScrollOnDisp"=dword:00000001'
		Add-Content $REGname '"EraseToScrollback"=dword:00000001'
		Add-Content $REGname '"LockSize"=dword:00000000'
		Add-Content $REGname '"BCE"=dword:00000001'
		Add-Content $REGname '"BlinkText"=dword:00000000'
		Add-Content $REGname '"X11Forward"=dword:00000000'
		Add-Content $REGname '"X11Display"=""'
		Add-Content $REGname '"X11AuthType"=dword:00000001'
		Add-Content $REGname '"X11AuthFile"=""'
		Add-Content $REGname '"LocalPortAcceptAll"=dword:00000000'
		Add-Content $REGname '"RemotePortAcceptAll"=dword:00000000'
		Add-Content $REGname '"PortForwardings"=""'
		Add-Content $REGname '"BugIgnore1"=dword:00000000'
		Add-Content $REGname '"BugPlainPW1"=dword:00000000'
		Add-Content $REGname '"BugRSA1"=dword:00000000'
		Add-Content $REGname '"BugIgnore2"=dword:00000000'
		Add-Content $REGname '"BugHMAC2"=dword:00000000'
		Add-Content $REGname '"BugDeriveKey2"=dword:00000000'
		Add-Content $REGname '"BugRSAPad2"=dword:00000000'
		Add-Content $REGname '"BugPKSessID2"=dword:00000000'
		Add-Content $REGname '"BugRekey2"=dword:00000000'
		Add-Content $REGname '"BugMaxPkt2"=dword:00000000'
		Add-Content $REGname '"BugOldGex2"=dword:00000000'
		Add-Content $REGname '"BugWinadj"=dword:00000000'
		Add-Content $REGname '"BugChanReq"=dword:00000000'
		Add-Content $REGname '"StampUtmp"=dword:00000001'
		Add-Content $REGname '"LoginShell"=dword:00000001'
		Add-Content $REGname '"ScrollbarOnLeft"=dword:00000000'
		Add-Content $REGname '"BoldFont"=""'
		Add-Content $REGname '"BoldFontIsBold"=dword:00000000'
		Add-Content $REGname '"BoldFontCharSet"=dword:00000000'
		Add-Content $REGname '"BoldFontHeight"=dword:00000000'
		Add-Content $REGname '"WideFont"=""'
		Add-Content $REGname '"WideFontIsBold"=dword:00000000'
		Add-Content $REGname '"WideFontCharSet"=dword:00000000'
		Add-Content $REGname '"WideFontHeight"=dword:00000000'
		Add-Content $REGname '"WideBoldFont"=""'
		Add-Content $REGname '"WideBoldFontIsBold"=dword:00000000'
		Add-Content $REGname '"WideBoldFontCharSet"=dword:00000000'
		Add-Content $REGname '"WideBoldFontHeight"=dword:00000000'
		Add-Content $REGname '"ShadowBold"=dword:00000000'
		Add-Content $REGname '"ShadowBoldOffset"=dword:00000001'
		Add-Content $REGname '"SerialLine"="COM1"'
		Add-Content $REGname '"SerialSpeed"=dword:00002580'
		Add-Content $REGname '"SerialDataBits"=dword:00000008'
		Add-Content $REGname '"SerialStopHalfbits"=dword:00000002'
		Add-Content $REGname '"SerialParity"=dword:00000000'
		Add-Content $REGname '"SerialFlowControl"=dword:00000001'
		Add-Content $REGname '"WindowClass"=""'
		Add-Content $REGname '"ConnectionSharing"=dword:00000000'
		Add-Content $REGname '"ConnectionSharingUpstream"=dword:00000001'
		Add-Content $REGname '"ConnectionSharingDownstream"=dword:00000001'
		Add-Content $REGname '"SSHManualHostKeys"=""'
		Add-Content $REGname ''
		Add-Content $REGname ''
		reg import $REGname
		Remove-Item $REGname
		Write-Host "Session Loaded to PuTTY" -ForegroundColor Green
	}
}
	
If ($Format.ToLower() -eq "rdman") {

	$RDGFileOutPut =  $PSScriptRoot + "\" +"CyberArk_Connections.rdg"
	#$RDGFileOutPut = "C:\Users\randy\Desktop\test.xml"
	$XmlWriter = New-Object System.XMl.XmlTextWriter($RDGFileOutPut,$Null)
	$XmlWriter.Formatting = 'Indented'
	$XmlWriter.Indentation = 1
	$XmlWriter.IndentChar = "`t"

	# Create header
	$XmlWriter.WriteStartDocument()

	# Create RDCMan element
	$XmlWriter.WriteStartElement('RDCMan')
	$XmlWriter.WriteAttributeString('programVersion', '2.7')
	$XmlWriter.WriteAttributeString('schemaVersion', '3')

	# Create file element
	$XmlWriter.WriteStartElement('file')

	# Create credentialsProfiles element
	$XmlWriter.WriteStartElement('credentialsProfiles')
	$XmlWriter.WriteEndElement()

	# Create properties element
	$XmlWriter.WriteStartElement('properties')
	$XmlWriter.WriteElementString('expanded', 'True')
	$XmlWriter.WriteElementString('name', 'CyberArk PSM Connections')
	$XmlWriter.WriteEndElement()

	ForEach ($row in $csv) {
		# Define Variables
		$psmServer = $row.PSM_Server
		$userName = $row.username
		$server = $row.server
		$epvUser = $row.epv_username
		$connectionString = "psm /u $userName /a $server /c PSM-RDP"

		# Create server element
		$XmlWriter.WriteStartElement('server')

		# Create properties element, add properties, close it
		$XmlWriter.WriteStartElement('properties')
		$XmlWriter.WriteElementString('displayName', $server)
		$XmlWriter.WriteElementString('name', $PSMserver)
		$XmlWriter.WriteEndElement()

		# Create logonCredentials element, add properties, close it
		$XmlWriter.WriteStartElement('logonCredentials')
		$XmlWriter.WriteAttributeString('inherit', 'None')
		$XmlWriter.WriteStartElement('profileName')
		$XmlWriter.WriteAttributeString('scope', 'Local')
		$XmlWriter.WriteRaw('Custom')
		$xmlWriter.WriteEndElement()
		$XmlWriter.WriteElementString('userName', $epvUser)
		$XmlWriter.WriteStartElement('password')
		$XmlWriter.WriteEndElement()
		$XmlWriter.WriteStartElement('domain')
		$XmlWriter.WriteEndElement()
		$XmlWriter.WriteEndElement()

		# Create connectionSettings element, add properties, close it
		$XmlWriter.WriteStartElement('connectionSettings')
		$XmlWriter.WriteAttributeString('inherit', 'None')
		$XmlWriter.WriteElementString('connectToConsole', 'False')
		$XmlWriter.WriteElementString('startProgram', $connectionString)
		$XmlWriter.WriteStartElement('workingDir')
		$XmlWriter.WriteEndElement()
		$XmlWriter.WriteElementString('port', '3389')
		$XmlWriter.WriteStartElement('loadBalanceInfo')
		$XmlWriter.WriteEndElement()
		$XmlWriter.WriteEndElement()

		# Close server element
		$XmlWriter.WriteEndElement()
	}
	
	# Close file element
	$XmlWriter.WriteEndElement()

	# Close RDMan element
	$XmlWriter.WriteEndElement()

	# Write to file
	$xmlWriter.WriteEndDocument()
	$xmlWriter.Flush()
	$xmlWriter.Close()
}