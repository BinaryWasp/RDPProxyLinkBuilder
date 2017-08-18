remove-item *.rdp
# write-output "RDP files removed. Housekeeping complete"

$csv = Import-Csv CSV.csv
foreach ($line in $csv) {
$RDPname=$line.filename + ".rdp"
$PSMserver=$line.PSM_Server
$Username=$line.username
$Server=$line.server
$EPVuser=$line.epv_username
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



}