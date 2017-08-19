function Set-PSMRDPFile {
    
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [String]
        $Path
    )

    foreach ($row in $csv) {

        # Import CSV file
        $csv = Import-Csv $Path

        # Set Variables
        $RDPname=$line.filename + ".rdp"
        $PSMserver=$line.PSM_Server
        $Username=$line.username
        $Server=$line.server
        $EPVuser=$line.epv_username

        # Create .rdp file
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
        Add-Content $RDPname "alternate shell:s:psm /u $Username /a $Server /c PSM-RDP"
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
    }
}

try {
    Set-PSMRDPFile -Path "CSV.csv"
}
catch {
    Write-Host "There was an error processing the CSV file.  Please try again." -ForegroundColor Red
}