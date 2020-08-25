<#
.SYNOPSIS
  Create HTML page with a mapping of VPN Users to VPN IP Addresses.
.DESCRIPTION
  This script will create a HTML page that wil display a list of VPN Users currently connected and their IP addresses. This is useful for support to check and for IP addresses when remote support
  users via VNC.

.INPUTS
  N/A
.OUTPUTS
  A pretty HTML file to list the current logged in users via MS VPN.
.NOTES
  Version:        1.0
  Author:         Peter Wahid
  Creation Date:  06/07/2020
  Purpose/Change: N/A
  
.EXAMPLE
  N/A
#>

Clear-Host
#Estimate script runtime
$stopwatch = [system.diagnostics.stopwatch]::StartNew()
$stopwatch.start()


$webFile = "\\YOURWEBSERVER\vpnstatus.htm"
$vpnSrv = "YOUR-MSVPN-SERVER"


function Get-VPNUserIP{
    $rasStatus = Get-RemoteAccessConnectionStatistics -ComputerName $vpnSrv
    #Ping -a the IP of each PC to get the hostname if possible
    foreach($property in $rasStatus){
        $dataHash = [ordered]@{
            UserName            = $($property.UserName);
            ClientIPAddress     = $($property.ClientIPAddress);
            ConnectDuration     = $($property.ConnectDuration);
        }
        New-Object PSObject -Property $dataHash
    }
}


function New-HTMLBody($report) {
    $head = @"
  <html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" xmlns="http://www.w3.org/TR/REC-html40"><head><meta http-equiv=Content-Type content="text/html; charset=us-ascii"><meta name=Generator content="Microsoft Word 15 (filtered medium)"><meta http-equiv="refresh" content="30"><style><!--
  /* Font Definitions */
    @font-face
    {font-family:"Courier New";
    panose-1:2 4 5 3 5 4 6 3 2 4;}
  @font-face
    {font-family:Courier New;
    panose-1:2 15 5 2 2 2 4 3 2 4;}
  /* Style Definitions */
  p.MsoNormal, li.MsoNormal, div.MsoNormal
    {margin:0cm;
    margin-bottom:.0001pt;
    font-size:11.0pt;
    font-family:"Courier New",sans-serif;
    mso-fareast-language:EN-US;}
  a:link, span.MsoHyperlink
    {mso-style-priority:99;
    color:#0563C1;
    text-decoration:underline;}
  a:visited, span.MsoHyperlinkFollowed
    {mso-style-priority:99;
     color:#954F72;
    text-decoration:underline;}
  span.EmailStyle17
    {mso-style-type:personal-compose;
    font-family:"Courier New",sans-serif;
    color:windowtext;}
  .MsoChpDefault
    {mso-style-type:export-only;
    font-family:"Courier New",sans-serif;
    mso-fareast-language:EN-US;}
  @page WordSection1
    {size:612.0pt 792.0pt;
    margin:72.0pt 72.0pt 72.0pt 72.0pt;}
  div.WordSection1
    {page:WordSection1;}
  --></style><!--[if gte mso 9]><xml>
  <o:shapedefaults v:ext="edit" spidmax="1026" />
  </xml><![endif]--><!--[if gte mso 9]><xml>
  <o:shapelayout v:ext="edit">
  <o:idmap v:ext="edit" data="1" />
  </o:shapelayout></xml><![endif]-->
  
  </head>
"@


    #Start Body

    $body = @"
  <body lang=EN-AU link="#0563C1" vlink="#954F72" bgcolor="#000000" text="green">
    <div class=WordSection1>
      <p class=MsoNormal>
"@

    $body += @"
    <H1 style=font-family:"Courier New"> VPN User Status </H1>
    <H4 style=font-family:"Courier New"><i> Page will refresh every 30 seconds, no historical data is kept.</i></H3>
    <!--- Loop deployment status ---!>
    <table class=MsoTable15Grid4Accent1 border=1 cellspacing=0 cellpadding=0 style='border-collapse:collapse;border:none'>
      <tr>
        <td width=89 valign=top style='width:66.75pt;border:solid #000000 1.0pt;border-right:solid #000000 1.0pt;background:#000000;padding:0cm 5.4pt 0cm 5.4pt'><p class=MsoNormal><b><span style='color:green'><o:p>Name</o:p></span></b></p></td>
        <td width=89 valign=top style='width:66.75pt;border:solid #000000 1.0pt;border-right:solid #000000 1.0pt;background:#000000;padding:0cm 5.4pt 0cm 5.4pt'><p class=MsoNormal><b><span style='color:green'><o:p>IP</o:p></span></b></p></td>
        <td width=89 valign=top style='width:66.75pt;border:solid #000000 1.0pt;border-right:solid #000000 1.0pt;background:#000000;padding:0cm 5.4pt 0cm 5.4pt'><p class=MsoNormal><b><span style='color:green'><o:p>Duration</o:p></span></b></p></td>
      </tr>
"@ 

    foreach ($item in $report) {
        [string]$UserName = $item.UserName
        [string]$ClientIPAddress = $item.ClientIPAddress
        [string]$ConnectDuration = $item.ConnectDuration


        #Test status colour
        #$deploymentStatus ="Active/Running"


       
        $body += @"
      <tr>
        <td width=89 valign=top style='width:80pt;border:solid #000000 1.0pt;border-right:solid #000000 1.0pt;background:#000000;padding:0cm 5.4pt 0cm 5.4pt'><p class=MsoNormal><b><span style='color:green'><o:p>$UserName</o:p></span></b></p></td>
        <td width=89 valign=top style='width:80pt;border:solid #000000 1.0pt;border-right:solid #000000 1.0pt;background:#000000;padding:0cm 5.4pt 0cm 5.4pt'><p class=MsoNormal><b><span style='color:green'><o:p>$ClientIPAddress</o:p></span></b></p></td>
        <td width=89 valign=top style='width:180pt;border:solid #000000 1.0pt;border-right:solid #000000 1.0pt;background:#000000;padding:0cm 5.4pt 0cm 5.4pt'><p class=MsoNormal><b><span style='color:green'><o:p>$ConnectDuration</o:p></span></b></p></td>
        
"@
        
        $body += @"
        
      </tr>
"@
    }
    $body += @"
    </table>
"@

    
 
    $body += @"

        </table>
"@
    $stopwatch.Stop()
    $runtime = $stopwatch.ElapsedMilliseconds / 60000
    $runtime = [math]::Round($runtime, 2)
    
    $generated = Get-Date -Format "dd-MM-yyyy hh:mm:ss"
     

    #End Body
    $body += @"
      </p>
    </div>
    <font style=font-family:"Consolas"; size="1">
    <br>Generated:$generated
    </font>
    <!--- <button type="submit"  onClick="refreshPage()">Andrew's I cant wait!!!</button> ---!>
  </body>
  </html>
"@
    #HTML Body complete
    $reportBody = $head.ToString() + $body.ToSTring()
    #Write to file
    $reportBody | Out-File $webFile
}


$statusDetail = @{}

$statusDetail = Get-VPNUserIP
New-HTMLBody($statusDetail)
