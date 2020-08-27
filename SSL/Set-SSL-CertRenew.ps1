<#
.SYNOPSIS
    Get latst SSL Let's Encrypt PFX from a location and install and bind to an IIS website.
.DESCRIPTION
    This script will grab the latest SSL certificiate from PFX location $global:pfxFile, install the certificate, bind to the new certificate and finally remove the expired certificate. If required restart IIS. Note, at this stage
    this script is only compatible with IIS based wbsits.
.EXAMPLE
    No CLI, uses scheduled task.
.INPUTS
    serverlist.json - see global variables
.OUTPUTS
    - Event Viewer logging.
    - Teams message on critical error.
.NOTES
    General notes
    EventID Numbers and explanations - #Error, Information, FailureAudit, SuccessAudit, Warning
    1. Information, things that are happening or about to happen  1000 - Entry Type Information
    2. Errors 1001 - Entry Type Error
    3. Successes 1002 - Entry Type SuccessAudit
    4. Warning 1003 - Entry Type Warning

#>

#Global Variables
$global:ProgressPreference = 'SilentlyContinue'
#Location of PFX file
$global:pfxFile = $null
$global:pfxPass = ""
$global:pfxFileHash = Get-FileHash $global:pfxFile

#Thumbprint will always be the same for a PFX file used across multiple servers
$global:pfxFileThumbprint = Get-PfxCertificate -FilePath $global:pfxFile | Select-Object Thumbprint

#Link to a json file with list of IP addresses of target IIS servers
$global:serverListLocation = $null
$global:serverList = Get-Content -Raw -Path $global:serverListLocation | ConvertFrom-Json
$global:srvCertStatus = $null
$global:installedCerts = $null

#WebHook Variables for Teams Messaging, convert to module and remove from initial build.
#URI of channel for WebHook
$global:channel = $null
$global:teamsMessage = $null
$global:teamTitle = $null
$global:teamSummary = $null
$global:teamSubTitle = $null
$global:teamActivityImage = $null
$global:teamStatusRED = "FC0000"
$global:teamStatusGRN = "00FC00"
$global:teamStatusORNG = "FC7E00"
$global:teamStatus = $null
$global:teamNameTXT = $null
$global:teamValueTXT = $null
$global:teamSRVName = $null
$global:teamSRVTSTHTTPS = $null
$global:teamSSLStatus = $null
$global:teamSSLStatusOK = "Current"
$global:teamSSLStatusNO = "Expired"

#Build the message body. Write your own messages as you please.
$global:teamsMessage = $null
$global:teamTitle = $null
$global:teamSubTitle = Get-Date -Format "HH:mm:ss dd/MM/yyyy"
$global:teamSummary = $null
#You need this image in your life.
$global:teamActivityImage = "https://i.ibb.co/99TvQNt/webhook.png"
$global:teamStatusRED = "FC0000"
$global:teamStatusGRN = "00FC00"
$global:teamStatusORNG = "FC7E00"
$global:teamStatus = $null #Set this to RED GRN ORNG variables
$global:teamStatusTXT = $null

#IIS Binding Configuration
$global:sslBindPath = $null
$global:sslUnassignedPath = "IIS:\SSLBindings\0.0.0.0!443"


#Encrypted credentials.
$global:winRMPassContent = $null
$global:winRMAESContent = $null
$global:winrMUName = $null
$global:winRMPass = Get-Content $global:winRMPassContent | ConvertTo-SecureString -Key (Get-Content $global:winRMAESContent)
$global:winRMCredential = New-Object System.Management.Automation.PsCredential($global:winrMUName, $global:winRMPass)

#Event Source messages
$global:eventSrcMsg = $null
$global:LogName = $null

#Servers that are not reachable on 443
$global:srvDown = @()


#Confirm event logs are configured properly to write events on localhost
function Set-EventSourceStatus {

    #Check if the LetsEncryptRenew event source does not exist, then create one.
    $eventSourceStatus = [System.Diagnostics.EventLog]::SourceExists("LetsEncryptRenew");
    if ($eventSourceStatus -ne $True) {
        New-EventLog -LogName $global:LogName -Source "LetsEncryptRenew"
    }
}


#Read json file of server list, test the ssl certificate for each server
function Get-ServersToCheck {
    foreach ($srv in $global:serverList) {
        #Check if 443 is responding
        $alive = Test-NetConnection -Port 443 $srv -InformationLevel Quiet
        $hostname = [System.Net.Dns]::GetHostEntry("$srv").HostName

        #Where the real action happens.
        if (($alive)) {
            #Check SSL certificate status for Let's Encrypt, return true requires a renew or false does not.
            $expiredCertThumbprint = Test-SSLCertExpiry $srv

            if ($false -eq $expiredCertThumbprint) {
                $global:eventSrcMsg = "Could not read SSL certificate repo on $hostname"
                Write-EventLog -LogName "$global:LogName" -Source "LetsEncryptRenew" -EventID 1001 -EntryType Error -Message $global:eventSrcMsg -Category 1 -RawData 10,20
                $global:eventSrcMsg = $null

                $global:teamSRVTSTHTTPS = $hostname
                $global:teamStatusTXT = "Could not read SSL certificate repo on $hostname"
                Send-TeamsMessage
            }
            else {
                if ($null -eq $expiredCertThumbprint) {
                    $global:eventSrcMsg = "No SSL change required on $hostname"
                    Write-EventLog -LogName "$global:LogName" -Source "LetsEncryptRenew" -EventID 1002 -EntryType SuccessAudit -Message $global:eventSrcMsg -Category 1 -RawData 10,20
                }
                else {
                    $global:eventSrcMsg = "Expired SSL certificate on $hostname, with Thumbprint: $expiredCertThumbprint"
                    Write-EventLog -LogName "$global:LogName" -Source "LetsEncryptRenew" -EventID 1003 -EntryType Warning -Message $global:eventSrcMsg -Category 1 -RawData 10,20

                    $pfxStatus = Test-LatestPFXFile
                    if ($pfxStatus) {

                        $global:eventSrcMsg = "PFX file on $global:LogName-NGINX $pfxStatus"
                        Write-EventLog -LogName "$global:LogName" -Source "LetsEncryptRenew" -EventID 1002 -EntryType SuccessAudit -Message $global:eventSrcMsg -Category 1 -RawData 10,20

                        #Install the PFX file on this server
                        Install-SSLCert $srv $expiredCertThumbprint
                    }
                    else {
                        $global:eventSrcMsg = "PFX file on $global:LogName-NGINX $pfxStatus - Processing terminated, no servers will be updated."
                        Write-EventLog -LogName "$global:LogName" -Source "LetsEncryptRenew" -EventID 1001 -EntryType Error -Message $global:eventSrcMsg -Category 1 -RawData 10,20
                        $global:teamSRVTSTHTTPS = "$global:LogName-NGINX"
                        $global:teamStatusTXT = "Problem with $global:LogName-NGINX PFX file - please investigate. Processing terminated."
                        Send-TeamsMessage
                        Exit
                    }
                }
            }
        }
        else {
            #Create a list of servers that are not responding to report on
            $global:srvDown += $srv
        }
    }
}

#Installs a PFX into IIS
function Install-SSLCert ($updateSRVCert, $expiredSSLThumbprint) {
    $global:eventSrcMsg = "Expired Thumbprint: $expiredSSLThumbprint on $updateSRVCert"
    Write-EventLog -LogName "$global:LogName" -Source "LetsEncryptRenew" -EventID 1001 -EntryType Error -Message $global:eventSrcMsg -Category 1 -RawData 10,20

    #Convert ip to hostname, destination to copy the pfx temporarily to later install and bind.
    $srv = [System.Net.Dns]::GetHostEntry("$updateSRVCert").HostName
    $destinationCopy = "\\$srv\c$\windows\temp"
    
    if (Test-Path $destinationCopy) {
        $global:eventSrcMsg = "Destination copy to: $destinationCopy is available"
        Write-EventLog -LogName "$global:LogName" -Source "LetsEncryptRenew" -EventID 1000 -EntryType Information -Message $global:eventSrcMsg -Category 1 -RawData 10,20
        Copy-Item $global:pfxFile $destinationCopy -Force
        $destHash = Get-FileHash "$destinationCopy\$global:LogName.com.au.pfx"
        if ($destHash.Hash -eq $global:pfxFileHash.Hash) {
            $global:eventSrcMsg = "Source and destination PFX file hash match, copy is good. Importing PFX file into certificate repository."
            Write-EventLog -LogName "$global:LogName" -Source "LetsEncryptRenew" -EventID 1000 -EntryType Information -Message $global:eventSrcMsg -Category 1 -RawData 10,20

            Invoke-Command -ComputerName $updateSRVCert -ScriptBlock { Import-PFXCertificate -CertStoreLocation Cert:\LocalMachine\My\ -FilePath C:\windows\Temp\NAMEOFYOURPFX.com.au.pfx } -Credential $global:winRMCredential
            #Bind new PFX to HTTPS website
            $tempthumb = $global:pfxFileThumbprint.Thumbprint
            $global:eventSrcMsg = "Binding new SSL Cert: $tempthumb to IIS."
            Write-EventLog -LogName "$global:LogName" -Source "LetsEncryptRenew" -EventID 1000 -EntryType Information -Message $global:eventSrcMsg -Category 1 -RawData 10,20

    #Check type of binding on target IIS server
            $global:sslBindPath = "IIS:\SSlBindings\$updateSRVCert!443"
            $scriptBlockParams = @{
                ComputerName = $updateSRVCert
                ScriptBlock = { Param ($param1) Import-Module WebAdministration ; Test-Path $param1 }
                Credential = $global:winRMCredential
                ArgumentList = "$global:sslBindPath"

            }
            $resultAssigned = Invoke-Command @scriptBlockParams

            #Check if IIS Site not Assigned an IP
            $global:sslBindPath = "IIS:\SSlBindings\0.0.0.0!443"
            $scriptBlockParams = @{
                ComputerName = $updateSRVCert
                ScriptBlock = { Param ($param1) Import-Module WebAdministration ; Test-Path $param1 }
                Credential = $global:winRMCredential
                ArgumentList = "$global:sslBindPath"

            }
            $resultUnAssigned = Invoke-Command @scriptBlockParams



            #If true, bind by IP, else bind by 0.0.0.0
            if($resultAssigned){
                $global:eventSrcMsg = "Server $updateSRVCert IIS Site is Assigned."
                Write-EventLog -LogName "CMRI" -Source "LetsEncryptRenew" -EventID 1000 -EntryType Information -Message $global:eventSrcMsg -Category 1 -RawData 10,20

                $currentPFX = $global:pfxFileThumbprint.Thumbprint
                $scriptBlockParamsBind = @{
                    ComputerName = $updateSRVCert
                    ScriptBlock  = { Param ($param1Bind, $param2Bind) Import-Module WebAdministration ; Get-ChildItem Cert:\LocalMachine\My\$param1Bind | Set-Item "IIS:\SSlBindings\$param2Bind!443" }
                    Credential   = $global:winRMCredential
                    ArgumentList = "$currentPFX", "$updateSRVCert"
                }
                Invoke-Command @scriptBlockParamsBind
            }elseif ($resultUnAssigned) {
                $global:eventSrcMsg = "Server $updateSRVCert IIS Site is UnAssigned."
                Write-EventLog -LogName "CMRI" -Source "LetsEncryptRenew" -EventID 1000 -EntryType Information -Message $global:eventSrcMsg -Category 1 -RawData 10,20                
                $currentPFX = $global:pfxFileThumbprint.Thumbprint
                $scriptBlockParamsUBind = @{
                    ComputerName = $updateSRVCert
                    ScriptBlock  = { Param ($param1UBind, $param2UBind) Import-Module WebAdministration ; Get-ChildItem Cert:\LocalMachine\My\$param1UBind | Set-Item "IIS:\SSlBindings\0.0.0.0!443" }
                    Credential   = $global:winRMCredential
                    ArgumentList = "$currentPFX", "$updateSRVCert"
                }
                Invoke-Command @scriptBlockParamsUBind
            }else{
                Write-host "Cannot determine binding, returning null" -ForegroundColor Red
            }


            #Clean up and remove the PFX file from c:\windows\temp
            $global:eventSrcMsg = "Cleaning up and removing temp PFX file from C:\Windows\temp."
            Write-EventLog -LogName "$global:LogName" -Source "LetsEncryptRenew" -EventID 1000 -EntryType Information -Message $global:eventSrcMsg -Category 1 -RawData 10,20
            Remove-Item "$destinationCopy\$global:LogName.com.au.pfx" -Force

            #If success, remove the old SSL Cert
            $global:eventSrcMsg = "Removing old Let's Encrypt certificate from IIS"
            Write-EventLog -LogName "$global:LogName" -Source "LetsEncryptRenew" -EventID 1000 -EntryType Information -Message $global:eventSrcMsg -Category 1 -RawData 10,20            
            #Remove-SSLCert $expiredSSLThumbPrint $updateSRVCert
            $scriptBlockParams = @{
                ComputerName = $updateSRVCert
                ScriptBlock  = { Param ($param1) Import-Module WebAdministration; Remove-Item -Path "Cert:\LocalMachine\My\$param1" }
                Credential   = $global:winRMCredential
                ArgumentList = "$expiredSSLThumbprint"
            }
            Invoke-Command @scriptBlockParams
        }
        else {
            $global:eventSrcMsg = "Source and destination PFX file hash do not match, something went wrong."
            Write-EventLog -LogName "$global:LogName" -Source "LetsEncryptRenew" -EventID 1001 -EntryType Error -Message $global:eventSrcMsg -Category 1 -RawData 10,20
            $global:teamSRVTSTHTTPS = $updateSRVCert
            $global:teamStatusTXT = "PFX hash does not match between source and $updateSRVCert. Please investigate."
            Send-TeamsMessage  
        }
    }
    else {
        $global:eventSrcMsg = "Destination copy to: $destinationCopy not available"
        Write-EventLog -LogName $global:LogName -Source "LetsEncryptRenew" -EventID 1001 -EntryType Error -Message $global:eventSrcMsg -Category 1 -RawData 10,20
        $global:teamSRVTSTHTTPS = $updateSRVCert
        $global:teamStatusTXT = "Unable to copy PFX file to $destinationCopy please investigate."
        Send-TeamsMessage
    }

}

#Check if the Let's Encrypt SSL cert requires renewal.
function Test-SSLCertExpiry ($testSRV) {
    $srvIP = Test-NetConnection -Port 443 $testSRV | Select-Object RemoteAddress
    $hostname = [System.Net.Dns]::GetHostEntry("$testSRV").HostName
    $global:eventSrcMsg = "Getting all SSL certificates on $hostname"
    Write-EventLog -LogName $global:LogName -Source "LetsEncryptRenew" -EventID 1000 -EntryType Information -Message $global:eventSrcMsg -Category 1 -RawData 10,20
    $global:installedCerts = Invoke-Command -ComputerName $srvIP.RemoteAddress -ScriptBlock { Get-ChildItem Cert:\LocalMachine\My } -Credential $global:winRMCredential

    if ($null -eq $global:installedCerts) {
        $global:eventSrcMsg = "Unable to check SSL certificates on $hostname"
        Write-EventLog -LogName $global:LogName -Source "LetsEncryptRenew" -EventID 1001 -EntryType Error -Message $global:eventSrcMsg -Category 1 -RawData 10,20
        $expiredCertThumbPrint = $false
        return $expiredCertThumbPrint
    }
    else {
        foreach ($cert in $global:installedCerts) {
            $certIssuer = $cert.Issuer.ToString()
            $expireAfter = $cert.NotAfter
            if ($certIssuer -eq "CN=Let's Encrypt Authority X3, O=Let's Encrypt, C=US") {
                $expiryDays = (Get-Date).AddDays(30)
                if ($expireAfter -le $expiryDays) {
                    $expiredCertThumbPrint = $cert.Thumbprint
                    return $expiredCertThumbPrint
                }
            } 
        }
    }
    $global:installedCerts = $null
}


#Check if a PFX file exists against the expird PFX files from the server list
function  Test-LatestPFXFile {
    $newSSLCertDate = Get-PfxCertificate -FilePath $global:pfxFile | Select-Object NotAfter
    $newSSLCertThumbPrint = Get-PfxCertificate -FilePath $global:pfxFile | Select-Object Thumbprint

    $expireAfter = $newSSLCertDate.NotAfter
    $expiryDays = (Get-Date).AddDays(30)
    if ($expireAfter -le $expiryDays) {

        $global:eventSrcMsg = "$global:LogName-NGINX SSL Cert is not useable."
        Write-EventLog -LogName "$global:LogName" -Source "LetsEncryptRenew" -EventID 1001 -EntryType Error -Message $global:eventSrcMsg -Category 1 -RawData 10,20
        $global:teamSRVTSTHTTPS = "$global:LogName-NGINX"
        $global:teamStatusTXT = "Problem with PFX file, please investigate"
        Send-TeamsMessage
        return $false
    }
    else {
        $tempthumb = $newSSLCertThumbPrint.Thumbprint
        $global:eventSrcMsg = "New SSL Cert Thumbprint: $tempthumb on Host"
        Write-EventLog -LogName "$global:LogName" -Source "LetsEncryptRenew" -EventID 1002 -EntryType SuccessAudit -Message $global:eventSrcMsg -Category 1 -RawData 10,20
        return $true
    }   
}

#Send process alerts to Teams.
function Send-TeamsMessage {
    $global:teamStatus = $global:teamStatusRED

    #Build the message
    $JSONBody = [PSCustomObject][Ordered]@{
        "@type"      = "MessageCard"
        "@context"   = "http://schema.org/extensions"
        "summary"    = $global:teamSummary
        "themeColor" = $global:teamStatus
        "sections"   = @(
            @{
                "activityTitle"    = $global:teamTitle
                "activitySubtitle" = $global:teamSubTitle
                "activityImage"    = $global:teamActivityImage
                "facts"            = @(
                    @{
                        "name"  = "Web Server:"
                        "value" = $global:teamSRVTSTHTTPS
                    },
                    @{
                        "name"  = "Certificate Status"
                        "value" = $global:teamStatusTXT
                    }
                )
                "markdown"         = $true
            }
        )
    }
    $TeamMessageBody = ConvertTo-Json $JSONBody -Depth 100
    $parameters = @{
        "URI"         = $global:channel
        "Method"      = 'POST'
        "Body"        = $TeamMessageBody
        "ContentType" = 'application/json'
    }
    Invoke-RestMethod @parameters
}


Clear-Host

#Confirm Event viewer is ready to write logs, always triggered first dont change this.
Set-EventSourceStatus

#Function will kick start entire process.
Get-ServersToCheck
