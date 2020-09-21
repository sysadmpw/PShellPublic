<#
.SYNOPSIS
  Get last nights daily backup and backup copy syncs, create report and send via email.
.DESCRIPTION
  Report on previous nights VM backups, Windows Agent (physical) backups, copy results and summarise failure reasons.
  Configured as scheduled task, if using a different VEEAM server replace value of $veeamServer to your target VBR server.
  Schedule the report after all VM and Windows agent backups are completed.

  This script is using an SMTP server to send the report via email, adjust Send-VBReport SMTP variables as required.
.INPUTS
  None.
.OUTPUTS
  Email containing formatted list of jobs, last status of the protection group, and summary of errors by protection group object (server).
.NOTES
  Version:        1.2
  Author:         Peter Wahid
  Creation Date:  05/07/19
  Purpose/Change: N/A
  
.EXAMPLE
  Just execute the script via CLI or configure a scheduled task on the source to execute the script.
#>
#Estimate script runtime
$stopwatch = [system.diagnostics.stopwatch]::StartNew()
$stopwatch.start()

Clear-Host
#---------------------------------------------------------[Variables to adjust] ------------------------------------------------------------------------#
$veeamServer = "server.local"

#-----------------------------------------------------------[Functions] --------------------------------------------------------------------------------#
#Build HTML Email and send.
Function New-HTMLEmail($BackupDaily, $BackupCopy, $BackupAgentJobs, $ReportTime, $DailyReasons) {
    $temp = 0
 
    Write-host "Building Email" -ForegroundColor Yellow
    #Header, add CSS style here
    $head = @"
  <html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" xmlns="http://www.w3.org/TR/REC-html40"><head><META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=us-ascii"><meta name=Generator content="Microsoft Word 15 (filtered medium)"><style><!--
    /* Font Definitions */
    @font-face
	    {font-family:"Cambria Math";
	    panose-1:2 4 5 3 5 4 6 3 2 4;}
    @font-face
	    {font-family:Calibri;
	    panose-1:2 15 5 2 2 2 4 3 2 4;}
    /* Style Definitions */
    p.MsoNormal, li.MsoNormal, div.MsoNormal
	    {margin:0cm;
	    margin-bottom:.0001pt;
	    font-size:11.0pt;
	    font-family:"Calibri",sans-serif;
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
	    font-family:"Calibri",sans-serif;
	    color:windowtext;}
    .MsoChpDefault
	    {mso-style-type:export-only;
	    font-family:"Calibri",sans-serif;
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
    </o:shapelayout></xml><![endif]--></head>
"@
  
  

    #Start Body, add data from VBR Sessions here
    $body = @"
  <body lang=EN-AU link="#0563C1" vlink="#954F72">
    <div class=WordSection1>
      <p class=MsoNormal>
"@

#Daily Backup Data to Table
    $body += @"
     <H1>VEEAM Backup and Replication Report</H1><br>
     <H2><u>Hyper-V</u></H2>
     <table class=MsoTable15Grid4Accent6 border=1 cellspacing=0 cellpadding=0 style='border-collapse:collapse;border:none'>
       <tr>
       <td width="400" valign="top" style="width:300.8pt;border:solid #70AD47 1.0pt;border-right:none;background:#70AD47;padding:0cm 5.4pt 0cm 5.4pt"><p class="MsoNormal"><b><span style="color:white"><o:p>Name</o:p></span></b></p></td>
       <td width="400" valign="top" style="width:300.8pt;border:solid #70AD47 1.0pt;border-right:none;background:#70AD47;padding:0cm 5.4pt 0cm 5.4pt"><p class="MsoNormal"><b><span style="color:white"><o:p>Start</o:p></span></b></p></td>
       <td width="400" valign="top" style="width:300.8pt;border:solid #70AD47 1.0pt;border-right:none;background:#70AD47;padding:0cm 5.4pt 0cm 5.4pt"><p class="MsoNormal"><b><span style="color:white"><o:p>End</o:p></span></b></p></td>
       <td width="400" valign="top" style="width:300.8pt;border:solid #70AD47 1.0pt;border-right:none;background:#70AD47;padding:0cm 5.4pt 0cm 5.4pt"><p class="MsoNormal"><b><span style="color:white"><o:p>Size GB</o:p></span></b></p></td>
       <td width="400" valign="top" style="width:300.8pt;border:solid #70AD47 1.0pt;border-right:none;background:#70AD47;padding:0cm 5.4pt 0cm 5.4pt"><p class="MsoNormal"><b><span style="color:white"><o:p>Result</o:p></span></b></p></td>
       </tr>
       <tr>
"@
    ForEach ($dailyitem in $BackupDaily) {
      

        [string]$dailyname = $dailyitem.Name
      
        #VBR returns start and end times in US Format
        #converting the Start and End objects to String/DateTime seems to allow the local parser to format the date time
        [string]$startTime = $dailyitem.Start.ToString()
        [string]$endTime = $dailyitem.End.ToString()

        [string]$dailystart = $startTime
        [string]$dailyend = $endTime
        [string]$dailysizegb = $dailyitem.SizeGB
        [string]$dailyresult = $dailyitem.Result
      
        $body += @"
                   <td width="400" valign="top" style="width:300.8pt;border:solid #A8D08D 1.0pt;border-top:none;background:#E2EFD9;padding:0cm 5.4pt 0cm 5.4pt"><p class="MsoNormal"><b><o:p>$dailyname</o:p></b></p></td>
                   <td width="400" valign="top" style="width:300.8pt;border:solid #A8D08D 1.0pt;border-top:none;background:#E2EFD9;padding:0cm 5.4pt 0cm 5.4pt"><p class="MsoNormal"><b><o:p>$dailystart</o:p></b></p></td>
                   <td width="400" valign="top" style="width:300.8pt;border:solid #A8D08D 1.0pt;border-top:none;background:#E2EFD9;padding:0cm 5.4pt 0cm 5.4pt"><p class="MsoNormal"><b><o:p>$dailyend</o:p></b></p></td>
                   <td width="400" valign="top" style="width:300.8pt;border:solid #A8D08D 1.0pt;border-top:none;background:#E2EFD9;padding:0cm 5.4pt 0cm 5.4pt"><p class="MsoNormal"><b><o:p>$dailysizegb GB</o:p></b></p></td>
"@
        If ($dailyresult -eq "Success") {
            $body += @"
                   <td width="400" valign="top" style="width:300.8pt;border:solid #A8D08D 1.0pt;border-top:none;background:#E2EFD9;padding:0cm 5.4pt 0cm 5.4pt"><p class="MsoNormal"><b><o:p>$dailyresult</o:p></b></p></td>
                 </tr>
"@    
        }
        elseif ($dailyresult -eq "Warning") {
            $body += @"
                   <td width="400" valign="top" style="width:300.8pt;border:solid #A8D08D 1.0pt;border-top:none;background:#f5da42;padding:0cm 5.4pt 0cm 5.4pt"><p class="MsoNormal"><b><o:p>$dailyresult</o:p></b></p></td>
                 </tr>
"@
        }
        elseif ($dailyresult -eq "Failed") {
            $body += @"
                   <td width="400" valign="top" style="width:300.8pt;border:solid #A8D08D 1.0pt;border-top:none;background:#f54242;padding:0cm 5.4pt 0cm 5.4pt"><p class="MsoNormal"><b><o:p>$dailyresult</o:p></b></p></td>
                 </tr>
"@        
        }else{
          $body += @"
          <td width="400" valign="top" style="width:300.8pt;border:solid #A8D08D 1.0pt;border-top:none;background:#f54242;padding:0cm 5.4pt 0cm 5.4pt"><p class="MsoNormal"><b><o:p>CHECK VBR CONSOLE</o:p></b></p></td>
        </tr>
"@         
        }
    }


    $body += @"
     </table>
"@

#Agent (Physical) Backups

    $body += @"
<br><br><br><H2><u>Windows Agent</u></H2>
<table class=MsoTable15Grid4Accent6 border=1 cellspacing=0 cellpadding=0 style='border-collapse:collapse;border:none'>
  <tr>
  <td width="400" valign="top" style="width:300.8pt;border:solid #70AD47 1.0pt;border-right:none;background:#70AD47;padding:0cm 5.4pt 0cm 5.4pt"><p class="MsoNormal"><b><span style="color:white"><o:p>Name</o:p></span></b></p></td>
  <td width="400" valign="top" style="width:300.8pt;border:solid #70AD47 1.0pt;border-right:none;background:#70AD47;padding:0cm 5.4pt 0cm 5.4pt"><p class="MsoNormal"><b><span style="color:white"><o:p>Start</o:p></span></b></p></td>
  <td width="400" valign="top" style="width:300.8pt;border:solid #70AD47 1.0pt;border-right:none;background:#70AD47;padding:0cm 5.4pt 0cm 5.4pt"><p class="MsoNormal"><b><span style="color:white"><o:p>End</o:p></span></b></p></td>
  <td width="400" valign="top" style="width:300.8pt;border:solid #70AD47 1.0pt;border-right:none;background:#70AD47;padding:0cm 5.4pt 0cm 5.4pt"><p class="MsoNormal"><b><span style="color:white"><o:p>Size GB</o:p></span></b></p></td>
  <td width="400" valign="top" style="width:300.8pt;border:solid #70AD47 1.0pt;border-right:none;background:#70AD47;padding:0cm 5.4pt 0cm 5.4pt"><p class="MsoNormal"><b><span style="color:white"><o:p>Result</o:p></span></b></p></td>
  </tr>
  <tr>
"@
    ForEach ($agentitem in $BackupAgentJobs) {
        $session = $agentitem.FindLastSession()
        if ($null -eq $session) {
            $temp += 1
            #Write-host ""$session.JobName" NULL Session found : $temp" -ForegroundColor Red
        }
        else {
            #Only add data to table if the last session for a Windows Agent backup is not Null
            $transferSize = ($session.Info.Progress.TransferedSize / 1GB -as [int])
            [string]$agentname = $session.JobName
            $agentstart = $session.CreationTime.tostring()
            $agentend = $session.EndTime.tostring()
            [string]$agentsizegb = $transferSize
            [string]$agentresult = $session.info.Result
                  
            #VBR returns start and end times in US Format
            #converting the Start and End objects to String/DateTime seems to allow the local parser to format the date time
            $body += @"
              <td width="400" valign="top" style="width:300.8pt;border:solid #A8D08D 1.0pt;border-top:none;background:#E2EFD9;padding:0cm 5.4pt 0cm 5.4pt"><p class="MsoNormal"><b><o:p>$agentname</o:p></b></p></td>
              <td width="400" valign="top" style="width:300.8pt;border:solid #A8D08D 1.0pt;border-top:none;background:#E2EFD9;padding:0cm 5.4pt 0cm 5.4pt"><p class="MsoNormal"><b><o:p>$agentstart</o:p></b></p></td>
              <td width="400" valign="top" style="width:300.8pt;border:solid #A8D08D 1.0pt;border-top:none;background:#E2EFD9;padding:0cm 5.4pt 0cm 5.4pt"><p class="MsoNormal"><b><o:p>$agentend</o:p></b></p></td>
              <td width="400" valign="top" style="width:300.8pt;border:solid #A8D08D 1.0pt;border-top:none;background:#E2EFD9;padding:0cm 5.4pt 0cm 5.4pt"><p class="MsoNormal"><b><o:p>$agentsizegb GB</o:p></b></p></td>
"@
            If ($agentresult -eq "Success") {
                $body += @"
                      <td width="400" valign="top" style="width:300.8pt;border:solid #A8D08D 1.0pt;border-top:none;background:#E2EFD9;padding:0cm 5.4pt 0cm 5.4pt"><p class="MsoNormal"><b><o:p>$agentresult</o:p></b></p></td>
                    </tr>
"@    
            }
            elseif ($agentresult -eq "Warning") {
                $body += @"
                    <td width="400" valign="top" style="width:300.8pt;border:solid #A8D08D 1.0pt;border-top:none;background:#f5da42;padding:0cm 5.4pt 0cm 5.4pt"><p class="MsoNormal"><b><o:p>$agentresult</o:p></b></p></td>
                  </tr>
"@
            }
            elseif ($agentresult -eq "Failed") {
                $body += @"
                  <td width="400" valign="top" style="width:300.8pt;border:solid #A8D08D 1.0pt;border-top:none;background:#f54242;padding:0cm 5.4pt 0cm 5.4pt"><p class="MsoNormal"><b><o:p>$agentresult</o:p></b></p></td>
                </tr>
"@        
            }
        }
    }


    $body += @"


</table>
"@

#Daily Sync Copy Data to Table   
    $body += @"
     <br><br><br><u><H2> Copy Sessions </H2></u>
     <table class=MsoTable15Grid4Accent6 border=1 cellspacing=0 cellpadding=0 style='border-collapse:collapse;border:none'>
       <tr>
       <td width="400" valign="top" style="width:300.8pt;border:solid #70AD47 1.0pt;border-right:none;background:#70AD47;padding:0cm 5.4pt 0cm 5.4pt"><p class="MsoNormal"><b><span style="color:white"><o:p>Name</o:p></span></b></p></td>
       <td width="400" valign="top" style="width:300.8pt;border:solid #70AD47 1.0pt;border-right:none;background:#70AD47;padding:0cm 5.4pt 0cm 5.4pt"><p class="MsoNormal"><b><span style="color:white"><o:p>Size GB</o:p></span></b></p></td>
       <td width="400" valign="top" style="width:300.8pt;border:solid #70AD47 1.0pt;border-right:none;background:#70AD47;padding:0cm 5.4pt 0cm 5.4pt"><p class="MsoNormal"><b><span style="color:white"><o:p>Result</o:p></span></b></p></td>
       </tr>
       <tr>
"@
    ForEach ($copyitem in $BackupCopy) {
        [string]$copyname = $copyitem.Name
        [string]$copysizegb = $copyitem.SizeGB
        [string]$copyresult = $copyitem.Result
        $body += @"
                   <td width="400" valign="top" style="width:300.8pt;border:solid #A8D08D 1.0pt;border-top:none;background:#E2EFD9;padding:0cm 5.4pt 0cm 5.4pt"><p class="MsoNormal"><b><o:p>$copyname</o:p></b></p></td>
                   <td width="400" valign="top" style="width:300.8pt;border:solid #A8D08D 1.0pt;border-top:none;background:#E2EFD9;padding:0cm 5.4pt 0cm 5.4pt"><p class="MsoNormal"><b><o:p>$copysizegb GB</o:p></b></p></td>
"@
        If ($copyresult -eq "Success") {
            $body += @"
                   <td width="400" valign="top" style="width:300.8pt;border:solid #A8D08D 1.0pt;border-top:none;background:#E2EFD9;padding:0cm 5.4pt 0cm 5.4pt"><p class="MsoNormal"><b><o:p>$copyresult</o:p></b></p></td>
                 </tr>
"@    
        }
        elseif ($copyresult -eq "Warning") {
            $body += @"
                   <td width="400" valign="top" style="width:300.8pt;border:solid #A8D08D 1.0pt;border-top:none;background:#f5da42;padding:0cm 5.4pt 0cm 5.4pt"><p class="MsoNormal"><b><o:p>$copyresult</o:p></b></p></td>
                 </tr>
"@
        }
        elseif ($copyresult -eq "Failed") {
            $body += @"
                   <td width="400" valign="top" style="width:300.8pt;border:solid #A8D08D 1.0pt;border-top:none;background:#f54242;padding:0cm 5.4pt 0cm 5.4pt"><p class="MsoNormal"><b><o:p>$copyresult</o:p></b></p></td>
                 </tr>
"@        
        }
    }


    $body += @"


     </table>
     <br>
"@

# Failure Reasonns to Table
    $body += @"
<br><br><br><H2><u>Failure Summary</u></H2>
<h5> Experimental: Reasons include retries.</h5>
<table class=MsoTable15Grid4Accent6 border=1 cellspacing=0 cellpadding=0 style='border-collapse:collapse;border:none'>
  <tr>
  <td width="400" valign="top" style="width:300.8pt;border:solid #ad0a0a 1.0pt;border-right:none;background:#e0523f;padding:0cm 5.4pt 0cm 5.4pt"><p class="MsoNormal"><b><span style="color:white"><o:p>Source Names</o:p></span></b></p></td>
  <td width="400" valign="top" style="width:300.8pt;border:solid #ad0a0a 1.0pt;border-right:none;background:#e0523f;padding:0cm 5.4pt 0cm 5.4pt"><p class="MsoNormal"><b><span style="color:white"><o:p>Failure Reason</o:p></span></b></p></td>
  </tr>
  <tr>
"@
    $goodservers = 0
    #Loop through each session and retrieve the final objects result/reason for failing. Only report on non null values, where null values are completed backups
    $DailyReasons = $DailyReasons | Sort-Object -Property ObjectName
    ForEach ($reasonItem in $DailyReasons) {
        $reasonName = $reasonItem.ObjectName
        $reasonDetail = $reasonItem.Reason
        if($reasonItem.Reason){
            $body += @"
            <td width="400" valign="top" style="width:300.8pt;border:solid #ad0a0a 1.0pt;border-top:none;background:#E2EFD9;padding:0cm 5.4pt 0cm 5.4pt"><p class="MsoNormal"><b><o:p>$reasonName</o:p></b></p></td>
            <td width="400" valign="top" style="width:300.8pt;border:solid #ad0a0a 1.0pt;border-top:none;background:#E2EFD9;padding:0cm 5.4pt 0cm 5.4pt"><p class="MsoNormal"><b><o:p>$reasonDetail</o:p></b></p></td>
  </tr>
"@
        }else{
            $goodservers += 1
        }
     }
    $body += @"
</table>
"@

    #End Body
    $body += @"
    </p>
   </div>
   <h4>Estimated runtime $ReportTime min</h4>
  </body>
  </html>
"@
 
    #Send Email
    $emailBody = $head.ToString() + $body.ToString()
    Write-Host "Sending email...." -ForegroundColor Yellow

    Send-VBReport($emailBody)

}

#Send the email with formatted body
Function Send-VBReport($email) {

    $au = New-Object System.Globalization.CultureInfo("en-AU")
    $reportDate = Get-Date -Format ($au.DateTimeFormat.ShortDatePattern)
    #SMTP Server Details  
    $ALERTsmtp = "smtp.server" 
    $ALERTto = "toemail@to.com.au" 
    $ALERTsubject = "$veeamServer VBR Report $reportDate"
    $ALERTfrom = "fromemail@domain.com.au"
    $smtpPort = "2525"
  
    #String from Build HTML Email
    $EmailBody = $email
    Send-MailMessage -To $ALERTto -Subject $ALERTsubject -From $ALERTfrom -Body $EmailBody -BodyAsHtml -SmtpServer $ALERTsmtp -Priority "High" -Port $smtpPort
}

#Windows Agent (Physical) Jobs
Function Get-VBRWindowsAgentData() {
    $BackupAgentJobs = [Veeam.Backup.Core.CBackupJob]::GetAll() | Where-Object { $_.JobType -eq "EpAgentBackup" }
    
    return $BackupAgentJobs
}

#Hyper-V Backup Jobs
Function Get-DailyBackupResults() {
    $BackupResult = Get-VBRBackupSession | Where-Object { $_.CreationTime -ge (Get-Date).AddDays(-1) } | Where-Object { $_.JobType -eq "Backup" }
    $ResultDaily = $BackupResult | Select-Object Name, @{Label = "Start"; Expression = { ($_.CreationTime) } }, @{Label = "End"; Expression = { ($_.EndTime) } }, @{Name = "SizeGB"; Expression = { [math]::round($_.BackupStats.BackupSize / 1GB, 2) } }, @{Label = "Result"; Expression = { ($_.Result) } } | Sort-Object Start -CaseSensitive

    return $ResultDaily
}

#Copy Jobs
Function Get-BackupCopyResults() {
  
    $BackupCopyResult = Get-VBRJob | Where-Object { $_.JobType -eq "BackupSync" }
    $ResultCopy = $BackupCopyResult | Select-Object Name, @{Name = "SizeGB"; Expression = { [math]::round($_.Info.IncludedSize / 1GB, 2) } }, @{Label = "Result"; Expression = { ($_.Info.LatestStatus) } } | Sort-Object Name -CaseSensitive

    return $ResultCopy
}

#Reasons for last nights failures, and the failed objects
Function Get-VBRBackupSessionReasons(){

    $allSessionResults = Get-VBRBackupSession | Where-Object { $_.CreationTime -ge (Get-Date).AddDays(-1) }
    $report = @()
    foreach($session in $allSessionResults){
      $info  = [Veeam.Backup.Core.CBackupTaskSession]::GetByJobSession($session.id)
      $report += $info
    }
    return $report 
}



#----------------------------------------------------------- [Execution] ------------------------------------------------------------------------#
#For now, execution must follow the below sequence

#Open VBR server session
Write-Host "Connecting to VEEAM Server" -ForegroundColor Green
Connect-VBRServer -Server $veeamServer

#Collect Hyper-V Result
Write-Host "Collecting backup results" -ForegroundColor Yellow
$BackupDailyResult = Get-DailyBackupResults

#Collect Windows Agent Result
Write-host "Collecting Windows Agent Backup Data" -ForegroundColor Yellow
$BackupWindowAgent = Get-VBRWindowsAgentData

#Collect Copy Results
Write-Host "Collecting copy results" -ForegroundColor Yellow
$CopyResult = Get-BackupCopyResults

#Collect Session Reasons
Write-Host "Collecting Session Reasons" -ForegroundColor Yellow
$BackupReasons = Get-VBRBackupSessionReasons

#Script processing runtime, not including building and sending email report
$stopwatch.Stop()
$runtime = $stopwatch.ElapsedMilliseconds / 60000
$runtime = [math]::Round($runtime, 2)

#Build and send report via Email
New-HTMLEmail -BackupDaily $BackupDailyResult -BackupCopy $CopyResult -BackupAgentJobs $BackupWindowAgent -ReportTime $runtime -DailyReasons $BackupReasons
[System.GC]::Collect()

#Close VBR server session
Write-Host "Disconnecting $veeamServer session" -ForegroundColor Red
Disconnect-VBRServer