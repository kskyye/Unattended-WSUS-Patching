#2017 08 21
#Written by: Kristin Anderson
#This script was intended to run as a scheduled task to apply deadlines on patches to server groups.


# Load .NET assembly
[reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration") | Out-Null

#Define variables
$count = 0
$reportDate = Get-Date -format "yyyy.MM.dd"
$todaysDate = Get-Date
$startofmonth = Get-Date $todaysDate -day 1 -hour 0 -minute 0 -second 0
$newPatches = [datetime]$startofmonth
$resultsLog = "C:\WSUS Reports\Declined Updates\$reportDate - WSUS Declined Updates.txt"
$email = new-object Net.Mail.MailMessage

# Connect to WSUS Server
$wsus = [Microsoft.UpdateServices.Administration.AdminProxy]::getUpdateServer(‘ADMIN-WSUS’,$False,"8530")
$updateScope = New-Object Microsoft.UpdateServices.Administration.UpdateScope
$updateScope.FromArrivalDate = $newPatches
<#$updateClassifications = $wsus.GetUpdateClassifications() | Where {
  $_.Title -Match "Critical Updates|Security Updates|Updates|Update Rollups"
}
$UpdateScope.Classifications.AddRange($updateClassifications)#>
$updates=$wsus.GetUpdates($updateScope)

$flag = 0


foreach ($update in $updates)
{
    #if ($update.IsSuperseded -eq ‘True’)
    if (($update.IsSuperseded -eq ‘True’) -and ($update.IsDeclined -ne ‘True’))    
    {
        $logUpdate = $($Update.Title) 
        $logUpdate | Add-Content $resultsLog
        $update.Decline()
        $count=$count + 1
        $flag = 1
    }

    #if ($update.Title -like "*Preview*")
    if (($update.Title -like "*Preview*") -and ($update.IsDeclined -ne ‘True’))
    {
        $logUpdate = $($Update.Title) 
        $logUpdate | Add-Content $resultsLog
        $update.Decline()
        $count=$count + 1
        $flag = 1
    }

    if (($update.Title -like "*LanguageFeatureOnDemand*") -or ($update.Title -like "*LanguageInterfacePack*") -or ($update.Title -like "*Language Pack*") -or ($update.Title -like "*LanguagePack*") -or ($update.Title -like "*Lang Pack*") -and ($update.IsDeclined -ne ‘True’))
    {
        $logUpdate = $($Update.Title) 
        $logUpdate | Add-Content $resultsLog
        $update.Decline()
        $count=$count + 1
        $flag = 1
    }
}
If ($flag -eq 1)
{
ren "C:\WSUS Reports\Declined Updates\$reportDate - WSUS Declined Updates.txt" "C:\WSUS Reports\Declined Updates\$reportDate - WSUS Declined Updates (total updates - $count updates).txt"
    $file = "C:\WSUS Reports\Declined Updates\$reportDate - WSUS Declined Updates (total updates - $count updates).txt"
    $attachment = new-object Net.Mail.Attachment($file)
    $email.Attachments.Add($attachment)
}

#Define variables for email and send report
$month = (Get-Culture).DateTimeFormat.GetMonthName((Get-Date).Month.ToString())
$smtpServer = "10.200.1.25"
$smtp = new-object Net.Mail.SmtpClient($smtpServer, 25)
$report = $null
$FromEmail = "WSUS Reports@sufs.org"
$ToEmail = "IT_Infrastructure@sufs.org"
$email.From = new-object Net.Mail.MailAddress($FromEmail)
$email.Priority = [System.Net.Mail.MailPriority]::High
$email.IsBodyHtml = $true
$email.Body = $report
$email.Subject = "WSUS: $month - WSUS Patching Schedule & Summary"
$email.To.Add($ToEmail)

$todaytesDateShort = $todaysDate = Get-Date $todaysDate -format "MM/dd/yy"
$dateWrks = ((Get-Date).adddays(+(0+(Get-Date -UFormat %u)))).date
$dateWrksShort = Get-Date $dateWrks -format "MM/dd/yy"
$dateSvrGrpA = ((Get-Date).adddays(+(7+(Get-Date -UFormat %u)))).date
$dateSvrGrpAShort = Get-Date $dateSvrGrpA -format "MM/dd/yy"
$dateSvrGrpB = ((Get-Date).adddays(+(14+(Get-Date -UFormat %u)))).date
$dateSvrGrpBShort = Get-Date $dateSvrGrpB -format "MM/dd/yy"
$dateSvrGrpC = ((Get-Date).adddays(+(21+(Get-Date -UFormat %u)))).date
$dateSvrGrpCShort = Get-Date $dateSvrGrpC -format "MM/dd/yy"
$startOfMonthShort = Get-Date $startOfMonth -format "MM/dd"

# HTML Style Definition
$report += "<style type='text/css'>"
$report += ".tg  {border-collapse:collapse;border-spacing:0;border-color:#999;}"
$report += ".tg td{font-family:Arial, sans-serif;font-size:14px;padding:10px 5px;border-style:solid;border-width:0px;overflow:hidden;word-break:normal;border-color:#999;color:#444;background-color:#F7FDFA;border-top-width:1px;border-bottom-width:1px;}"
$report += ".tg th{font-family:Arial, sans-serif;font-size:14px;font-weight:normal;padding:10px 5px;border-style:solid;border-width:0px;overflow:hidden;word-break:normal;border-color:#999;color:#fff;background-color:#26ADE4;border-top-width:1px;border-bottom-width:1px;}"
$report += ".tg .tg-wqx0{font-weight:bold;font-size:16px;font-family:Verdana, Geneva, sans-serif !important;;vertical-align:top}"
$report += ".tg .tg-c3rp{font-size:16px;font-family:Verdana, Geneva, sans-serif !important;;vertical-align:top}"
$report += ".tg .tg-gg7v{font-weight:bold;font-size:16px;font-family:Verdana, Geneva, sans-serif !important;;text-align:right;vertical-align:top}"
$report += ".tg .tg-jua3{font-weight:bold;font-family:Verdana, Geneva, sans-serif !important;;text-align:right;vertical-align:top}"
$report += "</style>"
$report += "<table class='tg'>"
$report +=   "<tr>"
$report +=     "<th class='tg-wqx0' colspan='2'>Updates Summary for $month</th>"
$report +=   "</tr>"
$report +=   "<tr>"
$report +=     "<td class='tg-gg7v'>New Updates:</td>"
$report +=     "<td class='tg-c3rp'>$($Updates.count) (from arrival date of $startOfMonthShort) </td>"
$report +=   "</tr>"
$report +=   "<tr>"
$report +=     "<td class='tg-gg7v'>Patches Declined:</td>"
$report +=     "<td class='tg-c3rp'>$count (report attached if applicable)</td>"
$report +=   "</tr>"
$report +=   "<tr>"
$report +=     "<th class='tg-wqx0' colspan='2'>Patching Schedule</th>"
$report +=   "</tr>"
$report +=   "<tr>"
$report +=     "<td class='tg-gg7v'>$todaysDate </td>"
$report +=     "<td class='tg-c3rp'>Microsoft Patch Tuesday.</td>"
$report +=   "</tr>"
$report +=   "<tr>"
$report +=     "<td class='tg-gg7v'>$dateWrksShort </td>"
$report +=     "<td class='tg-c3rp'>Workstations.</td>"
$report +=   "</tr>"
$report +=   "<tr>"
$report +=     "<td class='tg-gg7v'>$dateSvrGrpAShort </td>"
$report +=     "<td class='tg-c3rp'>Server Group A (DEV &amp; low risk ADMIN servers.)</td>"
$report +=   "</tr>"
$report +=   "<tr>"
$report +=     "<td class='tg-gg7v'>$dateSvrGrpBShort </td>"
$report +=     "<td class='tg-c3rp'>Server Group B (Low risk PROD &amp; remaining ADMIN servers.)</td>"
$report +=   "</tr>"
$report +=   "<tr>"
$report +=     "<td class='tg-gg7v'>$dateSvrGrpCShort </td>"
$report +=     "<td class='tg-c3rp'>Server Group C (Remaining PROD servers &amp; WSUS server.)</td>"
$report +=   "</tr>"
$report +=   "<tr>"
$report +=     "<td class='tg-gg7v'>&nbsp;</td>"
$report +=     "<td class='tg-c3rp'>&nbsp;</td>"
$report +=   "</tr>"
$report +=   "<tr>"
$report +=     "<td class='tg-gg7v'>Notes:</td>"
$report +=     "<td class='tg-c3rp'>Servers will install patches and reboot between 11PM and 5AM.  Workstations will begin patching when connected to the network.</td>"
$report +=   "</tr>"
$report += "</table>"

$email.body = $report
$smtp.Send($email)

If ($flag -eq 1)
{
del "C:\Scripts\WSUS\$reportDate - WSUS Deadline Report ($($Updates.count) updates).csv"#>
}