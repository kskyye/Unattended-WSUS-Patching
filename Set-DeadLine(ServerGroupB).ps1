#2017 08 16
#Written by: Kristin Anderson
#Script will set deadlines for Failed or Needed Windows updates on WSUS server groups.  This script was intended to run as a scheduled task.

#Initialize connection to server via .NET Accelerator
[reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration") | Out-Null

#Define variables
#'wsus' variable can be modified to connect to a different WSUS server and port
$wsus = [Microsoft.UpdateServices.Administration.AdminProxy]::getUpdateServer(‘ADMIN-WSUS’,$False,"8530")
#'group' variable can be changed to target other Computer groups defined in WSUS
$group = 'Server Group B'
$todaysDate = (Get-Date).AddDays(+2)
$reportDate = Get-Date -format "yyyy.MM.dd"
#String suffix of 'date' variable can be changed to set deadline at a different time of day.  The day the script runs is managed from the Windows Scheduled Task.
$date = $todaysDate.ToShortDateString()+" 03:00AM"
$deadline = [datetime]$date
$ComputerTargetGroups = $wsus.GetComputerTargetGroups()
#'UpdateScope' variables can be changed to select different sets of patches that the deadline will be set on.
$UpdateScope = New-Object Microsoft.UpdateServices.Administration.UpdateScope
$UpdateScope.ApprovedStates = 'LatestRevisionApproved'
$UpdateScope.IncludedInstallationStates = 'NotInstalled'
$UpdateScope.IncludedInstallationStates = 'Downloaded'
$updateClassifications = $wsus.GetUpdateClassifications() | Where {
  $_.Title -Match "Critical Updates|Security Updates"
}

$UpdateScope.Classifications.AddRange($updateClassifications)

#Set computer scope to match computer group configured in WSUS (variable defined above);
$ComputerScope = New-Object Microsoft.UpdateServices.Administration.ComputerTargetScope
$TargetGroup = $ComputerTargetGroups | Where {$_.Name -eq $Group}
[void]$computerscope.ComputerTargetGroups.Add($TargetGroup)

$Clients = $WSUS.GetComputerTargets($computerscope)
$Updates = $Clients.GetUpdateInstallationInfoPerUpdate($UpdateScope) | Select -Unique -ExpandProperty UpdateID


ForEach ($Item in $Updates) {    
    #Set update variable and write to CSV log
    $Update = $wsus.GetUpdate([guid]$Item)
    #Write update details to log file
    $UpdateLog = $wsus.GetUpdate($Item) | Export-Csv -Path "C:\WSUS Reports\Monthly Patching Deadlines\$reportDate - WSUS Deadline Report ($($Updates.count) updates).csv" -Append -NoTypeInformation
    
    Write-Host "Setting deadline for" $($Update.Title) "on" $group "for" $date "Update ID: " $Item

    #Set deadline on patch
    $Update.Approve("Install",$Targetgroup,$deadLine)
}



#Define variables for email and send report
$smtpServer = "10.200.1.25"
$smtp = new-object Net.Mail.SmtpClient($smtpServer, 25)
$report = $null
$FromEmail = "WSUS Reports@sufs.org"
$ToEmail = "IT_Infrastructure@sufs.org"
$email = new-object Net.Mail.MailMessage
$email.From = new-object Net.Mail.MailAddress($FromEmail)
$file = "C:\WSUS Reports\Monthly Patching Deadlines\$reportDate - WSUS Deadline Report ($($Updates.count) updates).csv"
$attachment = new-object Net.Mail.Attachment($file)
$email.Attachments.Add($attachment)
$email.Priority = [System.Net.Mail.MailPriority]::High
$email.IsBodyHtml = $true
$email.Body = $report
$email.Subject = "WSUS: $month - $($Updates.count) Updates Applied to $($TargetGroup.name)"
$email.To.Add($ToEmail)
$report += "The $($Updates.count) Windows updates listed in the attached report have been approved and a deadline has been set for $($TargetGroup.name).  Servers will install patches and reboot between 12AM and 5AM Wednesday"
$email.body = $report
$smtp.Send($email)
$attachment.Dispose()

#"C:\Scripts\WSUS\$reportDate - Windows Updates Report - $($TargetGroup.name): ($($Updates.count) updates).csv"