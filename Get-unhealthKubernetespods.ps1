param ([string]$info="Please provide password")
<#

Project Reporting script for VSTS

This script sends emails to all developers, testers, and product managers notifying them
of Open workitems list and send a gentle reminder to the Workitem owner and respective leads to say hey from these Many days your workitem is pending



#>
Add-Type -AssemblyName System.Web

# Get the script directory
$path= "C:\Users\kjnk\common\common.ps1"
# # Load and execute the common library script
$commonLibContent = Get-Content -Path $Path -Raw
Invoke-Expression -Command $commonLibContent
$date = Get-Date
$utcDate = ([datetime]$date).ToUniversalTime()
$currentDateTime = $utcDate.ToString("dd-MM-yyyy hh:mm:ss tt")
$report += "<br>"
$report += "<b>Date/Time (UTC): &nbsp;</b>$currentDateTime`r`n"
$report += "<br>"
$report += "<h2>Kubernetes pods listed below are those that, based on Azure Log Analytics, have been in Failed state for more than 4 hours</h2>"
$report += "<br>"
$report += "<tbody><table border='1' cellspacing='1' style='font-size:10px; border:1px solid black; border-collapse: collapse' nowrap>
        <tr style='font-size:12px;'>
            <th valign='top' nowrap><b>Service name</b></th>
            <th valign='top' nowrap><b>Namespace</b></th>
            <th valign='top' nowrap><b>Pod Name</b></th>
            <th valign='top' nowrap><b>Pod State</b></th>
            <th valign='top' nowrap><b>Pod Creation Date (UTC)</b></th>
        </tr>"
$WorkspaceID="Workspace ID"
$query = 'c'
# | where PodStatus !contains "Running"
$InsightsQuery = Invoke-AzOperationalInsightsQuery -WorkspaceId $WorkspaceID -Query $query
$InsightsQueryResult = $InsightsQuery.Results
$serviceNames = $InsightsQueryResult.serviceName|Sort-Object |Select-Object -Unique
foreach($serviceName in $serviceNames){
    $NameSpaces = $InsightsQueryResult | Where-Object {$_.serviceName -eq $serviceName}|Sort-Object -Property Namespace
    $podresults = $NameSpaces | Where-Object {$_.PodStatus -ne "Running"}
    foreach($Item in $podresults){
        $serviceNamedata = $Item.serviceName
        $NameSpace = $Item.Namespace
        $PodName = $Item.Name
        $PodState = $Item.PodStatus
        $PodCreationDate = $Item.PodCreationTimeStamp
        $color = "#FF0000"
        $report += "<tr style='font-size:12px;' bgcolor=$color nowrap>
        <td align ='left' nowrap>&nbsp;$($serviceNamedata)&nbsp;</a></td>
        <td align ='left' nowrap>&nbsp;$($NameSpace)&nbsp;</td> 
        <td align ='left' nowrap>&nbsp;$($PodName)&nbsp;</td> 
        <td align ='left' nowrap>&nbsp;$($PodState)&nbsp;</td>
        <td align ='left' nowrap>&nbsp;$($PodCreationDate)&nbsp;</td>
    </tr>" 

    }

}
$report += "</table></tbody>"
$body += $report
# # compose and send out email messages to individuals who have booked to at least one task in the sprint
Write-Host "-------------------------------------------------------------------------------------------"
Write-Host "Sending out email " 
$email=("")
$emailSubject = "Unhealthy Kubernetes Pods in non-Prod Environments"
Start-Sleep -s 15

$SMTP_SERVER = "smtp.socketlabs.com"
$SMTP_PORT = 587
$SMTP_USERNAME = "server4507"
$PWORD = ConvertTo-SecureString -String $info -AsPlainText -Force
$bccEmailList = ("")

# $cred = New-Object -TypeName System.Management.Automation.PSCredential($SMTP_USERNAME, $PWORD)
$SMTPClient = New-Object System.Net.Mail.SmtpClient
$SMTPClient.Host = $SMTP_SERVER
$SMTPClient.Port = $SMTP_PORT
$SMTPClient.EnableSsl = $true
$SMTPClient.Credentials = New-Object System.Net.NetworkCredential($SMTP_USERNAME, $PWORD)
$message= New-Object System.Net.Mail.MailMessage
$message.From = ""
foreach($mailid in $bccEmailList)
{
    $message.Bcc.Add($mailid)
}
# foreach($email in $emails){
#     $message.To.Add($email)
# }
# foreach($mail in $mails){
#     $message.CC.Add($mail)
# }
$message.To.Add($email)
$message.Subject = $emailSubject
$message.Body = $body
# $message.cc.Add($bccEmailList)
$message.IsBodyHtml = $true
# Send-MailMessage -SmtpServer $SMTP_SERVER -Port $SMTP_PORT -Credential $cred -To $email -Bcc $bccEmailList -From tfsbuild@apx.com -Subject $emailSubject -Body $body -BodyAsHtml
try {
    # Send-MailMessage -To $email -SmtpServer $SMTP_SERVER -Port $SMTP_PORT -Credential $cred -Bcc $bccEmailList -From tfsbuild@apx.com -Subject $emailSubject -Body $body -BodyAsHtml
    $SMTPClient.Send($message)
    $SMTPClient.Dispose()
    $message.Dispose()
    $body = $null
    $report = $null
    $serviceNames = $null   
}
catch {
    $_.Exception.message
}


