param ([string]$info="Password")

$jiraPATId="PAT Id"



$pair = "$(""):$($jiraPATId)"
$headers = @{
    Authorization = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($pair)"))
}
function Get-WorkItemInfo {
    $WorkItemInfo = New-Object -TypeName PSObject
    $WorkItemInfo| Add-Member -MemberType NoteProperty -Name AssignedTo -Value $null
    $WorkItemInfo| Add-Member -MemberType NoteProperty -Name EMail -Value $null
    # $WorkItemInfo| Add-Member -MemberType NoteProperty -Name Issueid -Value $null
    $WorkItemInfo| Add-Member -MemberType NoteProperty -Name Timespent -Value $null
    # $WorkItemInfo| Add-Member -MemberType NoteProperty -Name Status -Value $null
    return $WorkItemInfo
    
}
function Get-EpicInfo {
    $EpicInfo = New-Object -TypeName PSObject
    $EpicInfo| Add-Member -MemberType NoteProperty -Name ParentID -Value $null
    $EpicInfo| Add-Member -MemberType NoteProperty -Name summary -Value $null
    $EpicInfo| Add-Member -MemberType NoteProperty -Name ChildTask -Value $null
    $EpicInfo| Add-Member -MemberType NoteProperty -Name Timespent -Value $null
    return $EpicInfo
    
}
function Get-WorkItemUpdate {
    $WorkItemUpdate = New-Object -TypeName PSObject
    $WorkItemUpdate| Add-Member -MemberType NoteProperty -Name AssignedTo -Value $null
    $WorkItemUpdate| Add-Member -MemberType NoteProperty -Name EMail -Value $null
    $WorkItemUpdate| Add-Member -MemberType NoteProperty -Name Issueid -Value $null
    $WorkItemUpdate| Add-Member -MemberType NoteProperty -Name Timespent -Value $null
    $WorkItemUpdate| Add-Member -MemberType NoteProperty -Name Status -Value $null
    $WorkItemUpdate| Add-Member -MemberType NoteProperty -Name ParenID -Value $null
    $WorkItemUpdate| Add-Member -MemberType NoteProperty -Name ParentSummary -Value $null
    $WorkItemUpdate| Add-Member -MemberType NoteProperty -Name IssueSummary -Value $null
    return $WorkItemUpdate
    
}
function New-SprintDetails
{
    $SprintDetails = New-Object -TypeName PSObject
    $SprintDetails | Add-Member -MemberType NoteProperty -Name ID -Value $null
    $SprintDetails | Add-Member -MemberType NoteProperty -Name Name -Value $null
    $SprintDetails | Add-Member -MemberType NoteProperty -Name sprintNumber -Value $null
    $SprintDetails | Add-Member -MemberType NoteProperty -Name startDate -Value $null
    $SprintDetails | Add-Member -MemberType NoteProperty -Name endDate -Value $null
    $SprintDetails | Add-Member -MemberType NoteProperty -Name status -Value $null
    
    return $SprintDetails
}

$WorkitemDetailsHtml +="<tbody><table border='1' cellspacing='1' style='font-size:10px; border:1px solid black; border-collapse: collapse' nowrap>
        <tr style='font-size:12px;'>
            <th valign='top' nowrap><b>Issue ID</b></th>
            <th valign='top' nowrap><b>Issue Summary</b></th>
            <th valign='top' nowrap><b>Issue Type</b></th>
            <th valign='top' nowrap><b>Status</b></th>
            <th valign='top' nowrap><b>Peoject Key</b></th>
            <th valign='top' nowrap><b>Project Name</b></th>
            <th valign='top' nowrap><b>Assigned To</b></th>
            <th valign='top' nowrap><b>Created By</b></th>
            <th valign='top' nowrap><b>Created Date</b></th>
        
        </tr>"
$WorkitemDetailsHtml11 +="<tbody><table border='1' cellspacing='1' style='font-size:10px; border:1px solid black; border-collapse: collapse' nowrap>
        <tr style='font-size:12px;'>
            <th valign='top' nowrap><b>Issue ID</b></th>
            <th valign='top' nowrap><b>Issue Summary</b></th>
            <th valign='top' nowrap><b>Issue Type</b></th>
            <th valign='top' nowrap><b>Status</b></th>
            <th valign='top' nowrap><b>Peoject Key</b></th>
            <th valign='top' nowrap><b>Project Name</b></th>
            <th valign='top' nowrap><b>Assigned To</b></th>
            <th valign='top' nowrap><b>Created By</b></th>
            <th valign='top' nowrap><b>Created Date</b></th>
        
        </tr>"
$quarterlyWorkitemDetailsHtml +="<tbody><table border='1' cellspacing='1' style='font-size:10px; border:1px solid black; border-collapse: collapse' nowrap>
        <tr style='font-size:12px;'>
            <th valign='top' nowrap><b>Issue ID</b></th>
            <th valign='top' nowrap><b>Issue Summary</b></th>
            <th valign='top' nowrap><b>Issue Type</b></th>
            <th valign='top' nowrap><b>Status</b></th>
            <th valign='top' nowrap><b>Peoject Key</b></th>
            <th valign='top' nowrap><b>Project Name</b></th>
            <th valign='top' nowrap><b>Assigned To</b></th>
            <th valign='top' nowrap><b>Created By</b></th>
            <th valign='top' nowrap><b>Created Date</b></th>
        
        </tr>"
$quarterlyWorkitemDetailsHtml11 +="<tbody><table border='1' cellspacing='1' style='font-size:10px; border:1px solid black; border-collapse: collapse' nowrap>
        <tr style='font-size:12px;'>
            <th valign='top' nowrap><b>Issue ID</b></th>
            <th valign='top' nowrap><b>Issue Summary</b></th>
            <th valign='top' nowrap><b>Issue Type</b></th>
            <th valign='top' nowrap><b>Status</b></th>
            <th valign='top' nowrap><b>Peoject Key</b></th>
            <th valign='top' nowrap><b>Project Name</b></th>
            <th valign='top' nowrap><b>Assigned To</b></th>
            <th valign='top' nowrap><b>Created By</b></th>
            <th valign='top' nowrap><b>Created Date</b></th>
        
        </tr>"
# # $WorkitemsQueryUri = "https://$($organization).atlassian.net/rest/api/2/project/EMAD"
# $WorkitemsQueryUri = "https://$($organization).atlassian.net/rest/api/2/dashboard"
# $WorkitemsQueryUri = "https://$($organization).atlassian.net/rest/api/2/issue/EMAD-2699"
# https://$($organization).atlassian.net/rest/agile/1.0/sprint/520
$workflowuri= "https://$($organization).atlassian.net/rest/api/2/search?jql=ORDER%20BY%20Created&maxResults=100"
$workflowResult = Invoke-RestMethod -Uri $workflowuri -Method Get -Headers $headers
$Totalcount = ($workflowResult.total)/100
$currentdate = Get-Date
$currentmonth = $currentdate.Month
$lastmonth = $currentmonth-1
$quarterstartmonth = $currentmonth-3
for ($j = 0; $j -le $Totalcount; $j++) {
    $startAt = $j*100
    $EMADworkflowuri= "https://$($organization).atlassian.net/rest/api/2/search?jql=ORDER%20BY%20Created&maxResults=100&startAt=$($startAt)"
    $EMADworkflowResult = Invoke-RestMethod -Uri $EMADworkflowuri -Method Get -Headers $headers
    $EMAD = $EMADworkflowResult.issues|Where-Object{$_.fields.assignee.emailAddress -match ""}
    for ($i = 0; $i -lt $EMAD.Count; $i++) {
        $issueid = $EMAD[$i].key
        # $timespentinHours = $latestworklogs.fields.timetracking.timeSpent
        $issuestatusuri ="https://$($organization).atlassian.net/rest/api/2/issue/$($issueid)"
        $issuestatusResult = Invoke-RestMethod -Uri $issuestatusuri -Method Get -Headers $headers
        # $Parentinfo=Get-parentInfo -issueid $issueid
        $issuestatus = $EMAD | Where-Object{$_.key -match $issueid }
        # $assignedto = $issuestatus.fields.assignee.emailAddress
        $person = $issuestatus.fields.assignee.displayName
        $status = $issuestatus.fields.status.name
        $issuesummary = $issuestatusResult.fields.summary
        $issuetype = $issuestatusResult.fields.issuetype.name
        $projectName = $issuestatusResult.fields.project.name
        $projectkey = $issuestatusResult.fields.project.key
        $createdby = $issuestatusResult.fields.creator.displayName
        $createdtime = [datetime]$issuestatusResult.fields.created
        # $parentid = $Parentinfo.ID
        # $parentsummary = $Parentinfo.summary
        if ($createdtime.Month -eq $lastmonth) {
            $WorkitemDetailsHtml += "<tr style='font-size:12px;' nowrap>
                <td nowrap>&nbsp;$($issueid)&nbsp;</td>
                <td align ='left' nowrap>&nbsp;$($issuesummary)&nbsp;</td> 
                <td align ='left' nowrap>&nbsp;$($issuetype)&nbsp;</td> 
                <td align ='left' nowrap>&nbsp;$($status)&nbsp;</td>
                <td align ='left' nowrap>&nbsp;$($projectkey)&nbsp;</td>
                <td align ='left' nowrap>&nbsp;$($projectName)&nbsp;</td>
                <td align ='left' nowrap>&nbsp;$($person)&nbsp;</td>
                <td align ='left' nowrap>&nbsp;$($createdby)&nbsp;</td>
                <td align ='left' nowrap>&nbsp;$($createdtime)&nbsp;</td>
            </tr>" 
        }
        $querterMonth= 3,6,9,12
        if ($querterMonth -contains $currentmonth) {
            if ($createdtime.Month -ge $quarterstartmonth -and $createdtime.Month -le $lastmonth) {
                $quarterlyWorkitemDetailsHtml += "<tr style='font-size:12px;' nowrap>
                    <td nowrap>&nbsp;$($issueid)&nbsp;</td>
                    <td align ='left' nowrap>&nbsp;$($issuesummary)&nbsp;</td> 
                    <td align ='left' nowrap>&nbsp;$($issuetype)&nbsp;</td> 
                    <td align ='left' nowrap>&nbsp;$($status)&nbsp;</td>
                    <td align ='left' nowrap>&nbsp;$($projectkey)&nbsp;</td>
                    <td align ='left' nowrap>&nbsp;$($projectName)&nbsp;</td>
                    <td align ='left' nowrap>&nbsp;$($person)&nbsp;</td>
                    <td align ='left' nowrap>&nbsp;$($createdby)&nbsp;</td>
                    <td align ='left' nowrap>&nbsp;$($createdtime)&nbsp;</td>
                </tr>" 
            }
        }
                
    }
    $Assignees = ("")
    foreach($Assignee in $Assignees){
        $Synechron = $EMADworkflowResult.issues|Where-Object{$_.fields.assignee.emailAddress -match $Assignee}
        for ($i = 0; $i -lt $Synechron.Count; $i++) {
            $issueid = $Synechron[$i].key
            # $timespentinHours = $latestworklogs.fields.timetracking.timeSpent
            $issuestatusuri ="https://$($organization).atlassian.net/rest/api/2/issue/$($issueid)"
            $issuestatusResult = Invoke-RestMethod -Uri $issuestatusuri -Method Get -Headers $headers
            # $Parentinfo=Get-parentInfo -issueid $issueid
            $issuestatus = $Synechron | Where-Object{$_.key -match $issueid }
            # $assignedto = $issuestatus.fields.assignee.emailAddress
            $person = $issuestatus.fields.assignee.displayName
            $status = $issuestatus.fields.status.name
            $issuesummary = $issuestatusResult.fields.summary
            $issuetype = $issuestatusResult.fields.issuetype.name
            $projectName = $issuestatusResult.fields.project.name
            $projectkey = $issuestatusResult.fields.project.key
            $createdby = $issuestatusResult.fields.creator.displayName
            $createdtime = [datetime]$issuestatusResult.fields.created
            # $parentid = $Parentinfo.ID
            # $parentsummary = $Parentinfo.summary
            if ($createdtime.Month -eq $lastmonth) {
                $WorkitemDetailsHtml11 += "<tr style='font-size:12px;' nowrap>
                    <td nowrap>&nbsp;$($issueid)&nbsp;</td>
                    <td align ='left' nowrap>&nbsp;$($issuesummary)&nbsp;</td> 
                    <td align ='left' nowrap>&nbsp;$($issuetype)&nbsp;</td> 
                    <td align ='left' nowrap>&nbsp;$($status)&nbsp;</td>
                    <td align ='left' nowrap>&nbsp;$($projectkey)&nbsp;</td>
                    <td align ='left' nowrap>&nbsp;$($projectName)&nbsp;</td>
                    <td align ='left' nowrap>&nbsp;$($person)&nbsp;</td>
                    <td align ='left' nowrap>&nbsp;$($createdby)&nbsp;</td>
                    <td align ='left' nowrap>&nbsp;$($createdtime)&nbsp;</td>
                </tr>" 
            }
            $querterMonth= 3,6,9,12
            $q = 0
            if ($querterMonth -contains $currentmonth) {
                if ($createdtime.Month -ge $quarterstartmonth -and $createdtime.Month -le $lastmonth) {
                    $quarterlyWorkitemDetailsHtml11 += "<tr style='font-size:12px;' nowrap>
                        <td nowrap>&nbsp;$($issueid)&nbsp;</td>
                        <td align ='left' nowrap>&nbsp;$($issuesummary)&nbsp;</td> 
                        <td align ='left' nowrap>&nbsp;$($issuetype)&nbsp;</td> 
                        <td align ='left' nowrap>&nbsp;$($status)&nbsp;</td>
                        <td align ='left' nowrap>&nbsp;$($projectkey)&nbsp;</td>
                        <td align ='left' nowrap>&nbsp;$($projectName)&nbsp;</td>
                        <td align ='left' nowrap>&nbsp;$($person)&nbsp;</td>
                        <td align ='left' nowrap>&nbsp;$($createdby)&nbsp;</td>
                        <td align ='left' nowrap>&nbsp;$($createdtime)&nbsp;</td>
                    </tr>" 
                    $q++
                }
            }
                    
        }
    }
    
}
$month = (Get-Culture).DateTimeFormat.GetMonthName($lastmonth) 
    $WorkitemDetailsHtml += "</table></tbody>"
    $WorkitemDetailsHtml11 += "</table></tbody>"
    $WorkitemDetailsHtml1 += "<h1>Report for AWS OPS Team for the month of $($month)</h1>"
    $WorkitemDetailsHtml1 += "<br>"
    $quarterlyWorkitemDetailsHtml1 += "<h1>Report for AWS OPS Team for the Last quarter</h1>"
    $quarterlyWorkitemDetailsHtml1 += "<br>"
    # $quarterlyWorkitemDetailsHtml1 += "<h3>This report is filtered with Last month only</h3>"
    $WorkitemDetailsHtml12 += "<h1>Report for AWS Platform Team for the month of $($month)</h1>"
    $WorkitemDetailsHtml12 += "<br>"
    $quarterlyWorkitemDetailsHtml12 += "<h1>Report for AWS Platform Team for the last quarter</h1>"
    $quarterlyWorkitemDetailsHtml12 += "<br>"
    # $quarterlyWorkitemDetailsHtml12 += "<h3>This report is filtered with Last month only</h3>"
    
    $body += $WorkitemDetailsHtml1
    $body += $WorkitemDetailsHtml
    $body2 += $WorkitemDetailsHtml12
    $body2 += $WorkitemDetailsHtml11
    $body3 += $quarterlyWorkitemDetailsHtml1
    $body3 += $quarterlyWorkitemDetailsHtml
    $body4 += $quarterlyWorkitemDetailsHtml12
    $body4 += $quarterlyWorkitemDetailsHtml11
    

    # # compose and send out email messages to individuals who have booked to at least one task in the sprint
    Write-Host "-------------------------------------------------------------------------------------------"
    Write-Host "Sending out email for " $person
    $email=("")
    $emailSubject = "Work Items Summary report for Synechron Team (AWS Operations Team)"
    Start-Sleep -s 15

    $SMTP_SERVER = "smtp.socketlabs.com"
    $SMTP_PORT = 587
    $SMTP_USERNAME = "server4507"
    $PWORD = ConvertTo-SecureString -String $info -AsPlainText -Force
    $bccEmailList = ("")
    # $email = ("")
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
        $message.cc.Add($mailid)
    }
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
        $WorkitemDetailsHtml = $null
        $WorkitemDetailsHtml1 = $null
       
    }
    catch {
        $_.Exception.message
    }
    $body = $null
    $WorkitemDetailsHtml = $null
    $WorkitemDetailsHtml1 = $null
    # Stop-Transcript
    $emailSubject1 = "Work Items Summary report for Synechron Team (AWS Platform Team)"
    $SMTPClient1 = New-Object System.Net.Mail.SmtpClient
    $SMTPClient1.Host = $SMTP_SERVER
    $SMTPClient1.Port = $SMTP_PORT
    $SMTPClient1.EnableSsl = $true
    $SMTPClient1.Credentials = New-Object System.Net.NetworkCredential($SMTP_USERNAME, $PWORD)
    $message1= New-Object System.Net.Mail.MailMessage
    $message1.From = ""
    foreach($mailid in $bccEmailList)
    {
        $message1.cc.Add($mailid)
    }
    $message1.To.Add($email)
    $message1.Subject = $emailSubject1
    $message1.Body = $body2
    # $message.cc.Add($bccEmailList)
    $message1.IsBodyHtml = $true
    # Send-MailMessage -SmtpServer $SMTP_SERVER -Port $SMTP_PORT -Credential $cred -To $email -Bcc $bccEmailList -From tfsbuild@apx.com -Subject $emailSubject -Body $body -BodyAsHtml
    try {
        # Send-MailMessage -To $email -SmtpServer $SMTP_SERVER -Port $SMTP_PORT -Credential $cred -Bcc $bccEmailList -From tfsbuild@apx.com -Subject $emailSubject -Body $body -BodyAsHtml
        $SMTPClient1.Send($message1)
        $SMTPClient1.Dispose()
        $message1.Dispose()
        $body2 = $null
        $WorkitemDetailsHtml11 = $null
        $WorkitemDetailsHtml12 = $null
       
    }
    catch {
        $_.Exception.message
    }
    $body2 = $null
    $WorkitemDetailsHtml11 = $null
    $WorkitemDetailsHtml12 = $null
    # Stop-Transcript
$querterMonth= 1,4,6,10
if ($querterMonth -contains $currentmonth) {
    # # compose and send out email messages to individuals who have booked to at least one task in the sprint
    Write-Host "-------------------------------------------------------------------------------------------"
    Write-Host "Sending out email for " $person

    $emailSubject21 = "Quarterly Work Items Summary report for Synechron Team (AWS Operations Team)"
    Start-Sleep -s 15
    # $email = ("")
    # $cred = New-Object -TypeName System.Management.Automation.PSCredential($SMTP_USERNAME, $PWORD)
    $SMTPClient21 = New-Object System.Net.Mail.SmtpClient
    $SMTPClient21.Host = $SMTP_SERVER
    $SMTPClient21.Port = $SMTP_PORT
    $SMTPClient21.EnableSsl = $true
    $SMTPClient21.Credentials = New-Object System.Net.NetworkCredential($SMTP_USERNAME, $PWORD)
    $message21= New-Object System.Net.Mail.MailMessage
    $message21.From = ""
    foreach($mailid in $bccEmailList)
    {
        $message21.cc.Add($mailid)
    }
    $message21.To.Add($email)
    $message21.Subject = $emailSubject21
    $message21.Body = $body3
    # $message.cc.Add($bccEmailList)
    $message21.IsBodyHtml = $true
    # Send-MailMessage -SmtpServer $SMTP_SERVER -Port $SMTP_PORT -Credential $cred -To $email -Bcc $bccEmailList -From tfsbuild@apx.com -Subject $emailSubject -Body $body -BodyAsHtml
    try {
        # Send-MailMessage -To $email -SmtpServer $SMTP_SERVER -Port $SMTP_PORT -Credential $cred -Bcc $bccEmailList -From tfsbuild@apx.com -Subject $emailSubject -Body $body -BodyAsHtml
        $SMTPClient21.Send($message21)
        $SMTPClient21.Dispose()
        $message21.Dispose()
        $body3 = $null
        $quarterlyWorkitemDetailsHtml = $null
        $quarterlyWorkitemDetailsHtml1 = $null
   
    }
    catch {
        $_.Exception.message
    }
    $body3 = $null
    $quarterlyWorkitemDetailsHtml = $null
    $quarterlyWorkitemDetailsHtml1 = $null
    # Stop-Transcript
    $emailSubject22 = "Quarterly Work Items Summary report for Synechron Team (AWS Platform Team)"
    $SMTPClient22 = New-Object System.Net.Mail.SmtpClient
    $SMTPClient22.Host = $SMTP_SERVER
    $SMTPClient22.Port = $SMTP_PORT
    $SMTPClient22.EnableSsl = $true
    $SMTPClient22.Credentials = New-Object System.Net.NetworkCredential($SMTP_USERNAME, $PWORD)
    $message22= New-Object System.Net.Mail.MailMessage
    $message22.From = ""
    foreach($mailid in $bccEmailList)
    {
        $message22.cc.Add($mailid)
    }
    $message22.To.Add($email)
    $message22.Subject = $emailSubject22
    $message22.Body = $body4
    # $message.cc.Add($bccEmailList)
    $message22.IsBodyHtml = $true
    # Send-MailMessage -SmtpServer $SMTP_SERVER -Port $SMTP_PORT -Credential $cred -To $email -Bcc $bccEmailList -From "" -Subject $emailSubject -Body $body -BodyAsHtml
    try {
        # Send-MailMessage -To $email -SmtpServer $SMTP_SERVER -Port $SMTP_PORT -Credential $cred -Bcc $bccEmailList -From "" -Subject $emailSubject -Body $body -BodyAsHtml
        $SMTPClient22.Send($message22)
        $SMTPClient22.Dispose()
        $message22.Dispose()
        $body4 = $null
        $quarterlyWorkitemDetailsHtml11 = $null
        $quarterlyWorkitemDetailsHtml12 = $null
   
    }
    catch {
        $_.Exception.message
    }
    $body4 = $null
    $quarterlyWorkitemDetailsHtml11 = $null
    $quarterlyWorkitemDetailsHtml12 = $null
    # Stop-Transcript
}
