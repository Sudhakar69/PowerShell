param ([string]$info="Password")

$jiraPATId=""
$path1 = "project_reporting_common.ps1"
# Load and execute the common library script
$commonLibContent1 = Get-Content -Path $Path1 -Raw
Invoke-Expression -Command $commonLibContent1


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
function Get-JiraSprintDetails {
    Param ([DateTime]$forDate)
    $workupdates = @()
    for ($i = 500; $i -lt 1000; $i++) {
        try {
            $sprinturi ="https://$($organization).atlassian.net/rest/agile/1.0/sprint/$($i)"
            $sprintresult = Invoke-RestMethod -Method Get -Uri $sprinturi -Headers $headers
            if ($null -ne $sprintresult.psobject.Properties['startdate'] -and $sprintresult.originBoardId -eq "115") {
                $SprintDetails = New-SprintDetails
                $SprintDetails.Name = $sprintresult.name
                $SprintDetails.ID = $sprintresult.id
                $SprintDetails.startDate = $sprintresult.startDate
                $SprintDetails.endDate = $sprintresult.endDate
                $SprintDetails.status = $sprintresult.state
                $workupdates += $SprintDetails
            }
        }
        catch {
        
        }
    

    }
    $forDateUTC = [DateTime]::SpecifyKind($forDate.ToUniversalTime(), [DateTimeKind]::Utc)
    $forDateUTC3 = $forDateUTC.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
    $workupdates |Where-Object{[DateTime]$_.startDate -lt $forDateUTC3 -and [DateTime]$_.endDate -gt $forDateUTC3}
    
}
function Get-parentInfo {
    Param (
        [Parameter(Mandatory=$false)]$issueid
    )
    $issuestatusuri ="https://$($organization).atlassian.net/rest/api/2/issue/$($issueid)"
    $issuestatusResult = Invoke-RestMethod -Uri $issuestatusuri -Method Get -Headers $headers
    $issuelog = $issuestatusResult.fields
    if ($null -ne $issuelog.PSObject.Properties['Parent']) {
        $parentid = $issuelog.parent.key
        Get-parentInfo -issueid $parentid
    }
    else {
        $parentid = $issuestatusResult.key
        $summary = $issuestatusResult.fields.summary
        $parenddata = New-Object -TypeName PSObject
        $parenddata| Add-Member -MemberType NoteProperty -Name ID -Value $parentid
        $parenddata| Add-Member -MemberType NoteProperty -Name summary -Value $summary
       return $parenddata
       
    }
    
    
}
[Net.ServicePointManager]::SecurityProtocol = "tls12"
function Get-EmployeeAvailability {
    $EmployeeAvailability = New-Object -TypeName PSObject
    $EmployeeAvailability| Add-Member -MemberType NoteProperty -Name Name -Value $null
    $EmployeeAvailability| Add-Member -MemberType NoteProperty -Name Type -Value $null
    $EmployeeAvailability| Add-Member -MemberType NoteProperty -Name Date -Value $null
    $EmployeeAvailability| Add-Member -MemberType NoteProperty -Name Hours -Value $null
    return $EmployeeAvailability
    
}
function Get-EmployeeWorkAvailability {
    $EmployeeWorkAvailability = New-Object -TypeName PSObject
    $EmployeeWorkAvailability| Add-Member -MemberType NoteProperty -Name TaskId -Value $null
    $EmployeeWorkAvailability| Add-Member -MemberType NoteProperty -Name TaskName -Value $null
    $EmployeeWorkAvailability| Add-Member -MemberType NoteProperty -Name ParentId -Value $null
    $EmployeeWorkAvailability| Add-Member -MemberType NoteProperty -Name ParentTaskName -Value $null
    $EmployeeWorkAvailability| Add-Member -MemberType NoteProperty -Name WorkedBy -Value $null
    $EmployeeWorkAvailability| Add-Member -MemberType NoteProperty -Name Timespent -Value $null
    return $EmployeeWorkAvailability
    
}

# $name = "Dariusz Orlowski"
$RP_FILEPATH = "https://apxinc.sharepoint.com/sites/PortfolioManagementEMA/Shared%20Documents/General/EMA%20-%20NG%20-%20%20Leave%20Tracker%20and%20On%20Call.xlsx?web=1"
$sheet ="Time Off"
$excel = new-object -com excel.application
# $excel.Visible = $true
$rsWorkbook = $excel.workbooks.open($RP_FILEPATH)
$resourcesSheet = $rsWorkbook.Worksheets.Item($sheet)
$Availability = @()
for ($i = 2; $i -lt $resourcesSheet.UsedRange.Rows.Count; $i++) {
    $name = $resourcesSheet.Range("a$i").text
    $Type= $resourcesSheet.Range("b$i").text
    $date = $resourcesSheet.Range("c$i").text
    $Hours =$resourcesSheet.Range("d$i").text
    # Write-Host $name "::" $Type "::" $date "::" $Hours
        $EmployeeAvailability = Get-EmployeeAvailability
        $EmployeeAvailability.Name = $name
        $EmployeeAvailability.Type = $Type
        $EmployeeAvailability.Date = $date
        $EmployeeAvailability.Hours = $Hours
        $Availability += $EmployeeAvailability
    
}
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
[System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($excel)
$excel.Quit()
Get-Process -Name *excel* |Stop-Process
$latestSprint = Get-JiraSprintDetails -forDate (Get-Date)
$sprintstartdate = [DateTime]$latestSprint.startDate
$sprintenddate = [DateTime]$latestSprint.endDate
$startYear = $sprintstartdate.ToString("yyyy")
$endYear = $sprintenddate.ToString("yyyy")
[Int64]$yearCounter = [Int64]$startYear
while ($yearCounter -le [Int64]$endYear)
{
    $yearHolidayInfoUS = $HOLIDAYS.Item($COUNTRY_US + "|" + $yearCounter)
    $yearHolidayInfoRom = $HOLIDAYS.Item($COUNTRY_ROMANIA + "|" + $yearCounter)
    $yearHolidayInfoPol = $HOLIDAYS.Item($COUNTRY_POLAND + "|" + $yearCounter)

    $holidayInfosUS = $yearHolidayInfoUS.holidayInfos
    foreach ($holidayInfo in $holidayInfosUS)
    {
        if ($holidayInfo.date.DayOfWeek -ne $DAY_OF_WEEK_SATURDAY -and $holidayInfo.date.DayOfWeek -ne $DAY_OF_WEEK_SUNDAY )
        {
            $holidaysUS += ,$holidayInfo
        }
    }

    $holidayInfosRom = $yearHolidayInfoRom.holidayInfos
    foreach ($holidayInfo in $holidayInfosRom)
    {
        if ($holidayInfo.date.DayOfWeek -ne $DAY_OF_WEEK_SATURDAY -and $holidayInfo.date.DayOfWeek -ne $DAY_OF_WEEK_SUNDAY )
        {
            $holidaysRom += ,$holidayInfo
        }
    }

    $holidayInfosPol = $yearHolidayInfoPol.holidayInfos
    foreach ($holidayInfo in $holidayInfosPol)
    {
        if ($holidayInfo.date.DayOfWeek -ne $DAY_OF_WEEK_SATURDAY -and $holidayInfo.date.DayOfWeek -ne $DAY_OF_WEEK_SUNDAY )
        {
            $holidaysRom += ,$holidayInfo
        }
    }

    $yearCounter++  
}
$polandHolidays = $holidayInfosPol |Where-Object{$_.date -ge $sprintstartdate -and $_.date -le $sprintenddate -and $_.isweekend -ne $true}
$holidayscount = $polandHolidays.Count
# # $WorkitemsQueryUri = "https://$($organization).atlassian.net/rest/api/2/project/EMAD"
# $WorkitemsQueryUri = "https://$($organization).atlassian.net/rest/api/2/dashboard"
# $WorkitemsQueryUri = "https://$($organization).atlassian.net/rest/api/2/issue/EMAD-2699"
# https://$($organization).atlassian.net/rest/agile/1.0/sprint/520
$EMADworkflowuri= "https://$($organization).atlassian.net/rest/api/2/search?jql=ORDER%20BY%20Created&maxResults=10000"
$EMADworkflowResult = Invoke-RestMethod -Uri $EMADworkflowuri -Method Get -Headers $headers
$workupdates = @()
$EMAD = $EMADworkflowResult.issues|Where-Object{$_.fields.parent.key -match "EMAD"}
for ($i = 0; $i -lt $EMAD.Count; $i++) {
    $issueid = $EMAD[$i].key
    $issueuri ="https://$($organization).atlassian.net/rest/api/2/issue/$($issueid)/worklog"
    $issueResult = Invoke-RestMethod -Uri $issueuri -Method Get -Headers $headers
    $worklogs = $issueResult.worklogs
    if ($null -ne $worklogs) {    
        $timespentinseconds = 0
        $latestworklogs=$worklogs|Where-Object{[DateTime]$_.updated -gt $latestSprint.startDate}
        for ($j = 0; $j -lt $latestworklogs.Count; $j++) {
            $timespentinseconds += [int]$latestworklogs[$j].timeSpentSeconds
        }
        # $timespentinHours = $latestworklogs.fields.timetracking.timeSpent
        $issuestatusuri ="https://$($organization).atlassian.net/rest/api/2/issue/$($issueid)"
        $issuestatusResult = Invoke-RestMethod -Uri $issuestatusuri -Method Get -Headers $headers
        $Parentinfo=Get-parentInfo -issueid $issueid
        $issuestatus = $EMAD | Where-Object{$_.key -match $issueid}
        if ($null -ne $issuestatus.fields.assignee.displayName) {
            $assignedto = $issuestatus.fields.assignee.emailAddress
            $person = $issuestatus.fields.assignee.displayName
        }
        else {
            $assignedto = $null
            $person = $null
        }
        
        $status = $issuestatus.fields.status.name
        $issuesummary = $issuestatusResult.fields.summary
        $parentid = $Parentinfo.ID
        $parentsummary = $Parentinfo.summary
    }
   
    $WorkItemUpdate = Get-WorkItemUpdate
    $WorkItemUpdate.AssignedTo = $person
    $WorkItemUpdate.EMail = $assignedto
    $WorkItemUpdate.Issueid = $issueid
    $WorkItemUpdate.Timespent =$timespentinseconds
    $WorkItemUpdate.Status = $status
    $WorkItemUpdate.ParenID = $parentid
    $WorkItemUpdate.ParentSummary = $parentsummary
    $WorkItemUpdate.IssueSummary = $issuesummary
    $workupdates += $WorkItemUpdate
    
}
$issues = $workupdates|Group-Object -Property AssignedTo | ForEach-Object {$_.Group}
$uniqueNames = $workupdates|Select-Object -Property AssignedTo -Unique
$WorkInfo = @()
foreach($uniqueName in $uniqueNames.AssignedTo){
    $issuesbyname = $issues |Where-Object{$_.AssignedTo -eq $uniqueName}
    $email = $issuesbyname.EMail |Select-Object -Unique
    $timespentbyseconds =0
    for ($k = 0; $k -lt $issuesbyname.Count; $k++) {
        $timespentbyseconds += [int]$issuesbyname[$k].Timespent
    }
    # $timeinseconds =  [timespan]::fromseconds($timespentbyseconds)
    # $timespent = "$($timeinseconds.hours):$($timeinseconds.minutes):$($timeinseconds.seconds)"
    $timespent = $timespentbyseconds/3600
    $WorkItemInfo = Get-WorkItemUpdate
    $WorkItemInfo.AssignedTo = $uniqueName
    $WorkItemInfo.EMail = $email
    $WorkItemInfo.Issueid = $issuesbyname.issueid
    $WorkItemInfo.Timespent =$timespent
    # $WorkItemInfo.Status = $status
    $workInfo += $WorkItemInfo
    
}
[DateTime]$currentDateTime = Get-Date
$Assignees = $WorkInfo.AssignedTo |Select-Object -Unique
foreach($Assignee in $Assignees){
    
    [DateTime]$startDate=([DateTime]$latestSprint.startDate).ToUniversalTime()
    [DateTime]$endDate=([DateTime]$latestSprint.endDate).ToUniversalTime()
    $today =$currentDateTime.ToUniversalTime()
    $WorkitemSummaryHtml1 = "<h1>$($Assignee)</h1>"
    $WorkitemSummaryHtml1 += "<b>Your Time Booking Since Start of Sprint</b>"
    $WorkitemSummaryHtml1 += "<br><b>As Of Date/Time - UTC: </b>"+ $today.ToString("MM/dd/yyyy hh:mm:ss tt")
    $WorkitemSummaryHtml1 += "<br><b>Poland: </b> " + [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId($today, 'E. Europe Standard Time').toString()
    $WorkitemSummaryHtml1 += "<br><b>Spain/France: </b> " + [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId($today, 'Central Europe Standard Time').toString()
    $WorkitemSummaryHtml1 += "<br><b>UTC: </b>" + $today.ToString("MM/dd/yyyy hh:mm:ss tt")
    $WorkitemSummaryHtml1 += "<br>"
    $WorkitemSummaryHtml1 += "<br><b>Cuurent Sprint: </b>" + $latestSprint.Name
    $WorkitemSummaryHtml1 += "<br><b>Start Date: </b>" + $startDate.ToString("MM/dd/yyyy hh:mm:ss tt")
    $WorkitemSummaryHtml1 += "<br><b>End Date: </b>" + $endDate.ToString("MM/dd/yyyy hh:mm:ss tt")
    $WorkitemSummaryHtml1 += "<br>"
    $WorkitemSummaryHtml += "<br>"
    $WorkitemSummaryHtml += "<b>Breakdown by Task</b>`n"
    
    $EpicSummaryHtml  += "<br>"
    $EpicSummaryHtml  += "<b>Breakdown by Epics</b>`n"
    
    
    $data= $WorkInfo |Where-Object{$_.AssignedTo -eq $Assignee}
    $worklogTime = $data.Timespent
    if ($worklogTime -gt 0) {
        $time=$worklogTime.ToString("#.##")
    }
    else {
        $time = 0
    }
    $PTOList = $Availability |Where-Object{$_.Name -eq $Assignee -and [datetime]$_.Date -ge $latestSprint.startDate -and [datetime]$_.Date -le $latestSprint.endDate}
    $PTOTime =0
    foreach($PTO in $PTOList){
        $PTOTime += $PTO.Hours
    }
    $holidaystime = $holidayscount*8
    $totalhours = $PTOTime+$time+$holidaystime
    $lefthours = 80-$totalhours
    $WorkitemDetailsHtml += "<br>"
    $WorkitemDetailsHtml += "<b>Summary</b>`n"
    $WorkitemDetailsHtml +="<tbody><table border='1' cellspacing='1' style='font-size:10px; border:1px solid black; border-collapse: collapse' nowrap>
        <tr style='font-size:12px;'>
            <th valign='top' nowrap><b>Hours in Sprint</b></th>
            <th valign='top' nowrap><b>Hours Booked + PTO Holidays + PTO </b></th>
            <th valign='top' nowrap><b>Hours to be Booked </b></th>
        
        </tr>"
    $WorkitemDetailsHtml += "<tr style='font-size:12px;' nowrap>
    <td nowrap>80</td>
    <td align ='right' nowrap>$($totalhours)</td> 
    <td align ='right' nowrap>$($lefthours)</td> 

</tr>"
$PTOsHtml += "<br>"
$PTOsHtml += "<b>Breakdown: &nbsp;&nbsp;&nbsp; Hours Booked + PTO Holidays + PTO</b>`n"
$PTOsHtml += "<tbody><table border='1' cellspacing='1' style='font-size:10px; border:1px solid black; border-collapse: collapse' nowrap>
<tr style='font-size:12px;' nowrap>
    <td >Hours Booked</td>
    <td align ='center' >$($time)</td> 
</tr>"
$PTOsHtml += "<tr style='font-size:12px;' nowrap>
    <td >PTO - Holidays</td>
    <td align ='center' >$($holidaystime)</td> 
</tr>"
$PTOsHtml += "<tr style='font-size:12px;' nowrap>
    <td >PTO - Individual	&nbsp; </td>
    <td align ='center' >&nbsp; &nbsp; &nbsp;$($PTOTime) &nbsp; &nbsp; &nbsp; </td> 
</tr>"
$PTOsHtml += "<tr style='font-size:12px;' nowrap>
    <td ><b>Total</b></td>
    <td align ='center' ><b>$($totalhours)</b></td> 
</tr>"

    $issuesdetails = $WorkInfo |Where-Object{$_.AssignedTo -eq $Assignee -and $_.Timespent -gt 0}
    $issuesCount = $issuesdetails.Issueid
    
    if ($issuesdetails.Count -eq 0 ) 
    {
        $WorkitemSummaryHtml +="<p style='color:red;'>You have not yet booked time to any Tasks in this sprint </p>"
        $EpicSummaryHtml += "<p style='color:red;'>You have not yet booked time to any Epics in this sprint </p>"
    }
    else {
        $WorkitemSummaryHtml +="<tbody><table border='1' cellspacing='1' style='font-size:10px; border:1px solid black; border-collapse: collapse' nowrap>
        <tr style='font-size:12px;'>
            <th valign='top' nowrap><b>EPIC ID</b></th>
            <th valign='top' nowrap><b>Epic Name</b></th>
            <th valign='top' nowrap><b>Task ID</b></th>
            <th valign='top' nowrap><b>Task Name</b></th>
            <th valign='top' nowrap><b>Hours Booked</b></th>
        
        </tr>"
        $EpicSummaryHtml +="<tbody><table border='1' cellspacing='1' style='font-size:10px; border:1px solid black; border-collapse: collapse' nowrap>
        <tr style='font-size:12px;'>
            <th valign='top' nowrap><b>Epic ID</b></th>
            <th valign='top' nowrap><b>Epic Name</b></th>
            <th valign='top' nowrap><b>Hours Booked</b></th>
        
        </tr>"
        $totalTime = 0
        $epicdata = @()
        foreach($issue in $issuesCount){
            $data = $workupdates|Where-Object{$_.issueid -eq $issue -and $_.Timespent -gt 60}
            $summary = $data.IssueSummary
            $Epicinfos = Get-EpicInfo
            $Epicinfos.ParentID = $data.ParenID
            $Epicinfos.summary = $data.ParentSummary
            $Epicinfos.ChildTask = $issue
            
    
            $taskuri ="https://$($organization).atlassian.net/rest/api/2/issue/$($issue)/worklog"
            $taskResult = Invoke-RestMethod -Uri $taskuri -Method Get -Headers $headers
            $workinglogs = $taskResult.worklogs
            if ($null -ne $workinglogs) {    
                $timespentinseconds = 0
                $Totaltimespentinseconds = 0
                $latestworkinglogs=$workinglogs|Where-Object{[DateTime]$_.updated -gt $latestSprint.startDate}
                for ($n = 0; $n -lt $latestworkinglogs.Count; $n++) {
                    $Totaltimespentinseconds += [int]$latestworkinglogs[$n].timeSpentSeconds
                }
                if ($Totaltimespentinseconds -gt 60) {
                    $totalTime += $Totaltimespentinseconds
                    $Tasktimespentonwoklog = $Totaltimespentinseconds/3600
                    if ($Tasktimespentonwoklog -gt 0) {
                        $tasktime =$Tasktimespentonwoklog.ToString("#.##")
                    }
                    else {
                        $tasktime = 0
                    }
                    
                    $WorkitemSummaryHtml += "<tr style='font-size:12px;' nowrap>
                        <td nowrap> &nbsp;<a href='https://$($organization).atlassian.net/browse/$($data.ParenID)'> $($data.ParenID)</a>&nbsp;  </td>
                        <td nowrap>&nbsp;  $($data.ParentSummary)&nbsp;  </td>
                        <td nowrap>&nbsp; <a href='https://$($organization).atlassian.net/browse/$($data.Issueid)'> $($data.Issueid)</a> &nbsp; </td>
                        <td align ='left' nowrap>&nbsp;  $($summary) &nbsp; </td> 
                        <td align ='right' nowrap>&nbsp;  $($tasktime) &nbsp; </td> 
                    </tr>"
                    $Epicinfos.Timespent = $tasktime 
                    $epicdata += $Epicinfos 
                }
                
            } 
             
        }
        $parentinfo = $epicdata | Select-Object -Property ParentID -Unique
        $TotalEpicsTime = 0
        foreach($epicid in $Parentinfo.ParentID){
            $Totaltimespentinhours = 0
            $parentdata = $epicdata | Where-Object {$_.ParentID -eq $epicid}
            $parentid = $parentdata.ParentID |Select-Object -Unique
            $parentsummary=$parentdata.summary |Select-Object -Unique
            for ($l = 0; $l -lt $parentdata.Count; $l++) {
                $Totaltimespentinhours += $parentdata[$l].Timespent
            }
            $TotalEpicsTime += $Totaltimespentinhours
            $EpicSummaryHtml += "<tr style='font-size:12px;' nowrap>
                        <td nowrap> &nbsp; <a href='https://$($organization).atlassian.net/browse/$($parentid)'>$($parentid)</a>&nbsp;  </td>
                        <td align ='left' nowrap>&nbsp;  $($parentsummary)&nbsp;  </td> 
                        <td align ='right' nowrap> &nbsp; $($Totaltimespentinhours) &nbsp; </td>
                    </tr>"
        }
        if ($TotalEpicsTime -gt 0) {
            $EpicSummaryHtml += "<tr style='font-size:12px;' nowrap>
                        <td nowrap><b>&nbsp;  Total &nbsp;  </b></td>
                        <td align ='left' nowrap></td> 
                        <td align ='right' nowrap><b>&nbsp;  $($TotalEpicsTime) &nbsp; </b></td>
                    </tr>"
        }
        if ($totalTime -gt 60) {
            # $totatimeinSeconds = [timespan]::fromseconds($totalTime)
            # $Total="$($totatimeinSeconds.hours):$($totatimeinSeconds.minutes)"
            $Total = $totalTime/3600
            if ($Total -gt 0) {
                $AllTasksTime = $Total.ToString("#.##")
            }
            else {
                $AllTasksTime = 0
            }
            $WorkitemSummaryHtml += "<tr style='font-size:12px;' nowrap>
                <td nowrap><b>&nbsp;  Total &nbsp; </b></td>
                <td align ='right' nowrap></td> 
                <td align ='right' nowrap></td>
                <td align ='right' nowrap></td>
                <td align ='right' nowrap><b>&nbsp;  $($AllTasksTime) &nbsp; </b></td> 
        </tr>" 
        
        }
    }
    
    
    $WorkitemSummaryHtml += "</table></tbody>"
    $WorkitemDetailsHtml += "</table></tbody>"
    $EpicSummaryHtml += "</table></tbody>"
    $PTOsHtml += "</table></tbody>"
    $body += $WorkitemSummaryHtml1
    $body += $WorkitemDetailsHtml
    $body += $PTOsHtml
    $body += $EpicSummaryHtml
    $body += $WorkitemSummaryHtml

    # # compose and send out email messages to individuals who have booked to at least one task in the sprint
    Write-Host "-------------------------------------------------------------------------------------------"
    Write-Host "Sending out email for " $Assignee
    # $email=("")
    $emailSubject = "Time spent Summary report for " +$Assignee 
    Start-Sleep -s 15

    $SMTP_SERVER = "smtp.socketlabs.com"
    $SMTP_PORT = 587
    $SMTP_USERNAME = "server4507"
    $PWORD = ConvertTo-SecureString -String $info -AsPlainText -Force
    # $bccEmailList = ("")
    $email = ("")
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
    $message.To.Add($email)
    $message.Subject = $emailSubject
    $message.Body = $body
    # $message.cc.Add($bccEmailList)
    $message.IsBodyHtml = $true
    # Send-MailMessage -SmtpServer $SMTP_SERVER -Port $SMTP_PORT -Credential $cred -To $email -Bcc $bccEmailList -From "" -Subject $emailSubject -Body $body -BodyAsHtml
    try {
        # Send-MailMessage -To $email -SmtpServer $SMTP_SERVER -Port $SMTP_PORT -Credential $cred -Bcc $bccEmailList -From "" -Subject $emailSubject -Body $body -BodyAsHtml
        $SMTPClient.Send($message)
        $SMTPClient.Dispose()
        $message.Dispose()
        $body = $null
        $WorkitemSummaryHtml1 = $null
        $WorkitemDetailsHtml = $null
        $WorkitemSummaryHtml = $null
        $EpicSummaryHtml = $null
        $epicdata = $null
        $PTOsHtml = $null
    }
    catch {
        $_.Exception.message
    }

    # Stop-Transcript

}