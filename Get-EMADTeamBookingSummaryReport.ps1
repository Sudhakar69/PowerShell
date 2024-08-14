ssparam ([string]$info="Please provide password")

$jiraPATId="Personal Access Token"
$path1 = "vsts-scripts\projectreporting\project_reporting_common.ps1"
# Load and execute the common library script
$commonLibContent1 = Get-Content -Path $Path1 -Raw
Invoke-Expression -Command $commonLibContent1

# $project= "Task"

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
    $WorkItemUpdate| Add-Member -MemberType NoteProperty -Name updatedtime -Value $null
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
    $EmployeeWorkAvailability| Add-Member -MemberType NoteProperty -Name updatedtime -Value $null
    return $EmployeeWorkAvailability
    
}

# $name = "Dariusz Orlowski"
$RP_FILEPATH = "Sharepoint path"
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
$polandHolidays = $holidayInfosPol |Where-Object{[datetime]$_.date -ge $sprintstartdate -and [datetime]$_.date -le $sprintenddate -and $_.isweekend -ne $true}
$holidaysnames = $polandHolidays.name
$holidayscount = $holidaysnames.Count
Write-Host "Holidays Data has been collected"
# # $WorkitemsQueryUri = "https://$($organization).atlassian.net/rest/api/2/project/Task"
# $WorkitemsQueryUri = "https://$($organization).atlassian.net/rest/api/2/dashboard"
# $WorkitemsQueryUri = "https://$($organization).atlassian.net/rest/api/2/issue/Task-2699"
# https://$($organization).atlassian.net/rest/agile/1.0/sprint/520
$workflowuri= "https://$($organization).atlassian.net/rest/api/2/search?jql=ORDER%20BY%20updated&maxResults=10000"
$workflowResult = Invoke-RestMethod -Uri $workflowuri -Method Get -Headers $headers
$Total = $workflowResult.total
$workupdates = @()
[Int64]$Totalcount = $Total/100
#  $Taskkeys = @()
for ($j = 0; $j -le $Totalcount ; $j++) {
    $startAt = $j*100
    $Taskworkflowuri= "https://$($organization).atlassian.net/rest/api/2/search?jql=ORDER%20BY%20updated&maxResults=100&startAt=$($startAt)"
    $TaskworkflowResult = Invoke-RestMethod -Uri $Taskworkflowuri -Method Get -Headers $headers
    $Task = $TaskworkflowResult.issues.key |Where-Object{$_ -match "Task"}
    $keys = $Task|Select-Object -Unique |Sort-Object
    # $Taskkeys += $Task
# }
    # $keys = $Taskkeys|Select-Object -Unique |Sort-Object
    foreach($issueid in $keys) {
        # $issueid = $Task[$i].key
        $issueuri ="https://$($organization).atlassian.net/rest/api/2/issue/$($issueid)"
        $issueResult = Invoke-RestMethod -Uri $issueuri -Method Get -Headers $headers
        $issueworkloguri ="https://$($organization).atlassian.net/rest/api/2/issue/$($issueid)/worklog"
        $issueworklogResult = Invoke-RestMethod -Uri $issueworkloguri -Method Get -Headers $headers
        $issues = $issueResult.fields
        $worklogs = $issueworklogResult.worklogs
        if ($null -ne $worklogs) {    
            $timespentinseconds = 0
            $latestworklogs=$worklogs|Where-Object{[DateTime]$_.started -gt $latestSprint.startDate}
            for ($i = 0; $i -lt $latestworklogs.Count; $i++) {
                $timespentinseconds += [int]$latestworklogs[$i].timeSpentSeconds
                $WorkItemUpdate = Get-WorkItemUpdate
                $Parentinfo=Get-parentInfo -issueid $issueid
                if ($null -ne $latestworklogs[$i].updateAuthor.displayName) {
                    $WorkItemUpdate.EMail = $latestworklogs[$i].updateAuthor.emailAddress
                    $WorkItemUpdate.AssignedTo = $latestworklogs[$i].updateAuthor.displayName
                }
                else {
                    $WorkItemUpdate.EMail = $null
                    $WorkItemUpdate.AssignedTo = $null
                }
                $WorkItemUpdate.Issueid = $issueid
                $WorkItemUpdate.Status = $issues.status.name
                $WorkItemUpdate.IssueSummary = $issues.summary
                $WorkItemUpdate.ParenID = $Parentinfo.ID
                $WorkItemUpdate.ParentSummary = $Parentinfo.summary
                $WorkItemUpdate.Timespent =$latestworklogs[$i].timeSpentSeconds
                [datetime]$WorkItemUpdate.updatedtime =$latestworklogs[$i].started
                $workupdates += $WorkItemUpdate
            }
        }        
    }
}
# $Assignees = ("")
$Assignees = ("")
[DateTime]$currentDateTime = Get-Date
# $WorkInfo = @()
foreach($Assignee in $Assignees){
    [DateTime]$startDate=([DateTime]$latestSprint.startDate).ToUniversalTime()
    [DateTime]$endDate=([DateTime]$latestSprint.endDate).ToUniversalTime()
    $today =$currentDateTime.ToUniversalTime()
    $WorkitemSummaryHtml1 = "<h1>$($Assignee)</h1>"
    $WorkitemSummaryHtml1 += "<b>Your Time Booking Since Start of Sprint</b>"
    $WorkitemSummaryHtml1 += "<br><b>As Of Date/Time - UTC: </b>"+ $today.ToString("MM/dd/yyyy hh:mm:ss tt")
    $WorkitemSummaryHtml1 += "<br><b>Poland: </b> " + [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId($today, 'Central Europe Standard Time').toString()
    $WorkitemSummaryHtml1 += "<br><b>Spain/France: </b> " + [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId($today, 'Central Europe Standard Time').toString()
    $WorkitemSummaryHtml1 += "<br><b>UTC: </b>" + $today.ToString("MM/dd/yyyy hh:mm:ss tt")
    $WorkitemSummaryHtml1 += "<br>"
    $WorkitemSummaryHtml1 += "<br><b>Current Sprint: </b>" + $latestSprint.Name
    $WorkitemSummaryHtml1 += "<br><b>Start Date: </b>" + $startDate.ToString("MM/dd/yyyy hh:mm:ss tt")
    $WorkitemSummaryHtml1 += "<br><b>End Date: </b>" + $endDate.ToString("MM/dd/yyyy hh:mm:ss tt")
    $WorkitemSummaryHtml1 += "<br><b>Version: &nbsp; </b>V2.21" 
    $WorkitemSummaryHtml1 += "<br>"
    $WorkitemSummaryHtml += "<br>"
    $WorkitemSummaryHtml += "<b>Breakdown by Task</b>`n"
    
    $EpicSummaryHtml  += "<br>"
    $EpicSummaryHtml  += "<b>Breakdown by Epics</b>`n"
    $WorkitemDetailsHtml += "<br>"
    $WorkitemDetailsHtml += "<b>Summary</b>`n"
    $WorkitemDetailsHtml +="<tbody><table border='1' cellspacing='1' style='font-size:10px; border:1px solid black; border-collapse: collapse' nowrap>
        <tr style='font-size:12px;'>
            <th valign='top' nowrap><b>Hours in Sprint</b></th>
            <th valign='top' nowrap><b>Hours Booked + PTO Holidays + PTO </b></th>
            <th valign='top' nowrap><b>Hours to be Booked </b></th>
        
        </tr>"
    $workInfo = $workupdates |Where-Object{$_.AssignedTo -eq $Assignee}
    $email = $workInfo.email | Select-Object -Unique
    $Totaltimespentinseconds = 0
    for ($i = 0; $i -lt $workInfo.Count; $i++) {
        $Totaltimespentinseconds += $workInfo[$i].Timespent
    }
    $Totaltimespentinhours = $Totaltimespentinseconds/3600
    $totaltimeinhours = $Totaltimespentinhours.ToString("#.##")
    if ($totaltimeinhours -gt 0) {
        $totalinhours = $totaltimeinhours
    }
    else {
        $totalinhours = 0
    }
    $PTOList = $Availability |Where-Object{$_.Name -eq $Assignee -and [datetime]$_.Date -ge $latestSprint.startDate -and [datetime]$_.Date -le $latestSprint.endDate -and $_.Type -ne "Public Holiday" }
    $holidayscount=$Availability |Where-Object{$_.Name -eq $Assignee -and [datetime]$_.Date -ge $latestSprint.startDate -and [datetime]$_.Date -le $latestSprint.endDate -and $_.Type -eq "Public Holiday" }
    $PTOTime =0
    $holidaystime = 0
    foreach($PTO in $PTOList){
        $PTOTime += $PTO.Hours
    }
    # if ($holidayscount -gt 0) {
    #     $holidaystime = $holidayscount.Hours
    # }
    # else {
    #     $holidaystime = 0
    # }
    foreach($holiday in $holidayscount){
        $holidaystime += $holiday.Hours
    }
    # $holidaystime = $holidayscount*8
    $totalhours = $PTOTime+$totalinhours+$holidaystime
    $Totallefthours = 80-$totalhours
    $lefthours = $Totallefthours.ToString("F2")
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
            <td align ='center' >$($totalinhours)</td> 
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
    $issuesdetails = $workupdates |Where-Object{$_.AssignedTo -eq $Assignee -and $_.Timespent -gt 0 -and ([datetime]$_.updatedtime).ToUniversalTime() -ge [datetime]$latestSprint.startDate}
    $issuesCount = $issuesdetails.ParenID |Select-Object -Unique
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
        $TotalEpicsTime = 0
        
        foreach($issue in $issuesCount){
            $data = $issuesdetails|Where-Object{$_.ParenID -eq $issue }
            # $summary = $data.ParentSummary |Select-Object -Property ParentSummary -Unique
            $parentid = $issue
            $parentsummary=$data.ParentSummary|Select-Object -Unique
            $Totaltimespentinseconds = 0
            for ($l = 0; $l -lt $data.Count; $l++) {
                $Totaltimespentinseconds += [Int64]$data[$l].Timespent
            }
            $Totaltimespent = $Totaltimespentinseconds/3600
            $Totaltimespentinhours = $Totaltimespent.ToString("#.##")
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
        }
        
        $issueids = $issuesdetails.Issueid |Select-Object -Unique
        $TotalTimeonissuesinseconds =0
        foreach($issuesid in $issueids){
            $issueinfo = $issuesdetails |Where-Object{$_.Issueid -eq $issuesid}
            $issueparentid = $issueinfo.ParenID |Select-Object -Unique
            $parentidsummary = $issueinfo.ParentSummary|Select-Object -Unique
            $issuesummary = $issueinfo.IssueSummary|Select-Object -Unique
            $timespentonissueinseconds =0
            for ($k = 0; $k -lt $issueinfo.Count; $k++) {
                $timespentonissueinseconds += $issueinfo[$k].Timespent
            }
            
            $timespentonissue = $timespentonissueinseconds/3600
            $timespentonissueinhours = $timespentonissue.toString("#.##")
            $TotalTimeonissuesinseconds += $timespentonissueinseconds
            $WorkitemSummaryHtml += "<tr style='font-size:12px;' nowrap>
                        <td nowrap> &nbsp;<a href='https://$($organization).atlassian.net/browse/$($issueparentid)'> $($issueparentid)</a>&nbsp;  </td>
                        <td nowrap>&nbsp;  $($parentidsummary)&nbsp;  </td>
                        <td nowrap>&nbsp; <a href='https://$($organization).atlassian.net/browse/$($issuesid)'> $($issuesid)</a> &nbsp; </td>
                        <td align ='left' nowrap>&nbsp;  $($issuesummary) &nbsp; </td> 
                        <td align ='right' nowrap>&nbsp;  $($timespentonissueinhours) &nbsp; </td> 
                    </tr>"
        }
        if ($TotalTimeonissuesinseconds -gt 0) {
            $TotalTimeonissuesinhours = $TotalTimeonissuesinseconds/3600
            $TotalTimeonissues = $TotalTimeonissuesinhours.ToString("#.##")
            $WorkitemSummaryHtml += "<tr style='font-size:12px;' nowrap>
                <td nowrap><b>&nbsp;  Total &nbsp; </b></td>
                <td align ='right' nowrap></td> 
                <td align ='right' nowrap></td>
                <td align ='right' nowrap></td>
                <td align ='right' nowrap><b>&nbsp;  $($TotalTimeonissues) &nbsp; </b></td> 
            </tr>"
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
    $emailSubject = "Time Booking Summary for " +$latestSprint.Name+ "; " +$lefthours+" hour(s) to be booked in this sprint ( " +$Assignee+ "< " +$email+ " > )"
    Start-Sleep -s 15

    $SMTP_SERVER = "smtp.socketlabs.com"
    $SMTP_PORT = 587
    $SMTP_USERNAME = "server4507"
    $PWORD = ConvertTo-SecureString -String $info -AsPlainText -Force
    $bccEmailList = ("")
    $email = ("")
    # $email = ("")
    # $bccEmailList = ("")
    # $cred = New-Object -TypeName System.Management.Automation.PSCredential($SMTP_USERNAME, $PWORD)
    $SMTPClient = New-Object System.Net.Mail.SmtpClient
    $SMTPClient.Host = $SMTP_SERVER
    $SMTPClient.Port = $SMTP_PORT
    $SMTPClient.EnableSsl = $true
    $SMTPClient.Credentials = New-Object System.Net.NetworkCredential($SMTP_USERNAME, $PWORD)
    $message= New-Object System.Net.Mail.MailMessage
    $message.From = """"
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
}

