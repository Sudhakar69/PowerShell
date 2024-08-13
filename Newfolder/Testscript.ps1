param ([string]$info="PLEASE PROVIDE PASSWORD AS AN ARGUMENT TO THIS SCRIPT")


Add-Type -AssemblyName System.Web

$scriptDirectory = ($PSScriptRoot | Split-Path -Parent)
$commonLibPath = (Join-Path (Join-Path $scriptDirectory -ChildPath "common") -ChildPath "common.ps1")
. $commonLibPath

$infoLibPath = (Join-Path (Join-Path $scriptDirectory -ChildPath "common") -ChildPath "info.ps1")
. $infoLibPath

$infoLibPath = (Join-Path $PSScriptRoot -ChildPath "project_reporting_common.ps1")
. $infoLibPath
$queryId = "343d58e6-67ad-48a1-9317-1403d32e1dab"

# execute the query
Write-Host "Executing 'Project Reporting Snapshot' query (query id: $queryId)"
$projectReportingQueryUri = "https://apxinc.visualstudio.com/DefaultCollection/_apis/wit/wiql/$($queryId)?api-version=4.1"
$projectReportingQueryResult = Invoke-RestMethod -Uri $projectReportingQueryUri -Method Get -ContentType "application/json" -Headers @{Authorization=("Basic {0}" -f $base64AuthInfo)}
Write-Host "Project Reporting query result: $projectReportingQueryResult"

function Create-WorkItemInfo
{
    $names = @("workItemType","workItemId","workItemName","epicId","epicName","assignedTo","initialActualHours","subsequentActualHours")
    $workItemInfo = New-Object –TypeName PSObject
    foreach($name in $names){
        $workItemInfo | Add-Member –MemberType NoteProperty –Name $name -Value $null
    }
}

foreach ($email in ($hoursBookedByResource.Keys | Sort-Object))
{   
    # only send emails to those individuals present in WORK_RESOURCES hashtable and which are in the 
    # appropriate location for current time of day
    if ($WORK_RESOURCES.ContainsKey($email))  
    { 
        $resourceInfo = $WORK_RESOURCES.Item($email)
        $gmtHour = $currentDateTime.ToUniversalTime().Hour
        $locationsForHour = Get-LocationForHour -hourOfDay $gmtHour
        
        Write-Host "GMT hour: $gmtHour; locations for hour: $locationsForHour"
        if ($locationsForHour -contains $resourceInfo.location)
        {
            $timeBooking = $hoursBookedByResource.Item($email)
            $taskTimeBookings = $timeBooking.taskTimeBookings
            $timeBookingAggregates = $timeBooking.timeBookingAggregates

            # calculate holiday hours
            $country = $resourceInfo.country
            $holidays = $holidaysByCountry.Item($country)
            $holidayHours = $holidays.Count * 8

            # calculate total time booked to projects
            [decimal]$totalAcrossProjects = 0
            foreach ($agg in $timeBookingAggregates.Values)
            {
                $hoursBooked = $agg.hoursBooked
                $totalAcrossProjects += $hoursBooked
            }

            # calculate individual PTO
            $ptoInfos = $ptoInfosByResource.Item($email)
            [decimal]$ptoTotalHours = 0
            if ($ptoInfos -ne $null)
            {
                foreach ($ptoInfo in $ptoInfos)
                {
                    $ptoTotalHours += $ptoInfo.hours
                }
            }

            # Calculate estimated OH
            [decimal]$resourceOHRate = $resourceInfo.resourceOHRate
            [decimal]$estimatedOH = ([decimal]$resourceInfo.capacityPerSprint - $holidayHours - $ptoTotalHours) * $resourceOHRate
        
            # Calculate remaining hours to be booked
            [decimal]$hoursToBeBooked = [decimal]$resourceInfo.capacityPerSprint - $totalAcrossProjects - $estimatedOH - $holidayHours - $ptoTotalHours
                
            Write-Host ("Sending to " + $email)
        
            # Display header
            $body = "<h2>$email</h2>"
            $body += "<h3>Your Time Booking Since Start of Sprint</h3>"

            # Display sprint information. Note, per Microsoft documentaiton, time zones utilized below are daylight-savings-time-aware
            $body += ("<b>As Of Date - UTC:</b> " + $currentDateTime.ToUniversalTime() + "<br>")
            $body += ("<b>As Of Date - Romania:</b> " + ([System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId($currentDateTime, 'E. Europe Standard Time').toString()) + "<br>")
            $body += ("<b>As Of Date - Poland:</b> " + ([System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId($currentDateTime, 'E. Europe Standard Time').toString()) + "<br>")
            $body += ("<b>As Of Date - ET:</b> " + ([System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId($currentDateTime, 'Eastern Standard Time').toString()) + "<br>")
            $body += ("<b>As Of Date - PT:</b> " + ([System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId($currentDateTime, 'Pacific Standard Time').toString()) + "<br>")
            $body += ("<br><b>Sprint:</b> " + $sprintInfo.sprintNumber + "<br><b>Start Date:</b> " + $sprintInfo.startDate + "<br><b>End Date:</b> " + $sprintInfo.endDate)


            # Display summary table
            $isNegative = ($hoursToBeBooked -lt 0)        
            $body += "<h3>Summary</h3>"
            $body += "<table border='1' cellspacing='0' cellpadding='3'><tr>
                <td valign='top'><b>Hours in Sprint<b></td>
                <td valign='top'><b>Hours Booked + Estimated OH + PTO</b></td>
                <td valign='top'><b>Hours to Be Booked<br>
                    <font size='2'>Negative numbers indicate overbooking</font>
                </b></font></td></tr><tbody>"
            $body += "<tr><td align='right'>" + [decimal]$resourceInfo.capacityPerSprint.ToString("0.00") + "</td>
                            <td align='right'>" + ($totalAcrossProjects + $estimatedOH + $holidayHours + $ptoTotalHours).ToString("0.00") + "</td>
                            <td align='right'><b>" + ($hoursToBeBooked).ToString("0.00") + "</b></td></tr>"
            $body += "</tbody></table>"

            $body += "<br><font size='2'><b>Breakdown: Hours Booked + Estimated OH + PTO</b></font><br>"
            $body += "<table border='1' cellspacing='0' cellpadding='3'>
            <tr><td valign='top'><font size='2'>Hours Booked</font></td><td valign='top'><font size='2'>" + $totalAcrossProjects.ToString("0.00") + "</font></td></tr>
            <tr><td valign='top'><font size='2'>Estimated OH</font></td><td valign='top'><font size='2'>" + $estimatedOH.ToString("0.00") + "</font></td></tr>
            <tr><td valign='top'><font size='2'>PTO - Holidays</font></td><td valign='top'><font size='2'>" + $holidayHours.toString("0.00") + "</font></td></tr>
            <tr><td valign='top'><font size='2'>PTO - Individual</font></td><td valign='top'><font size='2'>" + $ptoTotalHours.toString("0.00") + "</font></td></tr>
            <tr><td valign='top'><font size='2'><b>Total<b></font></td><td valign='top'><font size='2'><b>" + ($totalAcrossProjects + $estimatedOH + $holidayHours + $ptoTotalHours).ToString("0.00") + "</b></font></td></tr>
            </table>"


            # Display table of projects
            $body += "<h3>Breakdown by Project</h3>"

            if ($timeBookingAggregates.Count -eq 0)
            {            
                $body += "<font size='3' color='red'><b>You have not yet booked time to any tasks/projects in this sprint</b></font>"
            }
            else
            {
                $body += "<table border='1' cellspacing='0' cellpadding='3'><thead><th>Project Name</th><th>Epic Id</th><th>Hours Booked</th></thead><tbody>"
                foreach ($agg in $timeBookingAggregates.Values)
                {
                    $epicId = $agg.epicId
                    $epicName = $agg.epicName
                    $hoursBooked = $agg.hoursBooked
                    $body += "<tr><td>$epicName</td>
                                    <td><a href='https://apxinc.visualstudio.com/Apx/_workitems/edit/$epicId'>$epicId</a></td>
                                    <td align='right'>" + $hoursBooked.ToString("0.00") + "</td></tr></td>"
                }
                $body += "<tr><td><b>Total</b></td><td></td><td align='right'><b>" + $totalAcrossProjects.ToString("0.00") + "</b></td></tr>"
                $body += "</tbody></table>"

                # Display table of tasks
                $body += "<h3>Breakdown by Task</h3>"
                $body += "<table border='1' cellspacing='0' cellpadding='3'><thead><th>Project Name</th><th>Epic Id</th><th>Task Name</th><th>Task Id</th><th>Hours Booked</th></thead><tbody>"
                [decimal]$totalAcrossTasks = 0
                foreach ($booking in $taskTimeBookings.Values)
                {
                    $taskId = $booking.taskId
                    $taskName = $booking.taskName
                    $epicId = $booking.epicId
                    $epicName = $booking.epicName
                    $hoursBooked = $booking.hoursBooked
                    $body += "<tr><td>$epicName</td><td><a href='https://apxinc.visualstudio.com/Apx/_workitems/edit/$epicId'>$epicId</a></td>
                            <td>$taskName</td>
                            <td><a href='https://apxinc.visualstudio.com/Apx/_workitems/edit/$taskId'>$taskId</a></td>
                            <td align='right'>" + $hoursBooked.ToString("0.00") + "</td></tr></td>"
    
                    $totalAcrossTasks += $hoursBooked
                }
                $body += "<tr><td><b>Total</b></td><td></td><td></td><td></td><td align='right'><b>" + $totalAcrossTasks.ToString("0.00") + "</b></td></tr>"
                $body += "</tbody></table>"
            }

            # Display holidays
            $body += "<h3>Holidays in Current Sprint</h3>"
            if ($holidays.Count -eq 0)
            {
                $body += "There are no holidays in the current sprint"
            }
            else
            {
                $body += "<table border='1' cellspacing='0' cellpadding='3'><thead><th>Holiday</th><th>Date</th><th>Hours</th></thead><tbody>"
                foreach($holidayInfo in $holidays)
                {
                    $holidayName = $holidayInfo.name
                    $holidayDateStr = $holidayInfo.date.toString("dd-MMM-yyyy")
                    $body += "<tr><td>$holidayName</td><td>$holidayDateStr</td><td align='right'>8.00</td></tr>"
                }

                $body += "<tr><td><b>Total</b></td><td></td><td align='right'><b>" + $holidayHours.ToString("0.00") + "</b></td></tr>"
                $body += "</tbody></table>"
            }

            # Display holidays
            $body += "<h3>Individual PTO in Current Sprint</h3>"
            $ptoInfos = $ptoInfosByResource.Item($email)
            if ($ptoInfos -eq $null -or $ptoInfos.Count -eq 0)
            {
                $body += "You have no individual PTO in the current sprint"        
            }
            else
            {
                $body += "<table border='1' cellspacing='0' cellpadding='3'><thead><th>Date</th><th>Hours</th></thead><tbody>"
                foreach($ptoInfo in $ptoInfos)
                {
                    $ptoDateStr = $ptoInfo.date.toString("dd-MMM-yyyy")
                    [decimal]$ptoHours = $ptoInfo.hours
                    $body += "<tr><td>$ptoDateStr</td><td align='right'>" + ($ptoHours).ToString("0.00") + "</td></tr>"
                }

                $body += "<tr><td><b>Total</b></td><td align='right'><b>" + ($ptoTotalHours).ToString("0.00") + "</b></td></tr>"
                $body += "</tbody></table>"        
            }

            if ($resourceInfo.roleCode -in $ROLE_CODE_PM, $ROLE_CODE_CIO, $ROLE_CODE_DEVOPS_QA)
            {
                # For select roles only, display matrix of time booking, by epic, by resource
                $body += "<br><br><br><h3>Team Time Booking</h3>"
                $body += $matrixHTML
                $body += "<br><br>"
            }

            $emailSubject = "Time Booking Summary for Sprint " + $sprintInfo.sprintNumber
            if (!$isNegative)
            {
                $emailSubject += "; " + ($hoursToBeBooked).ToString("0.00") + " hour(s) to be booked in this sprint"
            }
            $emailSubject += " (" + $email + ")"
        
            #if ($email.IndexOf("gurvits") -ge 0)
            #{
                Start-Sleep -s 15
		
                $SMTP_SERVER = "smtp.socketlabs.com"
                $SMTP_PORT = 25
                $SMTP_USERNAME = "server4507"
                $PWORD = ConvertTo-SecureString -String $info -AsPlainText -Force

                #$body | Out-file -Filepath "R:\vsts-scripts\projectreporting\temp\$($email).html"
                #$body | Out-file -Filepath "C:\_tmp\sprint-time-booking-emails.html" -Append
                #Write-Host ">>>> sending email to: " $email
                #$email = "agurvits@apx.com"
                
                $bccEmailList = @("agurvits@apx.com","kliang@apx.com")
                if ($resourceInfo.roleCode -eq $ROLE_CODE_DB_DEV)
                {
                    $bccEmailList += ,"rvenkatachalam@apx.com"
                }

                $cred = New-Object -TypeName System.Management.Automation.PSCredential($SMTP_USERNAME, $PWORD)
        	    Send-MailMessage -SmtpServer $SMTP_SERVER -Port $SMTP_PORT -Credential $cred -To $email -Bcc $bccEmailList -From tfsbuild@apx.com -Subject $emailSubject -Body $body -BodyAsHtml

                #$cred = New-Object -TypeName System.Management.Automation.PSCredential($SMTP_USERNAME, $PWORD)
        	    #Send-MailMessage -SmtpServer $SMTP_SERVER -Port 2525 -Credential $cred -To "agurvits@apx.com" -From tfsbuild@apx.com -Subject $emailSubject -Body $body -BodyAsHtml

            #}
        }
    }
}
Write-Host "Finished sending emails"