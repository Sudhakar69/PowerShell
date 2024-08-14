param ([string]$info="Password")
$path= "common.ps1"
# # Load and execute the common library script
$commonLibContent = Get-Content -Path $Path -Raw
Invoke-Expression -Command $commonLibContent
# Define your Snyk API token

$token = "Snyk token"
# Define the headers for the API request
$headers = @{
    "Content-Type" = "application/vnd.api+json"
    "Authorization" = "token $token"  
}
function Get-DevInfo
{
    $DevInfo = New-Object -TypeName PSObject
    $DevInfo | Add-Member -MemberType NoteProperty -Name ID -Value $null
    $DevInfo | Add-Member -MemberType NoteProperty -Name Title -Value $null
    $DevInfo | Add-Member -MemberType NoteProperty -Name vulnerabilityID -Value $null
    $DevInfo | Add-Member -MemberType NoteProperty -Name Ignored -Value $null
    $DevInfo | Add-Member -MemberType NoteProperty -Name Fixable -Value $null
    $DevInfo | Add-Member -MemberType NoteProperty -Name Upgradable -Value $null
    $DevInfo | Add-Member -MemberType NoteProperty -Name Severity -Value $null
    $DevInfo | Add-Member -MemberType NoteProperty -Name CreatedTime -Value $null
    $DevInfo | Add-Member -MemberType NoteProperty -Name ProjectId -Value $null
    $DevInfo | Add-Member -MemberType NoteProperty -Name ProjectName -Value $null
    $DevInfo | Add-Member -MemberType NoteProperty -Name vulnerabilityAge -Value $null

    return $DevInfo
}
function Get-ProdInfo
{
    $ProdInfo = New-Object -TypeName PSObject
    $ProdInfo | Add-Member -MemberType NoteProperty -Name ID -Value $null
    $ProdInfo | Add-Member -MemberType NoteProperty -Name Title -Value $null
    $ProdInfo | Add-Member -MemberType NoteProperty -Name vulnerabilityID -Value $null
    $ProdInfo | Add-Member -MemberType NoteProperty -Name Ignored -Value $null
    $ProdInfo | Add-Member -MemberType NoteProperty -Name Severity -Value $null
    $ProdInfo | Add-Member -MemberType NoteProperty -Name CreatedTime -Value $null
    $ProdInfo | Add-Member -MemberType NoteProperty -Name ProjectId -Value $null
    $ProdInfo | Add-Member -MemberType NoteProperty -Name vulnerabilityAge -Value $null
    return $ProdInfo
}
$currentdate = (Get-Date).ToUniversalTime()
# Define the org Ids 
$DevorgId ="Dev org ID"
$ProdorgId ="Prod Org ID"
# Define the URL for the API request
# example of getting orgs
# $Devurl ="https://api.snyk.io/rest/orgs/$($DevorgId)/issues?version=2024-05-23&limit=100&type=package_vulnerability&created_before=$($Createdbefore)&created_after=$($createdafter)&effective_severity_level=$($severity)&status=open"
# $url = "https://api.snyk.io/v1/org/$($orgId)"
# $url = "https://api.snyk.io/rest/orgs/$($orgId)"
$summaryreport +="<tbody><table border='1' cellspacing='1' style='font-size:10px; border:1px solid black; border-collapse: collapse' nowrap>
    <tr style='font-size:12px;'>
        <th valign='top' nowrap><b>&nbsp;</b></th>
        <th valign='top' nowrap><b>Ignored</b></th>
        <th valign='top' nowrap><b>Not Fixed</b></th>
        <th valign='top' nowrap><b>Fixed but Not <br>released to Production</b></th>
        <th valign='top' nowrap><b>Total</b></th>
    </tr>"

$severities = @("critical","high")
$Devresult = @()
$Prodresult = @()
foreach($severity in $severities){
    $i=0
    while ($i -le 1250) {
        # $j = $i*5
        $j = $i + 1
        $startdays = 0 - $i
        $tilldays = 0 - $j
        $daysbefore = $currentdate.AddDays($startdays)
        $afterdays = $currentdate.AddDays($tilldays)
        $Createdbefore = $daysbefore.ToString("yyyy-MM-ddTHH:mm:ss.ffZ")
        $createdafter = $afterdays.ToString("yyyy-MM-ddTHH:mm:ss.ffZ")
        $Devurl =  "https://api.snyk.io/rest/orgs/$($DevorgId)/issues?version=2024-05-23&limit=100&created_before=$($Createdbefore)&created_after=$($createdafter)&effective_severity_level=$($severity)&status=open"
        $produrl =  "https://api.snyk.io/rest/orgs/$($ProdorgId)/issues?version=2024-05-23&limit=100&created_before=$($Createdbefore)&created_after=$($createdafter)&effective_severity_level=$($severity)&status=open"
        try {
            $Devresponse = Invoke-RestMethod -Uri $Devurl -Headers $headers -Method Get
            $Prodresponse = Invoke-RestMethod -Uri $produrl -Headers $headers -Method Get
        }
        catch {
            $_.Exception.Message
        }
        
        $Devdata = $Devresponse.data
        $Proddata = $Prodresponse.data
        foreach($devitem in $Devdata){
            $attributes = $devitem.attributes
            $problems=$attributes.problems
            $coordinates = $attributes.coordinates
            for ($Devcount = 0; $Devcount -lt $problems.Count; $Devcount++) {
                $projectId = $devitem.relationships.scan_item.data.ID
                $projecturl = "https://api.snyk.io/rest/orgs/$($DevorgId)/projects/$($projectId)?version=2024-05-23"
                $DevProjectresponse = Invoke-RestMethod -Uri $projecturl -Headers $headers -Method Get
                if ($coordinates.PSObject.properties['Is_Fixable_manually'] -and $coordinates.Is_Fixable_manually -eq $true -or $coordinates.Is_Fixable_snyk -eq $true -or $coordinates.is_fixable_upstream -eq $true) {
                    $fixable = $true
                }
                else {
                    $fixable = $false
                }
                if ($coordinates.PSObject.properties['is_upgradeable'] -and $coordinates.is_upgradeable -eq $true) {
                    $upgradable = $true
                }
                else {
                    $upgradable = $false
                }
                $createdate = $attributes.created_at
                $issueage = New-TimeSpan -Start $createdate -End $currentdate
                $DevInfo = Get-DevInfo
                $DevInfo.ID = $devitem.ID
                $DevInfo.Title = $attributes.title
                $DevInfo.Ignored = $attributes.Ignored
                $DevInfo.Severity = $attributes.effective_severity_level
                $DevInfo.CreatedTime = $attributes.created_at
                $DevInfo.Fixable = $fixable
                $DevInfo.Upgradable = $upgradable
                $DevInfo.vulnerabilityID = $problems[$Devcount].ID
                $DevInfo.ProjectId = $devitem.relationships.scan_item.data.ID
                $DevInfo.ProjectName = $DevProjectresponse.data.attributes.name
                $DevInfo.vulnerabilityAge = $issueage.Days
                $Devresult += $DevInfo
            }            
        }
        foreach($proditem in $Proddata){
            $attributes = $proditem.attributes
            $problems=$attributes.problems
            for ($ProdCount = 0; $ProdCount -lt $problems.Count; $ProdCount++) {
                $vulnerability = $problems[$ProdCount]
                $createdate = $attributes.created_at
                $issueage = New-TimeSpan -Start $createdate -End $currentdate
                $ProdInfo = Get-ProdInfo
                $ProdInfo.ID = $proditem.ID
                $ProdInfo.Title = $attributes.title
                $ProdInfo.Ignored = $attributes.Ignored
                $ProdInfo.Severity = $attributes.effective_severity_level
                $ProdInfo.CreatedTime = $attributes.created_at
                $ProdInfo.vulnerabilityID = $vulnerability.ID
                $ProdInfo.ProjectId = $proditem.relationships.scan_item.data.ID
                $ProdInfo.vulnerabilityAge = $issueage.Days
                $Prodresult += $ProdInfo
            }            
        }
        Write-Host $i ": "$daysbefore "--"$severity
        $i++    
    }    
}
$past90days = $Devresult |Where-Object{$_.vulnerabilityAge -le 90} |Sort-Object -Property CreatedTime
$Past180days = $Devresult |Where-Object{$_.vulnerabilityAge -le 180 -and $_.vulnerabilityAge -gt 90}|Sort-Object -Property CreatedTime
$Past360days = $Devresult |Where-Object{$_.vulnerabilityAge -le 360 -and $_.vulnerabilityAge -gt 180}|Sort-Object -Property CreatedTime
$Past720days = $Devresult |Where-Object{$_.vulnerabilityAge -le 720 -and $_.vulnerabilityAge -gt 360}|Sort-Object -Property CreatedTime
$Past900days = $Devresult |Where-Object{$_.vulnerabilityAge -gt 720}|Sort-Object -Property CreatedTime
$details = @("Less than 3 months old","3 to 6 months old","6 to 12 months old","12 to 36 months old","Over 36 months old")
$Pastinfo = @($past90days,$Past180days,$Past360days,$Past720days,$Past900days)
$vulnerabilitydata =@()
foreach($severity in $severities){
    $summaryreport += "<tr style='font-size:12px;' nowrap>
            <td align ='Center' nowrap><a>$($severity)</a></td>
            <td align ='Center' nowrap></td> 
            <td align ='Center' nowrap></td> 
            <td align ='Center' nowrap></td>
            <td align ='Center' nowrap></td> 
        </tr>"
    
    for ($i = 0; $i -lt 5; $i++) {
        $ignored = $Pastinfo[$i]|Where-Object{$_.Ignored -eq $true -and $_.Severity -eq $severity}|Sort-Object -Property ProjectId,CreatedTime
        # $sorteddata = $Pastinfo[$i]|Where-Object{$_.Severity -eq $severity}|Sort-Object -Property ProjectId,created_at
        $vulnerabilitydata += $ignored
        $Notfixeditems = @()
        $FixedNotReleased = @()
        foreach($item in $Pastinfo[$i]){
            $Notfixeditems += $item |Where-Object{$_.ProjectId -ne $Prodresult.ProjectId -and $_.Ignored -eq $false -and $_.Severity -eq $severity}
            $FixedNotReleased += $item|Where-Object{$_.ProjectId -eq $Prodresult.ProjectId -and $_.Ignored -eq $false -and $_.Severity -eq $severity}
            
        }  
        $vulnerabilitydata += $Notfixeditems 
        $vulnerabilitydata += $FixedNotReleased
        $totalcount =$Pastinfo[$i] |Where-Object{$_.Severity -eq $severity}
        $summaryreport += "<tr style='font-size:12px;' nowrap>
            <td align ='Center' nowrap>&nbsp;$($details[$i])&nbsp;</a></td>
            <td align ='Center' nowrap>&nbsp;$($ignored.Count)&nbsp;</td> 
            <td align ='Center' nowrap>&nbsp;$($Notfixeditems.Count)&nbsp;</td> 
            <td align ='Center' nowrap>&nbsp;$($FixedNotReleased.Count)&nbsp;</td>
            <td align ='Center' nowrap>&nbsp;$($totalcount.Count)&nbsp;</td> 
        </tr>"  
    }
    
}
foreach($severity in $severities){
    $severityreport += "<br><b>$($severity)-Severity issues</b><br>"
    $severityreport += "<tbody><table border='1' cellspacing='1' style='font-size:10px; border:1px solid black; border-collapse: collapse' nowrap>
        <tr style='font-size:12px;'>
            <th valign='top' nowrap><b>IssueId</b></th>
            <th valign='top' nowrap><b>Issue Description</b></th>
            <th valign='top' nowrap><b># of Instances(Projects)</b></th>
            <th valign='top' nowrap><b>Publication Date</b></th>
            <th valign='top' nowrap><b>Issue age(days)</b></th>
            <th valign='top' nowrap><b>Is Fixable</b></th>
            <th valign='top' nowrap><b>Is Upgradable</b></th>
            <th valign='top' nowrap><b>Is Ignored</b></th>
            <th valign='top' nowrap><b>Is Fixed But Not<br>Released to Production</b></th>
        </tr>"
    $vulnerabilityinfo = $vulnerabilitydata|Where-Object{$_.Severity -eq $severity}|Sort-Object -Property Ignored
    $vulnerabilitiesinfo=$vulnerabilityinfo|Group-Object -Property vulnerabilityID|Sort-Object -Property Count -Descending
    $vulnerabilityids = $vulnerabilitiesinfo.Name
    foreach($vulnerabilityid in $vulnerabilityids){
        $vulnerability= $vulnerabilityinfo | Where-Object{$_.vulnerabilityID -eq $vulnerabilityID}
        $title = $vulnerability.title|Select-Object -Unique
        $projectcount = $vulnerability.ProjectId|Select-Object  -Unique
        $createdtime = $vulnerability.CreatedTime|Select-Object -Unique
        $vulnerabilityAge = $vulnerability.vulnerabilityAge|Select-Object -Unique |Sort-Object
        $fixable = $vulnerability.Fixable|Select-Object -Unique|Sort-Object -Descending
        $Upgradable = $vulnerability.Upgradable|Select-Object -Unique |Sort-Object -Descending
        $IsIgnored = $vulnerability.Ignored|Select-Object -Unique |Sort-Object -Descending
        $projectsCount =$projectcount.Count
        [datetime]$time= $createdtime[0]
        $issueage=$vulnerabilityAge[0]
        $IsFixable = $Fixable[0]
        $Isupgradable = $Upgradable[0]
        $IsIgnorable = $IsIgnored[0]
        $IsReleased = @()
        foreach($item in $vulnerability){
            if ($item.ProjectId -eq $Prodresult.ProjectId) {
                $IsReleased += $true
            }
            else {
                $IsReleased += $false
            }
        }
        $Released = $IsReleased|Sort-Object -Descending
        $FixReleased = $Released[0]            
        $severityreport += "<tr style='font-size:12px;' nowrap>
                <td align ='Center' nowrap>&nbsp;$($vulnerabilityid)&nbsp;</a></td>
                <td align ='Center' nowrap>&nbsp;$($title)&nbsp;</td> 
                <td align ='Center' nowrap>&nbsp;$($projectsCount)&nbsp;</td> 
                <td align ='Center' nowrap>&nbsp;$($time)&nbsp;</td>
                <td align ='Center' nowrap>&nbsp;$($issueage)&nbsp;</td> 
                <td align ='Center' nowrap>&nbsp;$($IsFixable)&nbsp;</td>
                <td align ='Center' nowrap>&nbsp;$($Isupgradable)&nbsp;</td>
                <td align ='Center' nowrap>&nbsp;$($IsIgnorable)&nbsp;</td>
                <td align ='Center' nowrap>&nbsp;$($FixReleased )&nbsp;</td>
            </tr>" 
    }
    $severityreport += "</table></tbody>"
}
$summaryreport += "<tr style='font-size:12px;' nowrap>
            <td align ='Center' nowrap>&nbsp;Total&nbsp;</a></td>
            <td align ='Center' nowrap>&nbsp;&nbsp;</td> 
            <td align ='Center' nowrap>&nbsp;&nbsp;</td> 
            <td align ='Center' nowrap>&nbsp;&nbsp;</td>
            <td align ='Center' nowrap>&nbsp;&nbsp;</td> 
        </tr>" 
for ($i = 0; $i -lt 5; $i++) {
    $ignored = $Pastinfo[$i] | Where-Object{$_.Ignored -eq $true}
    $TotalNotFixed = @()
    $Totalfixable = @()
    foreach($item in $Pastinfo[$i]){
        $TotalNotFixed += $item |Where-Object{$_.ProjectId -ne $Prodresult.ProjectId -and $_.Ignored -eq $false }
        $Totalfixable += $item|Where-Object{$_.ProjectId -eq $Prodresult.ProjectId -and $_.Ignored -eq $false }
            
    }

    $summaryreport += "<tr style='font-size:12px;' nowrap>
                <td align ='Center' nowrap>&nbsp;$($details[$i])&nbsp;</a></td>
                <td align ='Center' nowrap>&nbsp;$($ignored.Count)&nbsp;</td> 
                <td align ='Center' nowrap>&nbsp;$($TotalNotFixed.Count)&nbsp;</td> 
                <td align ='Center' nowrap>&nbsp;$($Totalfixable.Count)&nbsp;</td>
                <td align ='Center' nowrap>&nbsp;$($Pastinfo[$i].Count)&nbsp;</td> 
            </tr>"  
}

$summaryreport += "</table></tbody>"
$body += $summaryreport
$body += $severityreport
# # compose and send out email messages to individuals who have booked to at least one task in the sprint
Write-Host "-------------------------------------------------------------------------------------------"
Write-Host "Sending out email " 
$emails=("")
$emailSubject = "Snyk Report"
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
    $message.Bcc.Add($mailid)
}
foreach($email in $emails)
{
    $message.To.Add($email)
}

$message.Subject = $emailSubject
$message.Body = $body
# $message.cc.Add($bccEmailList)
$message.IsBodyHtml = $true
# Send-MailMessage -SmtpServer $SMTP_SERVER -Port $SMTP_PORT -Credential $cred -To $email -Bcc $bccEmailList -From " " -Subject $emailSubject -Body $body -BodyAsHtml
try {
    # Send-MailMessage -To $email -SmtpServer $SMTP_SERVER -Port $SMTP_PORT -Credential $cred -Bcc $bccEmailList -From " " -Subject $emailSubject -Body $body -BodyAsHtml
    $SMTPClient.Send($message)
    $SMTPClient.Dispose()
    $message.Dispose()
    $body = $null
    $summaryreport = $null
    $severityreport = $null
    
}
catch {
    $_.Exception.message
}
