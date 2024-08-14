param ([string]$info="Password")
Get-Date
$API_Key = "API Key"
$API_Secret ="Secret key"


$headers = @{
    Authorization = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("$($API_Key):$($API_Secret)"))
}
function Get-FivetranConnectorInfo {
    $FivetranConnectorInfo = New-Object -TypeName PSObject
    $FivetranConnectorInfo | Add-Member -MemberType NoteProperty -Name GroupName -Value $null
    $FivetranConnectorInfo | Add-Member -MemberType NoteProperty -Name ConnectorName -Value $null
    $FivetranConnectorInfo | Add-Member -MemberType NoteProperty -Name Service -Value $null
    $FivetranConnectorInfo | Add-Member -MemberType NoteProperty -Name DestinationName -Value $null
    $FivetranConnectorInfo | Add-Member -MemberType NoteProperty -Name Frequency -Value $null
    

    return $FivetranConnectorInfo
    
}
# $fivetranuri = "https://api.fivetran.com/v1/metadata/connector-types"
# $fivetranResponse = Invoke-RestMethod -Uri $fivetranuri -Headers $headers -Method Get
$ConnectorSummaryHtml = "<h1>Schedule configurations for Fivetran connectors</h1>"

$ConnectorSummaryHtml += "<tbody><table border='1' cellspacing='1' cellpadding='7' style='font-size:10px; border:1px solid black; border-collapse: collapse' nowrap>
    <tr style='font-size:12px;'>
        <th valign='top' nowrap><b>Group Name</b></th>
        <th valign='top' nowrap><b>Connector Name</b></th>
        <th valign='top' nowrap><b>Service</b></th>
        <th valign='top' nowrap><b>Destination Name</b></th>
        <th valign='top' nowrap><b>Frequency of execution</b></th>
    </tr>
"
$groupuri="https://api.fivetran.com/v1/groups"
$groupresponse=Invoke-RestMethod -Uri $groupuri -Headers $headers -Method Get
$groupids = $groupresponse.data.items.Id
$ConnectorInfo = @()
foreach($groupid in $groupids){
    $connectoruri= "https://api.fivetran.com/v1/groups/$($groupid)/connectors"
    $connectorresponse=Invoke-RestMethod -Uri $connectoruri -Headers $headers -Method Get
    $connectoritems = $connectorresponse.data.items
    $destinationNamefilter = $groupresponse.data.items |Where-Object {$_.id -eq $groupid}
    $destinationName = $destinationNamefilter.name
    foreach($connectoritem in $connectoritems){
        
        $FivetranConnectorInfo = Get-FivetranConnectorInfo
        $FivetranConnectorInfo.GroupName = $groupid
        $FivetranConnectorInfo.ConnectorName = $connectoritem.id
        $FivetranConnectorInfo.Service = $connectoritem.service
        $FivetranConnectorInfo.DestinationName= $destinationName
        $FivetranConnectorInfo.Frequency = $connectoritem.sync_frequency
        $ConnectorInfo += $FivetranConnectorInfo
    }
}
$sortedConnectorInfo = $ConnectorInfo |Sort-Object -Property Frequency -Descending
foreach($Item in $sortedConnectorInfo){
    $frequencytime = $Item.Frequency
    $frequencytimeinseconds = [timespan]::fromseconds($frequencytime)
    $frequencytimeinhours = "$($frequencytimeinseconds.hours) Hrs : $($frequencytimeinseconds.minutes) Mins : $($frequencytimeinseconds.seconds) Secs"
    $ConnectorSummaryHtml += "<tr nowrap>
                <td nowrap>$($Item.GroupName)</a></td>
                <td nowrap>$($Item.ConnectorName)</a></td>
                <td nowrap>$($Item.Service)</td>
                <td nowrap>$($Item.DestinationName)</td>
                <td nowrap>$($frequencytimeinhours)</td>
            </tr>"
}

$ConnectorSummaryHtml += "</table></tbody>"
$body += $ConnectorSummaryHtml
# # compose and send out email messages to individuals who have booked to at least one task in the sprint
Write-Host "-------------------------------------------------------------------------------------------"
Write-Host "Sending out the schedule configurations for Fivetran connectors"
$emailSubject = "Schedule configurations for Fivetran connectors"
Start-Sleep -s 15

$SMTP_SERVER = "smtp.socketlabs.com"
$SMTP_PORT = 587
$SMTP_USERNAME = "server4507"
$PWORD = ConvertTo-SecureString -String $info -AsPlainText -Force
$bccEmailList = ("")
$email = ("")
# $email = ("")
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
foreach($emailid in $email)
{
    $message.To.Add($emailid)
}
$message.CC.Add($bccEmailList)
$message.Subject = $emailSubject
$message.Body = $body
$message.IsBodyHtml = $true
try {
    $SMTPClient.Send($message)
    $SMTPClient.Dispose()
    $message.Dispose()
    $body = $null
    $ConnectorSummaryHtml = $null
    # $PullRequestData = $null

}
catch {
    $_.Exception.message
}

Get-date