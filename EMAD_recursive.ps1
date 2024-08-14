$jiraPATId=""




$pair = "$(""):$($jiraPATId)"
$headers = @{
    Authorization = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($pair)"))
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
# Get-parentInfo -issueid $issueid -Childtask $issueid