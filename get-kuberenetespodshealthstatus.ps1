function Create-PodInfo
{
    $podInfo = New-Object -TypeName PSObject
    $podInfo | Add-Member -MemberType NoteProperty -Name name -Value $null
    $podInfo | Add-Member -MemberType NoteProperty -Name environment -Value $null
    $podInfo | Add-Member -MemberType NoteProperty -Name namespace -Value $null
    $podInfo | Add-Member -MemberType NoteProperty -Name startTime -Value $null
    $podInfo | Add-Member -MemberType NoteProperty -Name containerStatus -Value $null
    $podInfo | Add-Member -MemberType NoteProperty -Name imageTag -Value $null

    return $podInfo
}

$date = Get-Date
$utcDate = ([datetime]$date).ToUniversalTime()
$currentDateTime = $utcDate.ToString("dd-MM-yyyy hh:mm:ss tt")

$WorkspaceID="workspace ID"
$query = 'KubePodInventory| join (ContainerInventory | project ContainerID, ImageTag, ImageID) on $left.ContainerID == $right.ContainerID| where TimeGenerated > now(-5m)| where Namespace != "kube-system"| where ClusterName contains "apxdevtest-usw-aks"| summarize arg_max(TimeGenerated, *) by Name| order by Namespace asc, TimeGenerated desc'
# | where PodStatus !contains "Running"

$InsightsQuery = Invoke-AzOperationalInsightsQuery -WorkspaceId $WorkspaceID -Query $query
$InsightsQueryResult = $InsightsQuery.Results

# Extract unique namespaces from the result
$Namespaces = $InsightsQueryResult.Namespace | Sort-Object -Unique
$Environments = @()
# Extract the last part of the namespace and use it as column headers
foreach ($Namespace in $Namespaces) {
    $environment = $Namespace.split('-')[-1]
    if ($environment.IndexOf("dashboard") -eq -1 -and $environment.IndexOf("internal") -eq -1 -and $environment.IndexOf("external") -eq -1) {
        $Environments += $environment
    }
}
$Environments = $Environments | Sort-Object -Unique

# Get unique service names
$ServiceNames = $InsightsQueryResult.ServiceName | Sort-Object -Unique 

$PodsByServiceName = @{}

# For each service, establish the listing of pods in each environment. This will be used later to render the HTML table.
foreach($Pod in $InsightsQueryResult) {
    $ServiceName = $Pod.ServiceName
    if ($null -ne $ServiceName -and $ServiceName -ne "" -and $ServiceName.IndexOf("internal") -eq -1 -and $ServiceName.IndexOf("external") -eq -1) {
        $PodsByEnvironment = $PodsByServiceName.Item($ServiceName)
        if ($null -eq $PodsByEnvironment) {
            $PodsByEnvironment = @{}
            $PodsByServiceName.Add($ServiceName, $PodsByEnvironment)
        }
        # get environment name for the current pod
        $Namespace = $Pod.Namespace
        $Environment = $Namespace.split('-')[-1]
        $PodInfo = Create-PodInfo
        $PodInfo.name = $Pod.Name
        $PodInfo.environment = $Environment
        $PodInfo.Namespace = $Pod.Namespace
        $PodInfo.startTime = $Pod.PodStartTime
        $PodInfo.containerStatus = $Pod.ContainerStatus
        $PodInfo.imageTag = $Pod.ImageTag
        $PodInfos = $PodsByEnvironment.Item($Environment)
        If ($null -eq $PodInfos) {
            $PodInfos = New-Object System.Collections.ArrayList
            $PodsByEnvironment.Add($Environment, $PodInfos)
        }
        # the below statement insures that for every service, when the HTML table is rendered, the pods in 'running' state
        # will be shown at the top
        if ($PodInfo.containerStatus -eq "running") {
            [void]$PodInfos.Insert(0, $PodInfo)
        }
        else {
            [void]$PodInfos.Add($PodInfo)
        }
    }
    
}

$report = @"
<!doctype html>
<html>
    <title>Kubernetes Pods Status Page</title>
    <head>
        <!-- Required meta tags -->
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no" >
        <meta http-equiv="refresh" content="30" />
        <style>
            /* Custom CSS to freeze the header row */
            table {
                border-collapse: collapse;
                width: 100%;
            }
            th, td {
                padding: 8px;
                text-align: left;
                border-style: none;
            }

            th {
                background-color: #f2f2f2;
                position: sticky;
                top: 0;
                z-index: 1;
            }
            td {
                height: 20px;
                width: 40px;
            }
            .expanded {
                height: 50px;
            }
        </style>
        <script>
            function expandCell(cell){
                cell.classList.toggle('expanded');
            }
        </script>
        <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/css/bootstrap.min.css" integrity="sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm" crossorigin="anonymous">
    </head>
    <body>
        <div class="container-fluid">
            <h1>The data is as of: $($currentDateTime) (UTC)</h1>
            <h3>Azure Log Analytics for last 5 minutes. And this page will auto refresh every 30 seconds</h3>
            <table class="table table-striped table-hover table-sm" border="1">
                <thead>
                    <tr>
                        <th>Service Name</th>
                        
"@
# <th>Pod Name</th>
foreach ($Environment in $Environments) {
    $report += "<th>$Environment</th>"
}

$report += @"
                    </tr>
                </thead>
                <tbody>
"@

$ServiceNames = $PodsByServiceName.Keys | Sort-Object
foreach($ServiceName in $ServiceNames) {
    # determine the number of table rows that will be needed for this service. 
    # the number of table rows will equal the max number of pods among the environments 
    # in which the service is deployed
    $maxPods = 0
    $PodsByEnvironment = $PodsByServiceName.Item($ServiceName)
    foreach($Environment in $Environments) {
        $podInfos = $PodsByEnvironment.Item($Environment)
        if ($null -ne $podInfos) {
            if ($podInfos.Count -gt $maxPods) {
                $maxPods = $PodInfos.Count
            }
        }
    }
    Write-Host "service name: $ServiceName; max pods: $maxPods"
    # now, render the rows. the first column - Service Name - spans all rows for that service. 
    $counter = 0
    for ($counter = 0; $counter -lt $maxPods; $counter++)
    {
        $report += "`n<tr>"
        if ($counter -eq 0) {
            $report += "`n`t<td rowspan=`"$maxPods`" style='font-size:14px' ><b>$ServiceName</b></td>"
        }
        foreach($Environment in $Environments) {    
            [System.Collections.ArrayList]$PodsForEnvironment = [System.Collections.ArrayList]$PodsByEnvironment.item($Environment)
            if ($null -ne $PodsForEnvironment) {
                if ($counter -lt $PodsForEnvironment.Count) {
                    $PodInfo = $PodsForEnvironment[$counter]
                    if ($PodInfo.containerStatus -match "fail") {
                        $color = "#ff0000"
                        $fontcolor = "#000000"
                    }
                    elseif ($PodInfo.containerStatus -match "terminate") {
                        $color = "#FFFFFF"
                        $fontcolor = "#808080"
                    }
                    else {
                        $color = "#FFFFFF"
                        $fontcolor = "#000000"
                    }
                }
            }
            else {
                $color = "#FFFFFF"
            }
            $report += "`n`t<td bgcolor=$color style='font-size:12px;color:$fontcolor' >"
            if ($null -ne $PodsForEnvironment) {
                if ($counter -lt $PodsForEnvironment.Count) { 
                    $PodInfo = $PodsForEnvironment[$counter]
                    $report += ("<nobr><b>Pod Name</b>: " + $PodInfo.name + "</nobr>")
                    $report += ("<br><nobr><b>Status</b>: " + $PodInfo.containerStatus + "</nobr>")
                    $report += ("<br><nobr><b>Namespace</b>: " + $PodInfo.namespace + "</nobr>")
                    $report += ("<br><nobr><b>Start Time</b>: " + $PodInfo.startTime + "</nobr>")
                    $report += ("<br><nobr><b>Image Tag</b>: " + $PodInfo.imageTag + "</nobr>")
                }
            }
            $report += "</td>"    
        }         
        $report += "`n</tr>"
    }
}

$report += @"
                </tbody>
            </table>
        </div>
    </body>
</html>
"@

$report | Out-File -FilePath "KQueryReport.html"
