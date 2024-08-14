param(
    [Parameter(Mandatory=$true)]
    [string]$ChangeRequestType,
    [string]$Server,
    [string]$IPAddress,
    [string]$UserName,
    [string]$QuarterYear,
    [string]$Servers
)
#Info.ps1 file path to obtain PAT token
$scriptDirectory = ($PWD.Path | Split-Path -Parent)
$Path = (Join-Path (Join-Path $scriptDirectory -ChildPath "common") -ChildPath "info.ps1")
$infoLibContent = Get-Content -Path $Path -Raw
Invoke-Expression -Command $infoLibContent

# # Load and execute the common library script
$CommonLibPath = (Join-Path (Join-Path $scriptDirectory -ChildPath "common") -ChildPath "common.ps1")
$commonLibContent = Get-Content -Path $CommonLibPath -Raw
Invoke-Expression -Command $commonLibContent

$latestSprint =Get-SprintInfo -forDate (Get-Date)
$latestSprintNumber = $latestSprint.sprintNumber
$presentsprint = "$($Project)\\Sprint " + $latestSprintNumber

$headers = @{
    Authorization = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($PATId)"))
}

# Function to validate IPv4 address
function Confirm-IPAddress {
    param([string]$IPAddress)
    [bool] $IsValid = $IPAddress -match '^(\d{1,3}\.){3}\d{1,3}$'
    return $IsValid
}
 
# Function to create PBI and tasks
function New-PBI {
    param(
        [string]$AssignedTo,
        [string]$PBITitle,
        [string]$DescriptionContent,
        [string]$AcceptanceCriteria,
        [string]$workitemType,
        [string]$ParentID
    )
    $parentIduri = [uri]::EscapeUriString("https://$($organization).visualstudio.com/DefaultCollection/$($Project)/_apis/wit/workItems/$($ParentID)")
    # Create PBI
    Write-Output "Creating PBI..."
    $PBI = "
    [
        {
            ""op"": ""add"",
            ""path"": ""/fields/System.Title"",
            ""from"": null,
            ""value"": ""$PBITitle""
        }
        ,
        {
            ""op"": ""add"",
            ""path"": ""/fields/System.Description"",
            ""from"": null,
            ""value"": ""$DescriptionContent""
        }
        ,
        {
            ""op"": ""add"",
            ""path"": ""/fields/Microsoft.VSTS.Common.AcceptanceCriteria"",
            ""from"": null,
            ""value"": ""$AcceptanceCriteria""
        }
        ,
        {
            ""op"": ""add"",
            ""path"": ""/fields/System.WorkItemType"",
            ""from"": null,
            ""value"":""Product Backlog Item"",
        }
        ,
        {
            ""op"": ""add"",
            ""path"": ""/fields/System.AreaPath"",
            ""from"": null,
            ""value"":""$($Project)\\Shared Infrastructure"",
        }
        ,
        {
            ""op"": ""add"",
            ""path"": ""/fields/System.IterationPath"",
            ""from"": null,
            ""value"":""$presentsprint"",
        }
        ,
        {   
            ""op"": ""add"",
            ""path"": ""/fields/System.state"",
            ""value"": ""New""
        }  
        ,
        {
            ""op"": ""add"",
            ""path"": ""/fields/System.AssignedTo"",
            ""value"": ""$AssignedTo""
        }   
        ,
        {
             ""op"": ""add"",
             ""path"": ""/relations/-"",
            ""value"": {
                ""rel"": ""System.LinkTypes.Hierarchy-Reverse"",
                ""url"": ""$parentIduri""
            },
        } 
    ]
"

    #$body = $task | ConvertTo-Json
    $body = $PBI
    #Write-Host "Body: $body"
    $PBIuri = [uri]::EscapeUriString("https://dev.azure.com/$($organization)/$($Project)/_apis/wit/workitems/`$Product Backlog Item?api-version=7.-preview.3")
    #creating new child task
    #Write-Host "body: $body"
    try {
        #Invoke a REST API call for each task to be created
        Write-Host "URI: $PBIuri"
        $result = Invoke-RestMethod -Uri $PBIuri -Method Post -ContentType "application/json-patch+json" -Headers $headers -Body $body
        $result
    }
    catch {
        $_.Exception.Message
    }  
    # Display information
    Write-Output "PBI created with title: $PBITitle"
}
function New-Tasks {
    param(
        [string]$AssignedTo,
        [string]$PBITitle,
        [string]$DescriptionContent,
        [string]$Skillset,
        [string]$workitemType,
        [string]$ParentID
    )
    $parentIduri = [uri]::EscapeUriString("https://$($organization).visualstudio.com/DefaultCollection/$($Project)/_apis/wit/workItems/$($ParentID)")
    # Create PBI
    Write-Output "Creating Task..."
    $task = "
    [
        {
            ""op"": ""add"",
            ""path"": ""/fields/System.Title"",
            ""from"": null,
            ""value"": ""$PBITitle""
        }
        ,
        {
            ""op"": ""add"",
            ""path"": ""/fields/System.Description"",
            ""from"": null,
            ""value"": ""$DescriptionContent""
        }
        ,
        {
            ""op"": ""add"",
            ""path"": ""/fields/System.WorkItemType"",
            ""from"": null,
            ""value"":""Task"",
        }
        ,
        {   
            ""op"": ""add"",
            ""path"": ""/fields/System.state"",
            ""value"": ""To do""
        }
        ,
        {
            ""op"": ""add"",
            ""path"": ""/fields/System.AreaPath"",
            ""from"": null,
            ""value"":""$($Project)\\Shared Infrastructure"",
        }
        ,
        {
            ""op"": ""add"",
            ""path"": ""/fields/System.IterationPath"",
            ""from"": null,
            ""value"":""$presentsprint"",
        }   
        ,
        {
            ""op"": ""add"",
            ""path"": ""/fields/System.AssignedTo"",
            ""value"": ""$AssignedTo""
        }
        ,
        {
            ""op"": ""add"",
            ""path"": ""/fields/System.Tags"",
            ""from"": null,
            ""value"":""Change Request"",
        }
        ,
        {
            ""op"": ""add"",
            ""path"": ""/fields/Microsoft.VSTS.Common.Activity"",
            ""from"": null,
            ""value"":""Deployment"",
        }
        ,
        {
            ""op"": ""add"",
            ""path"": ""/fields/Custom.Skillset"",
            ""from"": null,
            ""value"":""$skillset"",
        }  
        ,
        {
             ""op"": ""add"",
             ""path"": ""/relations/-"",
            ""value"": {
                ""rel"": ""System.LinkTypes.Hierarchy-Reverse"",
                ""url"": ""$parentIduri""
            },
        } 
        ,
        {
            ""op"": ""add"",
            ""path"": ""/fields/Microsoft.VSTS.Scheduling.RemainingWork"",
            ""value"": ""0.25""
        }
        ,
        {
            ""op"": ""add"",
            ""path"": ""/fields/$($organization).EMAScrum.OriginalHours"",
            ""value"": ""0.25""
        }
    ]
"

    #$body = $task | ConvertTo-Json
    $body = $task
    #Write-Host "Body: $body"
    $taskuri = [uri]::EscapeUriString("https://dev.azure.com/$($organization)/$($Project)/_apis/wit/workitems/`$task?api-version=7.-preview.3")
    #creating new child task
    #Write-Host "body: $body"
    try {
        #Invoke a REST API call for each task to be created
        Write-Host "URI: $taskuri"
        $result = Invoke-RestMethod -Uri $taskuri -Method Post -ContentType "application/json-patch+json" -Headers $headers -Body $body
        $result
    }
    catch {
        $_.Exception.Message
    }
 
    # Display information
    Write-Output "Task is created with title: $PBITitle"
}
function New-ParentPBI {
    param (
    [Parameter(Mandatory=$true)]
    [string]$UserName,
    [Parameter(Mandatory=$true)]
    [string]$Server,
    [Parameter(Mandatory=$true)]
    [string]$IpAddress,
    [Parameter(Mandatory=$true)]
    [string]$changerequesttype,
    [string]$ParentID
)
    if($changerequesttype -eq "AzureDBIPWhitelisting"){
        try {
            if ($server -eq "$($Project)-use-db" -or $server -eq "$($Project)devtest-nln-db") {
                $Title = "Shared Infrastructure RTB - Whitelist IP address "+ $IpAddress +" for User " + $UserName
                $Description= "Shared Infrastructure RTB - Whitelist IP address "+ $IpAddress +" for Server " +$Server
                $AcceptanceCriteria = "IP address $IPAddress is whitelisted on server $Server."
                $parentIduri = [uri]::EscapeUriString("https://$($organization).visualstudio.com/DefaultCollection/$($Project)/_apis/wit/workItems/$($ParentID)")
                # Example function to create a task
                Write-Output "Creating PBI..."
                $task = "
                [
                    {
                        ""op"": ""add"",
                        ""path"": ""/fields/System.Title"",
                        ""from"": null,
                        ""value"": ""$Title""
                    }
                    ,
                    {
                        ""op"": ""add"",
                        ""path"": ""/fields/System.Description"",
                        ""from"": null,
                        ""value"": ""$Description""
                    }
                    ,
                    {
                        ""op"": ""add"",
                        ""path"": ""/fields/Microsoft.VSTS.Common.AcceptanceCriteria"",
                        ""from"": null,
                        ""value"": ""$AcceptanceCriteria""
                    }
                    ,
                    {
                        ""op"": ""add"",
                        ""path"": ""/fields/System.WorkItemType"",
                        ""from"": null,
                        ""value"":""Product Backlog Item"",
                    }
                    ,
                    {
                        ""op"": ""add"",
                        ""path"": ""/fields/System.AreaPath"",
                        ""from"": null,
                        ""value"":""$($Project)\\Shared Infrastructure"",
                    }
                    ,
                    {
                        ""op"": ""add"",
                        ""path"": ""/fields/System.IterationPath"",
                        ""from"": null,
                        ""value"":""$presentsprint"",
                    }
                    ,
                    {
                        ""op"": ""add"",
                        ""path"": ""/fields/System.state"",
                        ""value"": ""New""
                    }
                    ,
                    {
                        ""op"": ""add"",
                        ""path"": ""/fields/System.AssignedTo"",
                        ""value"": ""Tarak""
                    }
                    ,
                    {
                        ""op"": ""add"",
                        ""path"": ""/relations/-"",
                        ""value"": {
                            ""rel"": ""System.LinkTypes.Hierarchy-Reverse"",
                            ""url"": ""$parentIduri""
                        },
                    } 
                    
                ]
    "
    
                #$body = $task | ConvertTo-Json
                $body = $task
                #Construct the URI
                $uri = [uri]::EscapeUriString("https://dev.azure.com/$($organization)/$($Project)/_apis/wit/workitems/`$Product Backlog Item?api-version=7.-preview.3")
                #creating new child task
                #Write-Host "body: $body"
                try {
                    #Invoke a REST API call for each task to be created
                    Write-Host "URI: $uri"
                    $DBtaskresult = Invoke-RestMethod -Uri $uri -Method Post -ContentType "application/json-patch+json" -Headers $headers -Body $body
                    $DBtaskresult
                }
                catch {
                    $_.Exception.Message
                }
            }
        }
        catch {
           $_.Exception.message
           Exit-PSHostProcess
        }
    }
    Write-Output "PBI is created with title: $Title"
}
# Main script logic
if ($ChangeRequestType -eq 'AzureDBIPWhitelisting') {
    # Validate Server parameter
    if ($Server -notin @('$($Project)-use-db', '$($Project)devtest-nln-db')) {
        
        Write-Error "Invalid value for Server parameter. Allowed values are '$($Project)-use-db' and '$($Project)devtest-nln-db'."
        return
    }
    else {
        # Validate IPAddress parameter
        if (-not (Confirm-IPAddress $IPAddress)) {
            Write-Error "Invalid IPv4 address specified for IPAddress parameter."
            return
        }
        else {
            $parentPBI= New-ParentPBI -UserName $UserName -IpAddress $IPAddress -Server $Server -changerequesttype $ChangeRequestType -ParentID "32420"
            $parentPBI
        }
        # Create PBI for AzureDBIPWhitelisting
        $id =$parentPBI.id
        # Create tasks for AzureDBIPWhitelisting
        $TaskTitle = "Shared Infrastructure RTB - Whitelist IP address for user $UserName on server $Server"
        $TaskDescription = "Deployment window start date/time: `n`r" + "`nDeployment window end date/time:"
        New-Tasks -AssignedTo 'Tarak' -PBITitle $TaskTitle -DescriptionContent $TaskDescription -Skillset "Other" -ParentID $id
    
       
    }    
}
elseif ($ChangeRequestType -eq 'QuarterlyServerPatching') {
    # Validate QuarterYear parameter
    if ($QuarterYear -notmatch '^Q[1-4] \d{4}$') {
        Write-Error "Invalid format for QuarterYear parameter. Correct format is 'Q1 2024'."
        return
    }
    $AllowedServers = ("EM Non-Prod","PM Non-Prod","EM DR","EM Prod","PM Geo","PM Prod")
    if ($Server -in $AllowedServers) {
        $parentPBI= New-ParentPBI -UserName $UserName -IpAddress $IPAddress -Server $Server -changerequesttype $ChangeRequestType
        $parentPBI
    }
    else {
        Write-Host "Please enter valid server details"
    }
    if ($Server -like "EM") {
        if ($server -eq "EM Non-Prod") {
            $EmNonProd = ("qqqdqwwf-db4", "qqqdqwwf-app")
            if ($value -in $EmNonProd) {
                # Create PBI for QuarterlyServerPatching
                $PBITaskTitle = "Shared Infrastructure RTB - Perform quarterly server patching for " +($value -join ",")+  " servers for $QuarterYear"
                $PBITaskDescription = "Quarterly patching of " +($value -join ",")+  " servers for $QuarterYear"
                $PBITaskAcceptanceCriteria ="Quarterly patching is completed on the following servers: `n" +($value -join ",`n")
                $parentTask =New-PBI -AssignedTo 'Alex Gurvits' -workitemType "Product Backlog Item" -PBITitle $PBITaskTitle -DescriptionContent $PBITaskDescription -AcceptanceCriteriaContent $PBITaskAcceptanceCriteria
                $parentTask
                $id =$parentTask.id
                # Create tasks for QuarterlyServerPatching to Ashok
                $TaskDescription = "Patching of database servers. `n"+"`nDeployment window start date/time: `n"+"`nDeployment window end date/time:"
                New-Tasks -AssignedTo 'Ashok Dubey' -PBITitle $PBITaskTitle -DescriptionContent $TaskDescription -Skillset "Database Administration" -ParentID $id

                # Create tasks for QuarterlyServerPatching to Pedro
                $TaskDescription = "Patching of application and/or web servers. `n"+"`nDeployment window start date/time: `n"+"`nDeployment window end date/time:"
                New-Tasks -AssignedTo 'Tarak' -PBITitle $PBITaskTitle -DescriptionContent $TaskDescription -Skillset "Other" -ParentID $id
            }
            else {
                Write-Host "Provided Server is not valid"
            }
            return
        }
        elseif ($server -eq "EM DR") {
            $EmDR = ("e", "p01")
            if ($value -in $EmDR) {
                # Create PBI for QuarterlyServerPatching
                $PBITaskTitle = "Shared Infrastructure RTB - Perform quarterly server patching for " +($value -join ",")+  " servers for $QuarterYear"
                $PBITaskDescription = "Quarterly patching of " +($value -join ",")+  " servers for $QuarterYear"
                $PBITaskAcceptanceCriteria ="Quarterly patching is completed on the following servers: `n" +($value -join ",`n")
                $parentTask =New-PBI -AssignedTo 'Tarak' -workitemType "Product Backlog Item" -PBITitle $PBITaskTitle -DescriptionContent $PBITaskDescription -AcceptanceCriteriaContent $PBITaskAcceptanceCriteria
                $parentTask
                $id =$parentTask.id
                # Create tasks for QuarterlyServerPatching to Ashok
                $TaskDescription = "Patching of database servers. `n"+"`nDeployment window start date/time: `n"+"`nDeployment window end date/time:"
                New-Tasks -AssignedTo 'Ashok' -PBITitle $PBITaskTitle -DescriptionContent $TaskDescription -Skillset "Database Administration" -ParentID $id

                # Create tasks for QuarterlyServerPatching to Pedro
                $TaskDescription = "Patching of application and/or web servers. `n"+"`nDeployment window start date/time: `n"+"`nDeployment window end date/time:"
                New-Tasks -AssignedTo 'Tarak' -PBITitle $PBITaskTitle -DescriptionContent $TaskDescription -Skillset "Other" -ParentID $id
            }
            else {
                Write-Host "Provided Server is not valid"
            }
            return
        }
        else {
            if ($server -eq "EM Prod") {
                $EmProd = ("e", "1")
                if ($value -in $EmProd) {
                    # Create PBI for QuarterlyServerPatching
                    $PBITaskTitle = "Shared Infrastructure RTB - Perform quarterly server patching for " +($value -join ",")+  " servers for $QuarterYear"
                    $PBITaskDescription = "Quarterly patching of " +($value -join ",")+  " servers for $QuarterYear"
                    $PBITaskAcceptanceCriteria ="Quarterly patching is completed on the following servers: `n" +($value -join ",`n")
                    $parentTask =New-PBI -AssignedTo 'Tarak' -workitemType "Product Backlog Item" -PBITitle $PBITaskTitle -DescriptionContent $PBITaskDescription -AcceptanceCriteriaContent $PBITaskAcceptanceCriteria
                $parentTask
                $id =$parentTask.id
                # Create tasks for QuarterlyServerPatching to Ashok
                $TaskDescription = "Patching of database servers. `n"+"`nDeployment window start date/time: `n"+"`nDeployment window end date/time:"
                New-Tasks -AssignedTo 'Ashok Dubey' -PBITitle $PBITaskTitle -DescriptionContent $TaskDescription -Skillset "Database Administration" -ParentID $id

                # Create tasks for QuarterlyServerPatching to Pedro
                $TaskDescription = "Patching of application and/or web servers. `n"+"`nDeployment window start date/time: `n"+"`nDeployment window end date/time:"
                New-Tasks -AssignedTo 'Tarak' -PBITitle $PBITaskTitle -DescriptionContent $TaskDescription -Skillset "Other" -ParentID $id
                }
                else {
                    Write-Host "Provided Server is not valid"
                }
                return
            }
            else {
                Write-Host "Provided Server is not valid"
            }
        }
    }
    elseif ($Server -like "PM") {
        if ($server -eq "PM Non-Prod") {
            $PmNonProd = ("dqwf-web", "dqwf-mid", "dqwf-db", "qqqwf-web", "qqqwf-mid", "qqqwf-db", "pmsb2-nln2-web", "pmsb2-nln2-mid", "pmsb2-nln2-db", "pmpp2-nln2-web", "pmpp2-nln2-mid", "pmpp2-nln2-db")
            if ($value -in $PmNonProd) {
                # Create PBI for QuarterlyServerPatching
                $PBITaskTitle = "Shared Infrastructure RTB - Perform quarterly server patching for " +($value -join ",")+  " servers for $QuarterYear"
                $PBITaskDescription = "Quarterly patching of " +($value -join ",")+  " servers for $QuarterYear"
                $PBITaskAcceptanceCriteria ="Quarterly patching is completed on the following servers: `n" +($value -join ",`n")
                $parentTask =New-PBI -AssignedTo 'Alex Gurvits' -workitemType "Product Backlog Item" -PBITitle $PBITaskTitle -DescriptionContent $PBITaskDescription -AcceptanceCriteriaContent $PBITaskAcceptanceCriteria
                $parentTask
                $id =$parentTask.id
                # Create tasks for QuarterlyServerPatching to Ashok
                $TaskDescription = "Patching of database servers. `n"+"`nDeployment window start date/time: `n"+"`nDeployment window end date/time:"
                New-Tasks -AssignedTo 'Ashok Dubey' -PBITitle $PBITaskTitle -DescriptionContent $TaskDescription -Skillset "Database Administration" -ParentID $id

                # Create tasks for QuarterlyServerPatching to Pedro
                $TaskDescription = "Patching of application and/or web servers. `n"+"`nDeployment window start date/time: `n"+"`nDeployment window end date/time:"
                New-Tasks -AssignedTo 'Tarak' -PBITitle $PBITaskTitle -DescriptionContent $TaskDescription -Skillset "Other" -ParentID $id
            }
            else {
                Write-Host "Provided Server is not valid"
            }
            return
        }
        elseif ($server -eq "PM Geo") {
            $PmNonProd = ("nln-nln-mid1", "nln-nln-mid2", "nln2-nln-db1", "nln2-nln-db2", "nln2-nln-web")
            if ($value -in $EmNonProd) {
                # Create PBI for QuarterlyServerPatching
                $PBITaskTitle = "Shared Infrastructure RTB - Perform quarterly server patching for " +($value -join ",")+  " servers for $QuarterYear"
                $PBITaskDescription = "Quarterly patching of " +($value -join ",")+  " servers for $QuarterYear"
                $PBITaskAcceptanceCriteria ="Quarterly patching is completed on the following servers: `n" +($value -join ",`n")
                $parentTask =New-PBI -AssignedTo 'Alex Gurvits' -workitemType "Product Backlog Item" -PBITitle $PBITaskTitle -DescriptionContent $PBITaskDescription -AcceptanceCriteriaContent $PBITaskAcceptanceCriteria
                $parentTask
                $id =$parentTask.id
                # Create tasks for QuarterlyServerPatching to Ashok
                $TaskDescription = "Patching of database servers. `n"+"`nDeployment window start date/time: `n"+"`nDeployment window end date/time:"
                New-Tasks -AssignedTo 'Ashok Dubey' -PBITitle $PBITaskTitle -DescriptionContent $TaskDescription -Skillset "Database Administration" -ParentID $id

                # Create tasks for QuarterlyServerPatching to Pedro
                $TaskDescription = "Patching of application and/or web servers. `n"+"`nDeployment window start date/time: `n"+"`nDeployment window end date/time:"
                New-Tasks -AssignedTo 'Tarak' -PBITitle $PBITaskTitle -DescriptionContent $TaskDescription -Skillset "Other" -ParentID $id
            }
            else {
                Write-Host "Provided Server is not valid"
            }
            return
        }
        else {
            if ($server -eq "PM Prod") {
                $nln = ("nln-use-mid1", "nln-use-mid2", "nln-use-web2", "nln2-use-db1", "nln2-use-db2")
                if ($value -in $nln) {
                    # Create PBI for QuarterlyServerPatching
                    $PBITaskTitle = "Shared Infrastructure RTB - Perform quarterly server patching for " +($value -join ",")+  " servers for $QuarterYear"
                    $PBITaskDescription = "Quarterly patching of " +($value -join ",")+  " servers for $QuarterYear"
                    $PBITaskAcceptanceCriteria ="Quarterly patching is completed on the following servers: `n" +($value -join ",`n")
                    $parentTask =New-PBI -AssignedTo 'Alex Gurvits' -workitemType "Product Backlog Item" -PBITitle $PBITaskTitle -DescriptionContent $PBITaskDescription -AcceptanceCriteriaContent $PBITaskAcceptanceCriteria
                    $parentTask
                    $id =$parentTask.id
                    # Create tasks for QuarterlyServerPatching to Ashok
                    $TaskDescription = "Patching of database servers. `n"+"`nDeployment window start date/time: `n"+"`nDeployment window end date/time:"
                    New-Tasks -AssignedTo 'Ashok Dubey' -PBITitle $PBITaskTitle -DescriptionContent $TaskDescription -Skillset "Database Administration" -ParentID $id

                    # Create tasks for QuarterlyServerPatching to Pedro
                    $TaskDescription = "Patching of application and/or web servers. `n"+"`nDeployment window start date/time: `n"+"`nDeployment window end date/time:"
                    New-Tasks -AssignedTo 'Tarak' -PBITitle $PBITaskTitle -DescriptionContent $TaskDescription -Skillset "Other" -ParentID $id
                }
                else {
                    Write-Host "Provided Server is not valid"
                }
                return
            }
            else {
                Write-Host "Provided Server is not valid"
            }
        }
    }
    else {
        Write-Host "Provide Server is not valid"
    }
    return
}
else {
    Write-Error "Invalid value for ChangeRequestType parameter. Allowed values are 'AzureDBIPWhitelisting' and 'QuarterlyServerPatching'."
}