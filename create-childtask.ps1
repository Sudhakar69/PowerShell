param (
    [Parameter(Mandatory=$true)]
    [string]$NewSprintNumber
)
$PATId = "Personal Access token id"

$headers = @{
    Authorization = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($PATId)"))
}
$sprintnumber = "$($Project)\\"+$NewSprintNumber

function New-PBI {
    param (
        [Parameter(Mandatory=$true)]
        [string]$sprintnumber,
        [Parameter(Mandatory=$true)]
        [string]$AssignedTo,
        [Parameter(Mandatory=$true)]
        [string]$PBIDescription,
        [Parameter(Mandatory=$true)]
        [string]$Parentid
    )
    $parenturi="https://$($Organization).visualstudio.com/DefaultCollection/$($Project)/_apis/wit/workItems/$($Parentid)"
    # Create Child Tasks for "Senthil","Ashok","Ram"
    # Example function to create a task
    $PBI = "
    [
        {
            ""op"": ""add"",
            ""path"": ""/fields/System.Title"",
            ""from"": null,
            ""value"": ""$PBIDescription""
        }
        ,
        {
            ""op"": ""add"",
            ""path"": ""/fields/System.Description"",
            ""from"": null,
            ""value"": ""$PBIDescription""
        }
        ,
        {
            ""op"": ""add"",
            ""path"": ""/fields/Microsoft.VSTS.Common.AcceptanceCriteria"",
            ""from"": null,
            ""value"": ""$PBIDescription""
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
            ""path"": ""/fields/System.IterationPath"",
            ""from"": null,
            ""value"":""$sprintnumber"",
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
                ""url"": ""$parenturi""
            },
        } 
    ]
"

    #$body = $task | ConvertTo-Json
    $body = $PBI
    #Write-Host "Body: $body"
    $PBIuri = [uri]::EscapeUriString("https://dev.azure.com/$($Organization)/$($Project)/_apis/wit/workitems/`$Product Backlog Item?api-version=7.-preview.3")
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
    
    
}


function New-Task {
    param (
        [Parameter(Mandatory=$true)]
        [string]$sprintnumber,
        [Parameter(Mandatory=$true)]
        [string]$AssignedTo,
        [Parameter(Mandatory=$true)]
        [string]$TaskDescription,
        [Parameter(Mandatory=$true)]
        [string]$Parentid
    )
    $parenturi="https://$($Organization).visualstudio.com/DefaultCollection/$($Project)/_apis/wit/workItems/$($Parentid)"
    $Task = "
    [
        {
            ""op"": ""add"",
            ""path"": ""/fields/System.Title"",
            ""from"": null,
            ""value"": ""$TaskDescription""
        }
        ,
        {
            ""op"": ""add"",
            ""path"": ""/fields/System.Description"",
            ""from"": null,
            ""value"": ""$TaskDescription""
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
            ""value"": ""To Do""
        }
        ,
        {
            ""op"": ""add"",
            ""path"": ""/relations/-"",
            ""value"": {
                ""rel"": ""System.LinkTypes.Hierarchy-Reverse"",
                ""url"": ""$parenturi""
            },
        },
        {
            ""op"": ""add"",
            ""path"": ""/fields/System.Iterationpath"",
            ""value"": ""$sprintnumber""
        }
        ,
        {
            ""op"": ""add"",
            ""path"": ""/fields/System.AssignedTo"",
            ""value"": ""$AssignedTo""
        }
    ]
"

    #$body = $task | ConvertTo-Json
    $taskbody = $task
    #Construct the Task URI
    $taskuri = [uri]::EscapeUriString("https://dev.azure.com/$($Organization)/$($Project)/_apis/wit/workitems/`$task?api-version=7.-preview.3")
    #creating new child task
    #Write-Host "body: $body"
    try {
        #Invoke a REST API call for each task to be created
        Write-Host "URI: $taskuri"
        $taskresult = Invoke-RestMethod -Uri $taskuri -Method Post -ContentType "application/json-patch+json" -Headers $headers -Body $taskbody
        $taskresult
    }
    catch {
        $_.Exception.Message
    }
    
}
######
#PBI Description for the Child Task
$DBTeam = @("Senthil ","Ashok ","Ram ")
$DBPBIDescription = "Under-30-minute general database administration-related tasks - $NewSprintNumber"
$PBIforDB = New-PBI -AssignedTo "Tarak" -sprintnumber $sprintnumber -PBIDescription $DBPBIDescription -Parentid "12345"
# Create PBI (Task) for Senthil ,Ashok ,Ram 
foreach ($DBteamMember in $DBTeam ) {
    # $DBPBIDescription = "Under-30-minute general database administration-related tasks - $NewSprintNumber"
    New-Task -AssignedTo $DBteamMember -sprintnumber $sprintnumber -TaskDescription $DBPBIDescription -Parentid $PBIforDB.Id
}

# Create PBI (Task) for Zafar
$PBITitleforZafar ="Under-30-minute Fivetran, dbt, and Snowflake support and maintenance tasks - Sprint $NewSprintNumber"

New-Task -AssignedTo "Zafar Khan" -sprintnumber $sprintnumber -TaskDescription $PBITitleforZafar -Parentid $PBIforDB.Id

# Create Child Task for Zafar

#Create PBI for Ashok
$PBITitleforAshok= "SQL Server Support for EMA -  $NewSprintNumber"

$AlexPBIforAshok = New-PBI -AssignedTo "Tarak" -sprintnumber $sprintnumber -PBIDescription $PBITitleforAshok -Parentid "12346"
# Create Child Task for Ashok

New-Task -AssignedTo "Ashok Dubey" -sprintnumber $sprintnumber -TaskDescription $PBITitleforAshok -Parentid $AlexPBIforAshok.Id
