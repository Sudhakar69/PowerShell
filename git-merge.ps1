<#

Git merge script for Azure DevOps

This script will merge all branches in a repository. In case if merge is not happend due to conflict it will abort the merge and give the info.
IMPORTANT: This script will not push changes to the remote. The push needs to be performed manually as a subasequent, manual step


Example of how to invoke this script: ".\git-merge.ps1 c:\vsts\billing\"

#>
param (
    [Parameter(Mandatory = $true)]
    [string]$gitRepoPath = "Plase provide full path for the location of the repo. For example 'c:\vsts\billing\'"
)

$originalLocation = Get-Location

$repoName = ""
try {
    if (!(Test-Path $gitRepoPath)) {
         $answer = Read-Host "Path '$gitRepoPath' does not exist. Do you want to clone the repo [Y/N]?"
         if ($answer -eq 'Y') {
             $base = Split-Path -Path $gitRepoPath
             if (!(Test-Path -Path $base -PathType Container)) {
                 Write-Host "Location '$base' does not exist or is not a folder/directory. Exiting."
                 Exit
             }
             else {
                 $repoName = Split-Path -Path $gitRepoPath -Leaf
                 $repoUri = "https://$($Organization).visualstudio.com/$($Project)/_git/"
                 $repoURL = $repoUri + $repoName

                 Write-Host "About to clone repo '$repoName' ($repoURL) into '$base' ..."
                 Set-Location $base
                 git clone $repoURL
                 Set-Location $gitRepoPath
             }
         }
         else {
             Write-Host "Exiting"
             Exit
         }
    }
    else {
	if (Test-Path -Path $gitRepoPath -PathType Container) {
            Set-Location $gitRepoPath
	    $repoName = Split-Path -Path $gitRepoPath -Leaf
            Write-Host "'$repoName' repo found on local. No need to clone." -ForegroundColor Green
	}
	else {
	    Write-Host "Specified git repo path is not a folder/directory. Exiting."
	    Exit
	}
    }

    $DoPush = $false
    Write-Host "NOTE: This script will perform merges and commit them to the local repo. The script WILL NOT push the merges to the remote." -BackgroundColor Green

    Write-Host "-------------------------------------------------------------------------------"

    # Remove unneeded remote branches (note, these have been deleted through Azure DevOps UI but still exist on the remote)
    Write-Host "Deleting unneeded remote branches"
    git pull --prune --all
    Write-Host "-------------------------------------------------------------------------------"

    # Remove local branches that do not exist on remote
    Write-Host "Deleting unneeded local branches"
    git branch -vv | Select-String -Pattern ": gone]" | % { $_.toString().Trim().Split(" ")[0]} | Select-String -notmatch "\*" | % {git branch -D $_}
    Write-Host "-------------------------------------------------------------------------------"

    # Determine the source branch (main or master)
    Write-Host "Determining source branch ('main' or 'master')"
    $SourceBranch = "main"
    if (!(git branch -a | Select-String -Pattern "main")) {
        if (git branch -a | Select-String -Pattern "master") {
            $SourceBranch = "master"
        } else {
            Write-Host "No main or master branch found." -BackgroundColor Yellow
            Exit
        }
    }
    Write-Host "Source branch is '$SourceBranch'"
    Write-Host "-------------------------------------------------------------------------------"

    # Checkout and pull the source branch
    Write-Host "Checking out and pulling '$SourceBranch' branch"
    git checkout $SourceBranch
    git pull
    Write-Host "-------------------------------------------------------------------------------"

    # Get all the branches that are not main, master, or locked
    Write-Host "Getting listing of branches to be merged"
    #$Branches = git branch -a | Where-Object {$_ -notmatch "main|master|locked"} | ForEach-Object { $_.Trim() } | ForEach-Object { $_.Substring($_.lastIndexOf("/") + 1) } 
    $Branches = git branch -a | ForEach-Object { $_.Trim() } | ForEach-Object { $_.Substring($_.lastIndexOf("/") + 1).replace("* ", "") } | Where-Object { !(@("main", "master", "locked") -contains $_) }
    Write-Host "Branches to be merged:"
    $Branches | ForEach-Object { Write-Host $_ }
    Write-Host "-------------------------------------------------------------------------------"

    $conflictBranches = @()
    $branchesProcessed = @()
    $branchesMerged = @()

    # Loop through each branch
    foreach ($Branch in $Branches) {
        # the check below is needed because a branch name could be duplicated, since "git branch -a" returns both local and remote branches
        if (!$branchesProcessed.Contains($Branch) -and !($repoName -eq "marketsuite" -and $branch -eq "V64.0")) {
        Write-Host "-------------------------------------------------------------------------------"
        Write-Host "Working on branch '$Branch'" -ForegroundColor Green

        # Checkout and pull the branch
        git checkout $Branch
        git pull

        # Attempt to merge from the source branch
        git merge --no-commit $SourceBranch

        # Check if there are any conflicts
        if ((git diff --name-only --diff-filter=U).Length -gt 0) {
            # Write the details of the conflicts to the screen
            Write-Host "Merge conflict in branch '$Branch'" -BackgroundColor Red
            git diff --name-only --diff-filter=U
            $conflictBranches += $Branch
            # Abort the merge
            git merge --abort
            Write-Host "Aborting the merge" -BackgroundColor Red
        } 
        else {
            # Commit the merge with a message
            Write-Host "About to execute git commit call: git commit -m Merged branch '$SourceBranch' into '$Branch' --no-edit"
            git commit -m "Merged branch '$SourceBranch' into '$Branch'" --no-edit
            if ($DoPush) {
               git push
            } else {
               Write-Host "Commit performed but was not pushed to the remote" -ForegroundColor Yellow
            }

            $branchesMerged += $Branch
        }
        }

        $branchesProcessed += $Branch
    }
    
    Write-Host "-------------------------------------------------------------------------------"
    Write-Host "Merge performed on $($branchesMerged.Length) branch(es)." -ForegroundColor Green
    $branchesMerged | ForEach-Object { Write-Host $_ -ForegroundColor Green }
    Write-Host "-------------------------------------------------------------------------------"

    if ($branchesMerged.Length -gt 0) {
        Write-Host "To push changes to the remote, please run 'git push --all' command." -BackgroundColor Green
    }

    # report branches for which merging was not done due to conflicts
    if ($conflictBranches.Length -gt 0) {
        Write-Host "*******************************************************************************"
        Write-Host "Branches NOT merged:" -ForegroundColor Red
        foreach ($conflictBranch in $conflictBranches) {
            Write-Host "Aborted merging for '$conflictBranch' branch, due to conflict" -ForegroundColor Red
        }        
        Write-Host "*******************************************************************************"
    }
    
    $answer = Read-Host "Do you want to stay in the folder/directory of the PowerShell script or navigate to the folder/directory of the repo [Script/Repo]?"
    If ($answer -eq "Repo") {
    	$originalLocation = Get-Location
    }
}
finally {
   Set-Location $originalLocation
}