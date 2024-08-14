

<#

Project Reporting script for VSTS

This script sends emails to all developers, testers, and product managers notifying them
of pending branches for merging and send a gentle reminder to the Application owner and respective leads to say hey these are Ahead and Behind Count for these particular branches



#>

param (
    [Parameter(Mandatory)]
    [string]$RepoName,
    [string]$Path
)

# Set the default path to the current directory if not specified
if ($null -eq $Path) {
    $Path = $PWD.Path
    Set-Location $Path
} else {
    Set-Location $Path
}

# Define the repository URL
$RepoUri = "https://$($Organization).visualstudio.com/$($Project)/_git/"

# Check if the repository exists locally
if (Test-Path -Path $RepoName) {
    Set-Location $RepoName
    Write-Host "$RepoName Repo is found on local." -BackgroundColor Green
} else {
    Write-Host "$RepoName Repo is not found on local. Cloning repo from remote." -ForegroundColor Yellow
    $RepoPath = $RepoUri + $RepoName
    git clone $RepoPath
    Set-Location $RepoName
}

# Update the local repository with the latest changes
git pull --prune --all

# Determine the source branch (main or master)
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

# Checkout and pull the source branch
git checkout $SourceBranch
git pull --all

# Get all the branches that are not main, master, or locked
$Branches = git branch -a | Where-Object {$_ -notmatch "main|master|locked"}
$conflictBranches = @()

# Loop through each branch
foreach ($Branch in $Branches) {
    Write-Host "---------------------------------------------------------------------------------------"
    Write-Host "Working on branch '$Branch'" -ForegroundColor Green

    # Checkout and pull the branch
    git checkout $Branch.Trim()
    git pull

    # Attempt to merge from the source branch
    git merge $SourceBranch

    # Check if there are any conflicts
    if (git ls-files -u) {
        # Write the details of the conflicts to the screen
        Write-Host "Merge conflict in branch $Branch" -BackgroundColor Red
        git diff --name-only --diff-filter=U
        $conflictBranches += $Branch
        # Abort the merge
        git merge --abort
        Write-Host "Aborting the merge" -BackgroundColor Red
    } 
    else {
        # Commit the merge with a message
        git commit -m "Merged branch '$SourceBranch' into '$Branch'" -ForegroundColor Yellow
    }
}

# Report branches with conflicts
foreach ($conflictBranch in $conflictBranches) {
    Write-Host "Aborted merging for $conflictBranch Branch, due to conflict" -ForegroundColor Red
}
