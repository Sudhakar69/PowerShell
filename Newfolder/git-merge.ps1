param ([Parameter(Mandatory)][string]$RepoName="PLEASE PROVIDE REPOSITORY NAME AS AN ARGUMENT TO THIS SCRIPT",
        [string]$Path="PLEASE PROVIDE REPOSITORY PATH AS AN ARGUMENT TO THIS SCRIPT")

#Defining Repository path here
if($null -eq $Path)
{
    $Path= $pwd.Path
}
Set-Location $Path

$childitems = Get-childitem
foreach($childitem in $childitems){
    if($RepoName -eq $childitem)
    {
	Set-Location $childitem
	Write-Host $RepoName "Repo is found on local."
    }
    else{
	Write-Host $RepoName "Repo is not found on local. Cloning repo from remote."
	$Repouri = "https://apxinc.visualstudio.com/Apx/_git/"
	
	#Combining RepoName to the Basic Url
	$Repopath = $Repouri + $RepoName
	git clone $Repopath
	# clsSet-Location $cc.Name
	Set-Location $RepoName

    }
    
    git pull --prune --all
    if (git branch -a | Select-String -Pattern "main") {
	# Set main as the source branch
	$SourceBranch = "main"
    }
    elseif (git branch -a | Select-String -Pattern "master") {
	# Set master as the source branch
	$SourceBranch = "master"
    }
    else {
	# Exit with an appropriate message
	Write-Host "No main or master branch found." -BackgroundColor Yellow
	Exit
    }
    Write-Host "Source branch is '$SourceBranch'"

    # Checkout and pull the source branch
    git checkout $SourceBranch
    git pull --all

    # Get all the branches that are not main, master, or locked
    $Branches = git branch -a | Where-Object {$_ -notmatch "main|master|locked"}

    # Loop through each branch
    foreach ($Branch in $Branches) {
	Write-Host "---------------------------------------------------------------------------------------"
	Write-Host "Working on branch '$Branch'"

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

	    # Abort the merge
	    git merge --abort
	    Write-Host "Aborting the merge" -BackgroundColor Red
	}
	else {
	    # Commit the merge with a message
	    git commit -m "Merged branch '$SourceBranch' into '$Branch'"
	}
    }
    return
}
