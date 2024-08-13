$Path = Read-Host "Please enter Repository path"
if($null -eq $Path)
{
    $Path= $pwd.Path
}
Set-Location $Path
$childitems = Get-childitem
$RepoName = "billing"
foreach($childitem in $childitems){
    if($RepoName -eq $childitem)
    {
	    Set-Location $childitem
        Write-Host $RepoName "repo is available"
        # return($RepoName)
    }
    else{
        Write-Host $RepoName "repo is not available"
        $Repouri = "https://apxinc.visualstudio.com/Apx/_git/"
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
    
    # Checkout and pull the source branch
    git checkout $SourceBranch
    git pull
    
    # Get all the branches that are not main, master, or locked
    $Branches = git branch -a | Where-Object {$_ -notmatch "main|master|locked"}
    
    # Loop through each branch
    foreach ($Branch in $Branches) {
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
