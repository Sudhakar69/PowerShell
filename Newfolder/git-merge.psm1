function DoGitmerge {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [string] $Repo,
        [Parameter(Mandatory = $true)]
        [string] $Path
    )

    # Change to the specified directory
    Set-Location $Path

    # Get the current branch name using git symbolic-ref
    $CurrentBranch = git symbolic-ref --short HEAD

    if ($CurrentBranch -eq $Repo) {
        # Pull changes from the remote repo
        git pull
    }
    else {
        # Check if Git has the specified remote repo
        $GitRepos = git remote -v | Select-String -Pattern $Repo

        if ($GitRepos) {
            # Clone the repo to the current folder
            git clone $GitRepos[0].Line.Split()[1]
        }
        else {
            # Exit with an error message
            Write-Error "Repo '$Repo' does not exist."
            Exit
        }
    }

    # Determine the source branch for merging
    $SourceBranch = "main"
    if (!(git branch -a --list "main" -eq $null)) {
        $SourceBranch = "main"
    }
    elseif (!(git branch -a --list "master" -eq $null)) {
        $SourceBranch = "master"
    }
    else {
        # Exit with a message if no main or master branch is found
        Write-Output "No main or master branch found."
        Exit
    }

    # Checkout and pull the source branch
    git checkout $SourceBranch
    git pull

    # Get all the branches that are not main, master, or locked
    $Branches = git branch -a | Where-Object { $_ -notmatch "main|master|locked" }

    foreach ($Branch in $Branches) {
        # Trim and checkout the branch
        $BranchName = $Branch.Trim()
        git checkout $BranchName
        git pull

        # Merge from the source branch
        git merge $SourceBranch

        # Check for conflicts
        if (git ls-files -u) {
            # Display merge conflict details
            Write-Output "Merge conflict in branch $BranchName"
            git diff --name-only --diff-filter=U

            # Abort the merge
            git merge --abort
        }
        else {
            # Commit the merge with a message
            git commit -m "Merged branch '$SourceBranch' into '$BranchName'"
        }
    }
}

# Example usage
# Do-Gitmerge -Repo "myRepo" -Path "C:\Path\To\Git\Repo"
