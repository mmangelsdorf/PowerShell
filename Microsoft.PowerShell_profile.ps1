#region ------ Initialisation ------
# Helper to display script loading steps :)
function Run-Step([string] $Description, [ScriptBlock]$script) {
   Write-Host  -NoNewline "Loading" $Description.PadRight(20)
   & $script
   Write-Host '✅'
}

Write-Host "Loading PowerShell $($PSVersionTable.PSVersion)..." -ForegroundColor DarkCyan
Write-Host

# Import posh-git
Run-Step "posh-git" {
   Import-Module posh-git
}

# Initialize oh-my-posh
Run-Step "oh-my-posh" {
   oh-my-posh --init --shell pwsh --config ".\oh-my-posh-themes\jandedobbeleer-mod.omp.json" | Invoke-Expression
}

# Update environment variables
Run-Step "update-path" {
   $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
}

Write-Host

$ErrorActionPreference = 'Inquire'

#endregion


#region ------ Misc ------

<# .SYNOPSIS
Shortcut for getting a filename-compatible timestamp string (yyyyMMdd_HHmmss) #>
function Get-TimeStamp { Get-Date -f yyyyMMdd_HHmmss }

<# .SYNOPSIS
Get the currently running command line for the given process #>
function Get-CommandLineForProcess {
   param( [string]$processName )
   (Get-CimInstance Win32_Process -Filter "name like '%$processName%'").CommandLine
}

<# .SYNOPSIS
Try to open VS Code workspace of current folder #>
function ws {
   # Try to find workspace in current folder
   $workspaceFileToOpen = (Get-ChildItem -Path . -Filter *.code-workspace | Select-Object -First 1).FullName

   # otherwise fall-back to parent folder
   if ([string]::IsNullOrWhiteSpace($workspaceFileToOpen)) {
      'No .code-workspace file in current folder, checking parent...'
      $workspaceFileToOpen = ('..\' + ( Split-Path . -Leaf ) + '.code-workspace')
   }

   if (Test-Path $workspaceFileToOpen) {
      "Opening '$workspaceFileToOpen'"
   }
   else {
      "No '$(( Split-Path . -Leaf ) + '.code-workspace')' file found either..."
      "Opening current folder directly"
      $workspaceFileToOpen = '.'
   }

   code $workspaceFileToOpen
}

#endregion


#region ------ PowerShell ------

<# .SYNOPSIS
Open an elevated PowerShell window (see: https://serverfault.com/a/993412/304668) #>
function GoAdmin { Start-Process wt -Verb RunAs }

<# .SYNOPSIS
Shortcut to get current PS version #>
function GetVersion { $PSVersionTable.PSVersion }

<# .SYNOPSIS
Shortcut to opening the global profile file in VS Code #>
function OpenProfile { code $profile --new-window }

<# .SYNOPSIS
Shortcut to reload the profile file (unfortunately this doesn't work reliably ☹️) #>
function ReloadProfile { . $profile }

<# .SYNOPSIS
Show list of all functions defined in $PROFILE (and all other profiles) #>
function Show-AvailableProfileCommands {
   Write-Host "Available Profile Functions:" -ForegroundColor Green
   $profileData = Get-Content $profile
   $profileData += Get-Content $profile.AllUsersAllHosts -ErrorAction SilentlyContinue
   $profileData += Get-Content $profile.AllUsersCurrentHost -ErrorAction SilentlyContinue
   $profileData -match "^function" `
   | ForEach-Object { $_.split(" ")[1] } `
   | Sort-Object -Unique `
   | ForEach-Object { Write-Host "`t$_" -ForegroundColor Green }
}

<# .SYNOPSIS
Shortcut to show history of commands entered in PowerShell terminals #>
function Show-CommandLineHistory { code "%AppData%\Microsoft\Windows\PowerShell\PSReadLine\ConsoleHost_history.txt" --new-window }

#endregion


#region ------ File Utils ------

<# .SYNOPSIS
Returns a nicely formatted string of the file/folder's size (eg. '4,20 MB') #>
function Get-FormattedSize {
   param (
      [Parameter(Mandatory)]
      [ValidateNotNullOrEmpty()]
      [string] $fileOrFolderPath
   )
   if (!(Test-Path $fileOrFolderPath)) {
      return "Doesn't exist"
   }

   if ((Get-Item $fileOrFolderPath -Force) -is [System.IO.DirectoryInfo]) {
      # Path is a directory
      $byteCount = (Get-ChildItem $fileOrFolderPath -Force -Recurse | Measure-Object -Property Length -Sum).Sum
   }
   else {
      # Path is a file
      $byteCount = (Get-Item $fileOrFolderPath -Force).Length
   }

   $formattedSize = switch ([math]::truncate([math]::log($byteCount, 1024))) {
      '0' { "$byteCount Bytes" }
      '1' { "{0:n2} KB" -f ($byteCount / 1KB) }
      '2' { "{0:n2} MB" -f ($byteCount / 1MB) }
      '3' { "{0:n2} GB" -f ($byteCount / 1GB) }
      '4' { "{0:n2} TB" -f ($byteCount / 1TB) }
      default { "{0:n2} PB" -f ($byteCount / 1pb) }
   }

   return $formattedSize
}

<# .SYNOPSIS
Compares the hashes of all files in a directory to check if they're all identical #>
function Test-IfAllFilesAreIdentical {
   param (
      [Parameter(Mandatory = $true)]
      [string]$folderPath
   )

   $files = Get-ChildItem -Path $folderPath -File
   $hashes = $files | ForEach-Object { (Get-FileHash $_.FullName).Hash }
   $uniqueHashes = $hashes | Sort-Object | Get-Unique

   if ($uniqueHashes.Count -eq 1) {
      Write-Output "All files are identical"
   }
   else {
      Write-Output "Not all files are identical"
   }
}

<# .SYNOPSIS
Given a list of filenames, tests their existence and displays the check result #>
function Write-FileExistenceStatus {
   param (
      [array]$filesToDisplay
   )

   foreach ($file in $filesToDisplay) {
      $existenceFlag = if (Test-Path $file) { '✅' } else { '❌' }
      "$existenceFlag $file"
   }
}

<# .SYNOPSIS
Gets a list of files with a given filetype in the current directory, filtered by $wordPart #>
function Get-FilteredFileList {
   param([string] $extensionFilter, [string] $wordPart)
   $filesWithExtension = (Get-ChildItem $extensionFilter).Name | ForEach-Object { "'.\$_'" }
   $filter = $wordPart ? $wordPart : ".*"
   return $filesWithExtension | Select-String "$filter"
}

<# .SYNOPSIS
Convert the given file path to an absolute path.
If the given file path doesn't exist yet, a temporary file will be created to get the path.
If the parent directory doesn't exist, the file path is return unchanged. #>
function ConvertTo-AbsolutePath {
   param (
      [string]$filePath
   )

   # If the file already exists, we can use existing functionality
   if (Test-Path $filePath) {
      return (Convert-Path -LiteralPath $filePath)
   }

   # Fail if the parent folder doesn't exist to avoid side-effects like creating a folder hierarchy
   $parentDirectory = [System.IO.Path]::GetDirectoryName($filePath)
   if (!(Test-Path $parentDirectory)) {
      Write-Error "Parent directory doesn't exist! Returning unaltered file path"
      return $filePath
   }

   # If we got here, we need to create the temporary file for Convert-Path...

   # Save original extension
   $originalExtension = [System.IO.Path]::GetExtension($filePath)

   # Create dummy file path that doesn't exist
   while ((!$guid) -or (Test-Path $dummyFile)) {
      $guid = New-Guid
      $dummyFile = [System.IO.Path]::ChangeExtension($filePath, $guid)
   }

   # Create a temporary file to be able to use Convert-Path
   New-Item -Path $dummyFile -Force | Out-Null
   $absolutePath = Convert-Path -LiteralPath $dummyFile
   Remove-Item $dummyFile -Force

   # Re-apply original extension
   $absolutePath = [System.IO.Path]::ChangeExtension($absolutePath, $originalExtension)

   return $absolutePath
}

#endregion


#region ------ winget Utils ------

<# .SYNOPSIS
Shortcut for winget update --all --include-unknown #>
function wu { winget update --all --include-unknown }

<# .SYNOPSIS
Shortcut to update to latest PS version #>
function Update-PS { winget update Microsoft.PowerShell }

<# .SYNOPSIS
Shortcut to update to latest git version
Currently installing the following options:
* Git LFS (Large File Support)
* Associate .git* configuration files with the default text editor
* Associate .sh files to run with Bash
* Check daily for Git for Windows updates
* Scalar (Git add-on to manage large-scale repositories)
(explicitly _NO_ shell extensions...)

For all available options, see:
https://github.com/git-for-windows/build-extra/blob/d1a9c4e920b476c4b4d439051ecddc9e281ca8a1/installer/install.iss#L107
#>
function Update-git { winget install --id=Git.Git --source winget --force --custom '/components="gitlfs,assoc,assoc_sh,autoupdate,scalar"' }

#endregion


#region ------ JSON Manipulation ------

<# .SYNOPSIS
Format *.json files (or --fileFilter) in the current folder (or --searchDirectory) and all non-hidden subdirectories #>
function Format-AllJsonHereAndInSubdirectories {
   param( [string]$searchDirectory, [string]$fileFilter )

   # Set defaults if parameters were not passed
   if ([String]::IsNullOrWhiteSpace($searchDirectory)) { $searchDirectory = '.' }
   if ([String]::IsNullOrWhiteSpace($fileFilter)) { $fileFilter = '*.json' }

   if (!(Test-Path $searchDirectory)) {
      "Directory not found: '$searchDirectory'"
      return
   }

   "Pretty-printing $((Get-ChildItem -Path $searchDirectory -Filter $fileFilter -Recurse).Length) files...`n"

   # Pretty print all json files
   $i = 1
   Get-ChildItem -Path $searchDirectory -Filter $fileFilter -Recurse | ForEach-Object {
      Write-Host "`r$i" -NoNewline
      $content = [System.IO.File]::ReadAllText($_.FullName) | ConvertFrom-Json | ConvertTo-Json -Depth 100
      [System.IO.File]::WriteAllText($_.FullName, $content)
      $i = $i + 1
   }
   Write-Host "`r" -NoNewline # Remove last displayed number

   'Done ✅'
   ''
   '⚠ Note that the parsing and re-writing might change some values, these are the changes to expect (at least):'
   '  * JSON files containing an array that has only one element will lose their array brackets'
   '    (this is automatically fixed for all content.json files but might affect other files in non-TSX contexts)'
   '  * "\/" will be replaced by "/"'
   '  * Numbers in scientific notation might be reformatted (e.g. "7.52E-8" -> "7.52E-08", "6.2414E24" -> "6.2414E+24")'
   '  * DateTime values will lose the milliseconds  (e.g. "2023-03-01T00:00:00.000+01:00" -> "2023-03-01T00:00:00+01:00")'
   '  * Unicode code-points might be replaced with their symbol (e.g. "\u00F6" -> "ö", "\u00B3" -> "³")'
   '  * The output file will definitely be UTF-8 encoded'
   '  * Hidden files/folders are excluded by default'
}

<#
.SYNOPSIS
  Sort all attributes in *.json files (or --fileFilter) in the current folder (or --searchDirectory) and all non-hidden subdirectories.
  Requires jq to be installed for json manipulation (see: https://jqlang.org/).
#>
function Sort-JsonAttributesHereAndInSubdirectories {
   param( [string]$searchDirectory, [string]$fileFilter )

   # Set defaults if parameters were not passed
   if ([String]::IsNullOrWhiteSpace($searchDirectory)) { $searchDirectory = '.' }
   if ([String]::IsNullOrWhiteSpace($fileFilter)) { $fileFilter = '*.json' }

   if (!(Test-Path $searchDirectory)) {
      "Directory not found: '$searchDirectory'"
      return
   }

   "Sorting attributes in $((Get-ChildItem -Path $searchDirectory -Filter $fileFilter -Recurse).Length) files...`n"

   # Sort attributes in JSON files (using JQ)
   $i = 1
   Get-ChildItem -Path $searchDirectory -Filter $fileFilter -Recurse | ForEach-Object {
      Write-Host "`r$i" -NoNewline
      [System.IO.File]::WriteAllLines($_.FullName, (& jq -S . $_.FullName))
      $i = $i + 1
   }
   Write-Host "`r" -NoNewline # Remove last displayed number
}

#endregion


#region ------ XML Manipulation ------

<# .SYNOPSIS
Format *.xml files (or --fileFilter) in the current folder (or --searchDirectory) and all non-hidden subdirectories #>
function Format-AllXmlHereAndInSubdirectories {
   param( [string]$searchDirectory, [string]$fileFilter )

   # Set defaults if parameters were not passed
   if ([String]::IsNullOrWhiteSpace($searchDirectory)) { $searchDirectory = '.' }
   if ([String]::IsNullOrWhiteSpace($fileFilter)) { $fileFilter = '*.xml' }

   if (!(Test-Path $searchDirectory)) {
      "Directory not found: '$searchDirectory'"
      return
   }

   "Pretty-printing $((Get-ChildItem -Path $searchDirectory -Filter $fileFilter -Recurse).Length) files...`n"

   # Pretty print all XML files
   Get-ChildItem -Path $searchDirectory -Filter $fileFilter -Recurse | ForEach-Object {
      Write-Host "`r$i" -NoNewline
      $xmlData = [xml]::new()
      $xmlData.PreserveWhitespace = $false
      $xmlData.Load($_.FullName)
      $xmlData.Save($_.FullName)
      $i = $i + 1
   }
   Write-Host "`r" -NoNewline # Remove last displayed number
}

<# .SYNOPSIS
Saves an XML document at the given $fileName using the given $indentation #>
function Save-XmlWithIndentation ($xml, $fileName, $indentation) {
   $settings = New-Object System.Xml.XmlWriterSettings
   $settings.Indent = $true
   $settings.IndentChars = ' ' * $indentation
   $writer = [System.Xml.XmlWriter]::Create($fileName, $settings)
   $xml.Save($writer)
   $writer.Close()
}

<# .SYNOPSIS
Get the fully-defined XPath for the given XmlNode #>
function Get-XPath {
   param(
      [Parameter(Mandatory)]
      [ValidateNotNullOrEmpty()]
      [System.Xml.XmlNode] $node
   )

   if ($node.NodeType -eq 'Document') {
      return ''
   }

   # Calculate index amongst siblings
   $index = 1
   $sibling = $node.PreviousSibling
   while ($null -ne $sibling) {
      if ($sibling.Name -eq $node.Name) {
         $index++
      }
      $sibling = $sibling.PreviousSibling
   }

   # Recursively build path to root
   $parentPath = Get-XPath $node.ParentNode
   return "$parentPath/$($node.Name)[$index]"
}

#endregion


#region ------ git Utils ------

<# .SYNOPSIS
Get git status #>
function gs { git status }

<# .SYNOPSIS
Return to develop branch #>
function gd { git-switch develop }

<# .SYNOPSIS
Start smartgit #>
function smartgit { & 'C:\Program Files\SmartGit\bin\smartgit.exe' . }
Set-Alias -Name sg -Value smartgit -Description "Start smartgit here"

<# .SYNOPSIS
List release branches available remotely #>
function git-branch-list-releases { "`nCurrent release branches:"; git branch --list -r origin/release/*; '' }

<# .SYNOPSIS
Open TortoiseGit commit view #>
function git-commit-tortoise-git { TortoiseGitProc.exe /command:commit | Out-Null } # Piping will wait for the command to finish

<# .SYNOPSIS
Show TortoiseGit repository status #>
function git-diff-tortoise-git { TortoiseGitProc.exe /command:repostatus }

<# .SYNOPSIS
Create branch with name suggestions from Azure DevOps #>
function git-create-branch {
   param(
      [Parameter(Mandatory)]
      [ValidateNotNullOrEmpty()]
      [ArgumentCompleter( {
            param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)
            # Get all work items from the current sprint that the current user has ever been assigned to, ordered by ID
            # ! Make sure to replace set your actual Azure DevOps team information here for the query to work
            # (To find the team info and ID, create a query using the "Iteration Path = @currentIteration" filter and then look at the WIQL editor)
            $azureTeam = '[AZD PROJECT]\My Team <id:01fa9802-5814-4f20-9fe3-fa5e98df0d02>'
            $wiql = "SELECT [System.Title] FROM workitems WHERE EVER [System.AssignedTo] = @me AND ( [System.IterationPath] = @currentIteration('$azureTeam') OR [System.IterationPath] = @currentIteration('$azureTeam') + 1 ) ORDER BY [System.Id]"
            $workItemsJson = az boards query --wiql $wiql

            # If query failed, we probably need to login
            if (-not $?) { az login }

            # Create list of branch name suggestions based on work item titles
            $branchSuggestions = $workItemsJson
            | ConvertFrom-Json
            | Select-Object -Property id -ExpandProperty fields
            | ForEach-Object {
               $workitemID = $_.id
               $workitemTitle = ConvertTo-ValidBranchName( $_.'System.Title' )
               "feature/$workitemID-$workitemTitle"
            }

            $filter = $wordToComplete ? $wordToComplete : ".*"
            return $branchSuggestions | Select-String "$filter"
         })]
      [string] $newBranchName )

   git switch -c $newBranchName
}

<# .SYNOPSIS
Sanitize branch name candidate by replacing spaces with underscores and removing any non-word characters except '-' #>
function ConvertTo-ValidBranchName {
   param (
      [Parameter(Mandatory)]
      [ValidateNotNullOrEmpty()]
      [string] $unclean
   )
   return $unclean -replace (' ', '_') -replace ('[^\w-]', '')
}

<# .SYNOPSIS
Delete unused local branches #>
function git-delete-unused-branches { git branch | ForEach-Object { git branch -d $_.Trim() } }

<# .SYNOPSIS
git pull (used to also do submodule update) #>
function git-pull-and-update-submodules {
   Write-Host "`nPulling updates..." -ForegroundColor DarkCyan
   git pull
   Write-Host "`nDone! :)`n" -ForegroundColor Green

   # "Skipping submodules update because we're lazy..."
   # git submodule update --init --recursive
}
Set-Alias -Name gu -Value git-pull-and-update-submodules -Description "Pull and update submodules" -Force
# Set-Alias -Name gp -Value git-pull-and-update-submodules -Description "Pull and update submodules" -Force
# (This alias would override gp for Get-ItemProperty that is defined by default in PowerShell)

<# .SYNOPSIS
Stash changes - git update - pop stash #>
function gus {
   Write-Host "`nStashing current changes..." -ForegroundColor DarkMagenta
   git stash
   git-pull-and-update-submodules
   Write-Host "Re-applying stashed changes..." -ForegroundColor DarkGreen
   git stash pop
}

# Shortcut to git reset --hard because I used it so often at some point :P
function git-reset-hard { git reset --hard }

<# .SYNOPSIS
Do git checkout with branch name autocompletion and automatic pull after checkout (Inspiration: https://stackoverflow.com/a/58844032/2822719) #>
function git-switch {
   param(
      [Parameter(Mandatory)]
      [ValidateNotNullOrEmpty()]
      [ArgumentCompleter( {
            param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)
            $branchList = (git branch --all --format='%(refname:short)') -replace 'origin/', ''
            $filter = $wordToComplete ? $wordToComplete : ".*"
            return $branchList | Select-String "$filter"
         })]
      [string] $branch
   )

   "`nSwitching to $branch..."
   git fetch

   git switch $branch

   git-pull-and-update-submodules
}

#endregion


#region ------ Silly Stuff ------

<# .SYNOPSIS
Shortcut to execute a file once it exists (checked every 500ms) #>
function Start-WhenExists {
   param(
      [Parameter(Mandatory)]
      [ValidateNotNullOrEmpty()]
      [ArgumentCompletions('.\Tests.exe', '.\TestInsightProject.exe',
         '.\bin\Win32\Debug\Tests.exe', '.\bin\Win32\Release\Tests.exe',
         '.\bin\Win64\Debug\Tests.exe', '.\bin\Win64\Release\Tests.exe',
         '.\bin\Win32\Debug\TestInsightProject.exe', '.\bin\Win32\Release\TestInsightProject.exe',
         '.\bin\Win64\Debug\TestInsightProject.exe', '.\bin\Win64\Release\TestInsightProject.exe'
      )]
      [string] $fileToExecute )

   # Try to get one of the spinners from the web
   try {
      [string[]]$favorites = "dots12", "simpleDotsScrolling", "bouncingBall", "moon", "pong", "fistBump", "orangeBluePulse", "aesthetic", "hearts", "clock"

      $json = (Invoke-WebRequest -Uri "https://raw.githubusercontent.com/sindresorhus/cli-spinners/main/spinners.json").Content | ConvertFrom-Json
      $randomSpinner = Get-Random $favorites
      # Uncomment the following line if you want to use a random of all the available spinners
      # $randomSpinner = $json | Get-Member -MemberType Properties | Select-Object -ExpandProperty Name | Get-Random
      $spin = $json.$randomSpinner.frames
   }
   catch {
      # Otherwise fall back to our manual ones
      [string[]]$spin1 = "🕛", "🕐", "🕑", "🕒", "🕓", "🕔", "🕕", "🕖", "🕗", "🕘", "🕙", "🕚"
      [string[]]$spin2 = "🔴", "🟠", "🟡", "🟢", "🔵", "🟣", "🟪", "🟦", "🟩", "🟨", "🟧", "🟥"
      [string[]]$spin3 = "🤜        🤛", "🤜        🤛", "🤜        🤛", " 🤜      🤛 ", "  🤜    🤛  ", "   🤜  🤛   ", "    🤜🤛    ", "   🤜✨🤛   ", "  🤜 ✨ 🤛  ", " 🤜      🤛 "

      $spin = Get-Random ($spin1, $spin2, $spin3)
   }

   while (!(Test-Path $fileToExecute)) {
      Write-Host "`r$($spin[$i++%$spin.Length]) Waiting for '$fileToExecute'" -NoNewline
      # Write-Progress -Activity "$($spin[$i++%$spin.Length]) Waiting for $fileToExecute" -PercentComplete ($i * 10 % 100)
      Start-Sleep 0.5 # seconds
   }
   Write-Host "`r '$fileToExecute' now exists! 🎉`nStarting it right away..."
   Start-Sleep 5 # Wait five more seconds just to make time for additional files to be copied (in the case of waiting for a compilation result)
   & $fileToExecute
}

<# .SYNOPSIS
Start a countdown timer on timeanddate.com, e.g. to share during a break in a long-running meeting #>
function Start-CountdownTimer {
   param(
      [string] $delayOrTime,
      [string] $subject
   )
   # Default to a 10-minute timer and the subject 'Meeting continues in...'
   [System.DateTime]$targetTime = (Get-Date).AddMinutes(10)
   if ([string]::IsNullOrWhiteSpace($subject)) {
      $subject = 'Meeting continues in...'
   }

   # Interpret everything with max. 2 digits as a delay
   if ($delayOrTime.Length -le 3) {
      $targetTime = (Get-Date).AddMinutes($delayOrTime)
   }
   else {
      $delayOrTime = $delayOrTime.Replace(':', '')    # Allow 19:30 as well as 1930
      $targetTime = (Get-Date).Date + (New-TimeSpan -Hours $delayOrTime.Substring(0, 2) -Minutes $delayOrTime.Substring(2, 2))
   }

   "delayOrTime: $delayOrTime"
   "targetTime:  $targetTime"
   "subject:     $subject"

   $subject = $subject.Replace(' ', '+') # Should actually be URL encoded in the future...
   $timeString = $targetTime.ToString('yyyyMMddTHHmmss')
   $url = "https://www.timeanddate.com/countdown/launch?iso=${timeString}&p0=317&msg=${subject}&font=slab&csz=1#"

   'URL: ' + $url
   Start-Process $url
}

#endregion
