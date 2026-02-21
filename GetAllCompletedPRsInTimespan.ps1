# Define the parameters
$organization = "my-org"
$project = "my-project"
$repositoryName = "my-repo"
$startDate = "2025-07-16"
$endDate = "2025-08-31"
$personalAccessToken = "<my-PAT>"   # PAT with Code-Read entitlement (create one in AZD user settings)

# Base64 encode the PAT
$base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($personalAccessToken)"))

# Paginate through all completed PRs until we go past the start date
$pullRequestsInRange = @()
$top = 100
$skip = 0
$keepGoing = $true

while ($keepGoing) {
  $uri = "https://dev.azure.com/$organization/$project/_apis/git/repositories/$repositoryName/pullrequests" +
  "?searchCriteria.status=completed" +
  "&`$top=$top" +
  "&`$skip=$skip" +
  "&api-version=7.0"

  $response = Invoke-RestMethod -Uri $uri -Method Get -Headers @{Authorization = ("Basic {0}" -f $base64AuthInfo) }
  if ($response.value.Count -eq 0) {
    break
  }

  foreach ($pr in $response.value) {
    $closed = [DateTime]$pr.closedDate

    # PR closed after our window — skip this PR but keep paginating (results are newest-first)
    if ($closed.Date -gt $endDate.Date) { 
      continue 
    }

    # PR closed before our window — everything from here on is too old, stop paginating
    if ($closed.Date -lt $startDate.Date) {
      $keepGoing = $false
      break
    }

    # PR is within range
    $pullRequestsInRange += $pr
  }

  $skip += $top

  # If a page returned fewer results than requested, we've hit the end and need to quit
  if ($response.value.Count -lt $top) {
    break
  }
}

# Keep only the interesting properties
$formattedPRs = $pullRequestsInRange | Select-Object `
@{Name = "PR ID"; Expression = { $_.pullRequestId } },
@{Name = "Target Branch"; Expression = { $_.targetRefName } },
@{Name = "Completion Date"; Expression = { $_.closedDate } },
@{Name = "PR Title"; Expression = { $_.title } },
@{Name = "Link"; Expression = { "https://dev.azure.com/$organization/$project/_git/$repositoryName/pullrequest/$($_.pullRequestId)" } }

# Use your desired output
$formattedPRs | Format-Table -AutoSize

$formattedPRs | ConvertTo-Csv

Write-Host "Total PRs found: $($formattedPRs.Count)"