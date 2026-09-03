param(
  [Parameter(Mandatory = $true)]
  [int]$rows,

  [Parameter(Mandatory = $true)]
  [int]$columns
)

function main {
  # Calculate name based on rows/columns ratio
  $rowsByColumns = "$($rows)x$($columns)"
  $ratio = $rows / $columns

  $name = switch ($true) {
    ($ratio -lt 1) { "Landscape $rowsByColumns" }
    ($ratio -eq 1) { "Square $rowsByColumns" }
    ($ratio -gt 1) { "Portrait $rowsByColumns" }
    default { "$rowsByColumns Grid" }
  }

  # Calculate percentages
  $rowPercentages = @(for ($i = 0; $i -lt $rows; $i++) { [int][Math]::Round(100 / $rows * 100, 0) })
  $columnPercentages = @(for ($i = 0; $i -lt $columns; $i++) { [int][Math]::Round(100 / $columns * 100, 0) })

  # Percentages must sum up to 100%, otherwise there's a parse error on import
  # So we need to calculate the actual sum after rounding and correct for it if it's not 100%
  ScaleToOneHundredPercent($rowPercentages)
  ScaleToOneHundredPercent($columnPercentages)

  # Create cell-child-map
  # The comma before @() in the for loop forces PowerShell to treat each sub-array as a separate array, preventing it from flattening the array.
  $cellChildMap = @(for ($i = 0; $i -lt $rows; $i++) { , @(for ($j = 0; $j -lt $columns; $j++) { $i * $columns + $j }) })

  # Create JSON object
  $jsonObject = @{
    uuid = (New-Guid).Guid
    name = $name
    type = "grid"
    info = @{
      rows                 = $rows
      columns              = $columns
      "rows-percentage"    = $rowPercentages
      "columns-percentage" = $columnPercentages
      "cell-child-map"     = $cellChildMap
      "show-spacing"       = $false
      spacing              = 16
      "sensitivity-radius" = 20
    }
  }

  # Convert to JSON
  $json = $jsonObject | ConvertTo-Json -Depth 5

  # Write to file
  $filename = "$name ($($jsonObject.uuid)).json"
  $json | Out-File -FilePath $filename

  "Successfully created '$(Convert-Path -LiteralPath $filename)'"
}

<#
.SYNOPSIS
Scale the given percentages array to sum up to 100, applying
the correction values to the outermost rows/columns until the
sum reaches 100% (10 000 in FancyZones values)
#>
function ScaleToOneHundredPercent {
  param (
    $percentageArray
  )

  # Calculate if our sum is too small or too big
  $correction = 10000 - ($percentageArray | Measure-Object -Sum).Sum

  # and set the correction step accordingly
  $correctionStep = $correction -gt 0 ? 1 : -1 

  $correctFromRight = $true
  $leftIndex = 0
  $rightIndex = $percentageArray.Count - 1
  while (10000 -ne ($percentageArray | Measure-Object -Sum).Sum) {
    if ($correctFromRight) {
      $percentageArray[$rightIndex] = $percentageArray[$rightIndex] + $correctionStep
      $rightIndex = $rightIndex - 1
    }
    else {
      $percentageArray[$leftIndex] = $percentageArray[$leftIndex] + $correctionStep
      $leftIndex = $leftIndex + 1
    }

    $correctFromRight = !$correctFromRight
  }
}

main

