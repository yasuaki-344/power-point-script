# Import the function from the external script
. .\..\..\scripts\Export-PowerPointTitlesAndMessages.ps1

# Get the current script directory
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition

# Define the paths for the PowerPoint and CSV files
# $pptPath = ".\presentation.pptx"
$pptPath = "$scriptDir\presentation.pptx"
$csvPath = "$scriptDir\output.csv"

# Call the function with the specified paths
Export-PowerPointTitlesAndMessages -pptPath $pptPath -csvPath $csvPath
