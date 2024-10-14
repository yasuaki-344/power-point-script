# Import the function from the external script
. .\..\..\scripts\Import-TitlesAndMessagesFromCSV.ps1

# Get the current script directory
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition

# Define the paths for the PowerPoint and CSV files
$csvPath = "$scriptDir\input.csv"
$pptTemplatePath = "$scriptDir\template.pptx"
$customLayoutName = "title-and-key-message"
$outputPptPath = "$scriptDir\output.pptx"

# Example usage
Import-TitlesAndMessagesFromCSV -csvPath $csvPath -pptTemplatePath $pptTemplatePath -customLayoutName $customLayoutName -outputPptPath $outputPptPath
