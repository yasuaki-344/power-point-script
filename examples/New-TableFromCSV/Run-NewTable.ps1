# Import the function from the external script
. .\..\..\scripts\New-TableFromCSV.ps1

# Get the current script directory
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition

$csvPath = "$scriptDir\table.csv"
$pptTemplatePath = "$scriptDir\template.pptx"
$outputPptPath = "$scriptDir\output.pptx"

New-TableFromCSV -csvPath $csvPath -pptTemplatePath $pptTemplatePath -outputPptPath $outputPptPath
