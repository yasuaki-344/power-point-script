# PowerShell script to insert titles and key messages to the specified PowerPoint presentation

function Insert-TitlesAndMessagesIntoPPT {
    param (
        [string]$csvPath,
        [string]$pptTemplatePath,
        [string]$customLayoutName,
        [string]$outputPptPath
    )

    # Load PowerPoint application
    $app = New-Object -ComObject PowerPoint.Application
    # $app.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse
    $app.WindowState = [Microsoft.Office.Interop.PowerPoint.PpWindowState]::ppWindowMinimized

    # Open the template PowerPoint file
    $presentation = $app.Presentations.Open($pptTemplatePath,
        [Microsoft.Office.Core.MsoTriState]::msoFalse,
        [Microsoft.Office.Core.MsoTriState]::msoTrue,
        [Microsoft.Office.Core.MsoTriState]::msoFalse
    )

    # Get the slide master and custom layout
    $slideMaster = $presentation.Designs[1].SlideMaster
    $customLayout = $slideMaster.CustomLayouts | Where-Object { $_.Name -eq $customLayoutName }
    if (-not $customLayout) {
        throw "Custom layout '$customLayoutName' not found."
    }

    # Read the CSV file
    $data = Import-Csv -Path $csvPath

    # Loop through each row in the CSV file
    foreach ($row in $data) {
        # Add a new slide with specified custom layout
        $slide = $presentation.Slides.AddSlide($presentation.Slides.Count + 1, $customLayout)
    }

    # Save the new presentation
    $presentation.SaveAs($outputPptPath)

    # Close the presentation and release COM objects
    $presentation.Close()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($presentation) | Out-Null
    $app.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

# Example usage
Insert-TitlesAndMessagesIntoPPT -csvPath "D:\dev\power-point-script\examples\Insert-ContentToPowerPoint\input.csv" -pptTemplatePath "D:\dev\power-point-script\examples\Insert-ContentToPowerPoint\template.pptx" -customLayoutName "title-and-key-message" -outputPptPath "D:\dev\power-point-script\examples\Insert-ContentToPowerPoint\result.pptx"
