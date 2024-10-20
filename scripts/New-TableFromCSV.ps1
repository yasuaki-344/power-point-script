function New-TableFromCSV {
    param (
        [string]$csvPath,
        [string]$pptTemplatePath,
        [string]$outputPptPath
    )

    # Load PowerPoint application
    $app = New-Object -ComObject PowerPoint.Application
    # $app.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse
    $app.WindowState = [Microsoft.Office.Interop.PowerPoint.PpWindowState]::ppWindowMinimized

    # Open the template PowerPoint file
    $presentation = $app.Presentations.Open(
        $pptTemplatePath,
        [Microsoft.Office.Core.MsoTriState]::msoFalse,
        [Microsoft.Office.Core.MsoTriState]::msoTrue,
        [Microsoft.Office.Core.MsoTriState]::msoFalse
    )

    # Read the CSV file
    $data = Import-Csv -Path $csvPath

    # Add a new slide
    $slide = $presentation.Slides.Add($presentation.Slides.Count + 1, [Microsoft.Office.Interop.PowerPoint.PpSlideLayout]::ppLayoutText)

    # Define table dimensions
    $rows = $data.Count + 1 # Include header row
    $columns = ($data[0].PSObject.Properties | Measure-Object).Count

    # Insert a table into the slide
    $table = $slide.Shapes.AddTable($rows, $columns).Table

    # Insert header row
    $headers = $data[0].PSObject.Properties | ForEach-Object { $_.Name }
    for ($col = 1; $col -le $columns; $col++) {
        $table.Cell(1, $col).Shape.TextFrame.TextRange.Text = $headers[$col - 1]
    }

    # Insert data rows
    for ($row = 2; $row -le $rows; $row++) {
        $rowData = $data[$row - 2].PSObject.Properties | ForEach-Object { $_.Value }
        for ($col = 1; $col -le $columns; $col++) {
            $table.Cell($row, $col).Shape.TextFrame.TextRange.Text = [string]$rowData[$col - 1]
        }
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
