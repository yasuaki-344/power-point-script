# PowerShell script to extract titles and key messages from a PowerPoint presentation
function Export-PowerPointTitlesAndMessages {
    param (
        [string]$pptPath,
        [string]$csvPath
    )

    # Load PowerPoint application
    $app = New-Object -ComObject PowerPoint.Application
    $app.WindowState = [Microsoft.Office.Interop.PowerPoint.PpWindowState]::ppWindowMinimized

    # $app.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse

    # Open the PowerPoint file
    $presentation = $app.Presentations.Open(
        $pptPath,
        [Microsoft.Office.Core.MsoTriState]::msoFalse,
        [Microsoft.Office.Core.MsoTriState]::msoTrue,
        [Microsoft.Office.Core.MsoTriState]::msoFalse
    )

    # Initialize arrays to store titles, key messages, and slide numbers
    $results = @()

    # Loop through each slide in the presentation
    foreach ($slide in $presentation.Slides) {
        $slideNumber = $slide.SlideNumber

        # Extract the slide title from the placeholder
        $slideTitle = $slide.Shapes |
            Where-Object { $_.PlaceholderFormat.Type -eq 1 -and $_.TextFrame.HasText } |
            ForEach-Object { $_.TextFrame.TextRange.Text }

        # Extract key messages from placeholders (assuming they are in the content placeholders)
        $slideMessages = $slide.Shapes |
            Where-Object { $_.PlaceholderFormat.Type -eq 2 -and $_.TextFrame.HasText } |
            ForEach-Object { $_.TextFrame.TextRange.Text }

        # Store results in an object
        $result = [PSCustomObject]@{
            SlideNumber = $slideNumber
            Title       = $slideTitle -join "; "
            KeyMessages = $slideMessages -join "; "
        }
        $results += $result
    }

    # Export results to UTF-8 encoded CSV
    $results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

    # Close the presentation and release COM objects
    $presentation.Close()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($presentation) | Out-Null
    $app.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
