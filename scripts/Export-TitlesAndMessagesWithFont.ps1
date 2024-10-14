function Export-TitlesAndMessagesWithFont {
    param (
        [string]$pptPath,
        [string]$csvPath
    )

    # Load PowerPoint application
    $app = New-Object -ComObject PowerPoint.Application
    # $app.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse
    $app.WindowState = [Microsoft.Office.Interop.PowerPoint.PpWindowState]::ppWindowMinimized

    # Open the PowerPoint file
    $presentation = $app.Presentations.Open(
        $pptPath,
        [Microsoft.Office.Core.MsoTriState]::msoFalse,
        [Microsoft.Office.Core.MsoTriState]::msoTrue,
        [Microsoft.Office.Core.MsoTriState]::msoFalse
    )

    # Initialize array to store titles, key messages, font sizes, and fonts
    $results = @()

    # Loop through each slide in the presentation
    foreach ($slide in $presentation.Slides) {
        $slideNumber = $slide.SlideNumber

        # Extract the title from the placeholder
        $titlePlaceholder = $slide.Shapes.Placeholders |
            Where-Object { $_.PlaceholderFormat.Type -eq [Microsoft.Office.Interop.PowerPoint.PpPlaceholderType]::ppPlaceholderTitle }
        if ($titlePlaceholder) {
            $titleText = $titlePlaceholder.TextFrame.TextRange.Text
            $titleFontSize = $titlePlaceholder.TextFrame.TextRange.Font.Size
            $titleJapaneseFont = $titlePlaceholder.TextFrame.TextRange.Font.NameFarEast
            $titleEnglishFont = $titlePlaceholder.TextFrame.TextRange.Font.NameAscii
        } else {
            $titleText = ""
            $titleFontSize = ""
            $titleJapaneseFont = ""
            $titleEnglishFont = ""
        }

        # Extract the key messages from the placeholder
        $bodyPlaceholder = $slide.Shapes.Placeholders |
            Where-Object { $_.PlaceholderFormat.Type -eq [Microsoft.Office.Interop.PowerPoint.PpPlaceholderType]::ppPlaceholderBody }
        if ($bodyPlaceholder) {
            $bodyText = $bodyPlaceholder.TextFrame.TextRange.Text
            $bodyFontSize = $bodyPlaceholder.TextFrame.TextRange.Font.Size
            $bodyJapaneseFont = $bodyPlaceholder.TextFrame.TextRange.Font.NameFarEast
            $bodyEnglishFont = $bodyPlaceholder.TextFrame.TextRange.Font.NameAscii
        } else {
            $bodyText = ""
            $bodyFontSize = ""
            $bodyJapaneseFont = ""
            $bodyEnglishFont = ""
        }

        # Store results in an object
        $result = [PSCustomObject]@{
            SlideNumber = $slideNumber
            Title       = $titleText
            TitleFontSize = $titleFontSize
            TitleJapaneseFont = $titleJapaneseFont
            TitleEnglishFont = $titleEnglishFont
            KeyMessages = $bodyText
            BodyFontSize = $bodyFontSize
            BodyJapaneseFont = $bodyJapaneseFont
            BodyEnglishFont = $bodyEnglishFont
        }
        $results += $result
    }

    # Export results to CSV
    $results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

    # Close the presentation and release COM objects
    $presentation.Close()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($presentation) | Out-Null
    $app.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
