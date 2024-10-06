function New-PowerPoint {
    $app = New-Object -ComObject PowerPoint.Application
    $app.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue

    return $app
}
