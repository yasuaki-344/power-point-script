function New-PowerPoint {
    $app = New-Object -ComObject PowerPoint.Application
    $app.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue

    return $app
}

function Exit-PowerPoint {
    param (
        $app
    )
    # $app.Quit()

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) | Out-Null
    $app = $null
    # Remove-Variable -Name Application -ErrorAction SilentlyContinue
    [System.GC]::Collect()
}

function Add-Presentaion {
    param (
        $app
    )
    $presentation = $Application.Presentations.Add()
    return $presentation
}

$Application = New-PowerPoint
$presentation = Add-Presentaion($Application)
Exit-PowerPoint($Application)
