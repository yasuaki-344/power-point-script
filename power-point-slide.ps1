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
    $presentation = $app.Presentations.Add()
    return $presentation
}

function Get-Layout {
    param (
        $presentation,
        [string]
        $layoutName
    )

    if ($null -eq $presentation) 
    {
        Write-Output "null object"
    }

    $layout = $presentation.SlideMaster.CustomLayouts |
        Where-Object { $_.Name -eq $layoutName } |
        Select-Object -First 1

    Write-Output $layout.Name

    return $layout
}

function Add-Slide {
    param (
        $presentation,
        [int]
        $pageNumber,
        $layout
    )
    $slide = $presentation.Slides.AddSlide($pageNumber, $layout)
    return $slide
}

$Application = New-PowerPoint
$presentation = Add-Presentaion($Application)
if ($null -eq $presentation) 
{
    Write-Output "error"
}


$CustomLayout = $presentation.SlideMaster.CustomLayouts |
    Where-Object { $_.Name -eq 'タイトルとコンテンツ' } |
    Select-Object -First 1
# $CustomLayout = Get-Layout($presentation, 'タイトルとコンテンツ')
if ($null -eq $CustomLayout) 
{
    Write-Output "null object"
}
else
{
    Write-Output "not null"
}



# $slide = $presentation.Slides.AddSlide(1, $CustomLayout)
$slide = Add-Slide($presentation, 1, $layout)

$Shape = $slide.Shapes |
    Where-Object { $_.Name -match 'Content Placeholder \d+' } |
    Select-Object -First 1
$Shape.TextFrame.TextRange.Text = "message1`nmessage2"


Exit-PowerPoint($Application)


