Type DrawArea
    startX As Single
    startY As Single
    drawWidth As Single
    drawHeight As Single
End Type

Sub InsertComparison()
    ' This subroutine inserts comparison visual into the active sheet.

    ' Parameters for the matrix visual
    Const cm2pt As Single = 28.35
    Const spacing As Single = 10
    Dim params As DrawArea
    With params
        .startX = 0.9 * cm2pt
        .startY = 4.8 * cm2pt
        .drawWidth = 32 * cm2pt
        .drawHeight = 12.7 * cm2pt
    End With

    ' Get the current slide
    Dim slide As slide: Set slide = ActivePresentation.Slides(1)

End Sub
