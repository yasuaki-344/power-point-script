Sub InsertMatrixVisual()
    ' This subroutine inserts a matrix visual into the active sheet.

    ' Parameters for the matrix visual
    Const cm2pt As Single = 28.35
    Const startX As Single = 0.9 * cm2pt
    Const startY As Single = 4.8 * cm2pt
    Const drawWidth As Single = 32 * cm2pt
    Const drawHeight As Single = 12.7 * cm2pt
    Const spacing As Single = 10

    ' Get the current slide
    Dim slide As slide: Set slide = ActivePresentation.Slides(1)

End Sub
