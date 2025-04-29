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

    Dim shape As shape
    ' Add title object
    Const titleHeight As Single = 1 * cm2pt
    Set shape = slide.Shapes.AddShape(msoShapeRectangle, startX, startY, drawWidth, titleHeight)
    Set shape = slide.Shapes.AddLine(startX, startY + titleHeight, startX + drawWidth, startY + titleHeight)

    ' Add matrix elements
    Const rowCount As Integer = 3
    Const colCount As Integer = 3

    Const matrixHeight As Single = drawHeight - spacing - titleHeight
    Const shapeWidth  As Single = (drawWidth - (colCount - 1) * spacing) / colCount
    Const shapeHeight As Single = (matrixHeight - (rowCount - 1) * spacing) / rowCount

    Const startMatrixY As Single = startY + titleHeight + spacing
    Dim row As Integer: For row = 0 To rowCount - 1
        Dim col As Integer: For col = 0 To colCount - 1
            Set shape = slide.Shapes.AddShape(msoShapeRectangle, startX + (shapeWidth + spacing) * col, startMatrixY + (shapeHeight + spacing) * row, shapeWidth, shapeHeight)
        Next col
    Next row
End Sub
