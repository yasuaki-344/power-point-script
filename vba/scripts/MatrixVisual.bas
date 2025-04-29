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
    Call AddTitleWithLine(slide, startX, startY, drawWidth, titleHeight)

    ' Add matrix elements
    Const rowCount As Integer = 3
    Const colCount As Integer = 3

    Const matrixHeight As Single = drawHeight - spacing - titleHeight
    Const shapeWidth  As Single = (drawWidth - (colCount - 1) * spacing) / colCount
    Const shapeHeight As Single = (matrixHeight - (rowCount - 1) * spacing) / rowCount

    Const startMatrixY As Single = startY + titleHeight + spacing
    Call AddMatrixElements(slide, startX, startMatrixY, shapeWidth, shapeHeight, spacing, rowCount, colCount)
End Sub

Private Sub AddTitleWithLine(slide As slide, startX As Single, startY As Single, drawWidth As Single, titleHeight As Single)
    ' タイトル用の矩形を追加
    With slide.Shapes.AddShape(msoShapeRectangle, startX, startY, drawWidth, titleHeight)
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
        .TextFrame.TextRange.Text = "Matrix Visual"
        .TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter
        .Line.Visible = msoFalse
    End With
    ' タイトルの下に線を追加
    With slide.Shapes.AddLine(startX, startY + titleHeight, startX + drawWidth, startY + titleHeight)
        .Line.Weight = 2.25
        .Line.ForeColor.RGB = RGB(118, 113, 113)
    End With
End Sub

Private Sub AddMatrixElements(slide As slide, startX As Single, startY As Single, shapeWidth As Single, shapeHeight As Single, spacing As Single, rowCount As Integer, colCount As Integer)
    Dim row As Integer, col As Integer
    Dim shape As shape
    For row = 0 To rowCount - 1
        For col = 0 To colCount - 1
            Set shape = slide.Shapes.AddShape(msoShapeRectangle, startX + (shapeWidth + spacing) * col, startY + (shapeHeight + spacing) * row, shapeWidth, shapeHeight)
        Next col
    Next row
End Sub
