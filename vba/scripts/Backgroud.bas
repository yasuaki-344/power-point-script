Sub InsertBackground()
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

    ' Add title object
    Const titleHeight As Single = 1.2 * cm2pt
    Call AddTitleWithLine(slide, startX, startY, drawWidth, titleHeight)

    ' Add matrix elements
    Const colCount As Integer = 3

    Const shapeHeight As Single = drawHeight - spacing - titleHeight
    Const shapeWidth  As Single = (drawWidth - (colCount - 1) * spacing) / colCount
    Const startMatrixY As Single = startY + titleHeight + spacing
    Call AddBackgroundElements(slide, startX, startMatrixY, shapeWidth, shapeHeight, spacing, titleHeight, colCount)
End Sub

Private Sub AddTitleWithLine(slide As slide, startX As Single, startY As Single, drawWidth As Single, titleHeight As Single)
    ' タイトル用の矩形を追加
    With slide.Shapes.AddShape(msoShapeRectangle, startX, startY, drawWidth, titleHeight)
        .TextFrame.TextRange.Text = "Background Title"
        .TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter
    End With
End Sub

Private Sub AddBackgroundElements(slide As slide, startX As Single, startY As Single, shapeWidth As Single, shapeHeight As Single, spacing As Single, titleHeight As Single, colCount As Integer)
    Dim col As Integer: For col = 0 To colCount - 1
        With slide.Shapes.AddShape(msoShapeRectangle, startX + (shapeWidth + spacing) * col, startY, shapeWidth, titleHeight)
            .TextFrame.TextRange.Font.Size = 16
        End With

        With slide.Shapes.AddShape(msoShapeRectangle, startX + (shapeWidth + spacing) * col, startY + titleHeight, shapeWidth, shapeHeight - titleHeight)
            .Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
            .TextFrame.TextRange.Font.Size = 14
        End With
    Next col
End Sub
