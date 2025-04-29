Type DrawArea
    startX As Single
    startY As Single
    drawWidth As Single
    drawHeight As Single
End Type

Sub InsertBackground()
    ' This subroutine inserts a matrix visual into the active sheet.

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

    ' Add title object
    Const titleHeight As Single = 1.2 * cm2pt
    Call AddTitleWithLine(slide, params, titleHeight)

    ' Add background elements
    Const colCount As Integer = 3
    Dim backgroundParams As DrawArea
    backgroundParams = params
    With backgroundParams
        .startY = .startY + titleHeight + spacing
        .drawHeight = .drawHeight - titleHeight - spacing
    End With

    Call AddBackgroundElements(slide, backgroundParams, spacing, titleHeight, colCount)
End Sub

Private Sub AddTitleWithLine(slide As slide, ByRef params As DrawArea, titleHeight As Single)
    ' タイトル用の矩形を追加
    With slide.Shapes.AddShape(msoShapeRectangle, params.startX, params.startY, params.drawWidth, titleHeight)
        .TextFrame.TextRange.Text = "Background Title"
        .TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter
    End With
End Sub

Private Sub AddBackgroundElements(slide As slide, ByRef params As DrawArea, spacing As Single, titleHeight As Single, colCount As Integer)
    Dim shapeWidth As Single: shapeWidth = (params.drawWidth - (colCount - 1) * spacing) / colCount

    Dim col As Integer: For col = 0 To colCount - 1
        Dim positionX As Single: positionX = params.startX + (shapeWidth + spacing) * col
        With slide.Shapes.AddShape(msoShapeRectangle, positionX, params.startY, shapeWidth, titleHeight)
            .TextFrame.TextRange.Font.Size = 16
        End With

        With slide.Shapes.AddShape(msoShapeRectangle, positionX, params.startY + titleHeight, shapeWidth, params.drawHeight - titleHeight)
            .Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
            .TextFrame.TextRange.Font.Size = 14
        End With
    Next col
End Sub
