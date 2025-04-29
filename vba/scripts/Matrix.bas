Type DrawArea
    startX As Single
    startY As Single
    drawWidth As Single
    drawHeight As Single
End Type

Sub InsertMatrix()
    ' This subroutine inserts a matrix visual into the active sheet.

    ' Parameters for the matrix visual
    Const cm2pt As Single = 28.35
    Const spacing As Single = 5
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
    Const titleHeight As Single = 1 * cm2pt
    Call AddTitleWithLine(slide, params, titleHeight)

    ' Add matrix elements
    Const rowCount As Integer = 4
    Const colCount As Integer = 5

    Dim matrixParams As DrawArea
    matrixParams = params
    With matrixParams
        .startY = .startY + titleHeight + spacing
        .drawHeight = .drawHeight - titleHeight - spacing
    End With
    Call AddMatrixElements(slide, matrixParams, spacing, rowCount, colCount)
End Sub

Private Sub AddTitleWithLine(slide As slide, ByRef params As DrawArea, titleHeight As Single)
    ' タイトル用の矩形を追加
    With slide.Shapes.AddShape(msoShapeRectangle, params.startX, params.startY, params.drawWidth, titleHeight)
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
        .TextFrame.TextRange.Text = "Matrix Visual"
        .TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter
        .Line.Visible = msoFalse
    End With
    ' タイトルの下に線を追加
    With slide.Shapes.AddLine(params.startX, params.startY + titleHeight, params.startX + params.drawWidth, params.startY + titleHeight)
        .Line.Weight = 2.25
        .Line.ForeColor.RGB = RGB(118, 113, 113)
    End With
End Sub

Private Sub AddMatrixElements(slide As slide, ByRef params As DrawArea, spacing As Single, rowCount As Integer, colCount As Integer)
    Dim shapeWidth As Single: shapeWidth = (params.drawWidth - colCount * spacing) / (colCount + 1)
    Dim shapeHeight As Single: shapeHeight = (params.drawHeight - rowCount * spacing) / (rowCount + 1)

    Dim row As Integer: For row = 0 To rowCount
        Dim col As Integer: For col = 0 To colCount
            If row > 0 Or col > 0 Then
                With slide.Shapes.AddShape(msoShapeRectangle, params.startX + (shapeWidth + spacing) * col, params.startY + (shapeHeight + spacing) * row, shapeWidth, shapeHeight)
                    If row = 0 Or col = 0 Then
                        .TextFrame.TextRange.Font.Size = 16
                    Else
                        .Fill.ForeColor.RGB = RGB(255, 255, 255)
                        .TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
                        .TextFrame.TextRange.Font.Size = 14
                    End If
                End With
            End If
        Next col
    Next row
End Sub
