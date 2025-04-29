Sub InsertEnumeration()
    ' This subroutine inserts a matrix visual into the active sheet.

    ' Parameters for the matrix visual
    Const cm2pt As Single = 28.35
    Const startX As Single = 0.9 * cm2pt
    Const startY As Single = 4.8 * cm2pt
    Const drawWidth As Single = 32 * cm2pt
    Const drawHeight As Single = 12.7 * cm2pt
    Const spacing As Single = 5

    ' Get the current slide
    Dim slide As slide: Set slide = ActivePresentation.Slides(1)

    ' Add matrix elements
    Const rowCount As Integer = 5
    Call AddMatrixElements(slide, startX, startY, drawWidth, drawHeight, spacing, rowCount)
End Sub

Private Sub AddMatrixElements(slide As slide, startX As Single, startY As Single, drawWidth As Single, drawHeight As Single, spacing As Single, rowCount As Integer)
    Dim shapeWidth As Single: shapeWidth = (drawWidth - spacing) / 2
    Dim shapeHeight As Single: shapeHeight = (drawHeight - (rowCount - 1) * spacing) / rowCount

    Dim row As Integer: For row = 0 To rowCount - 1
        Dim col As Integer: For col = 0 To 1
            If col = 0 Then
                With slide.Shapes.AddShape(msoShapeRectangle, startX + (shapeWidth + spacing) * col, startY + (shapeHeight + spacing) * row, shapeWidth, shapeHeight)
                    .TextFrame.TextRange.Font.Size = 16
                End With
            Else
                With slide.Shapes.AddShape(msoShapeRectangle, startX + (shapeWidth + spacing) * col, startY + (shapeHeight + spacing) * row, shapeWidth, shapeHeight)
                    .Fill.ForeColor.RGB = RGB(255, 255, 255)
                    .TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
                    .TextFrame.TextRange.Font.Size = 14
                End With
            End If
        Next col
    Next row
End Sub
