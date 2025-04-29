Sub InsertEnumeration()
    ' This subroutine inserts a matrix visual into the active sheet.

    ' Parameters for the matrix visual
    Const cm2pt As Single = 28.35
    Const startX As Single = 0.9 * cm2pt
    Const startY As Single = 4.8 * cm2pt
    Const drawWidth As Single = 32 * cm2pt
    Const drawHeight As Single = 12.7 * cm2pt
    Const spacing As Single = 10
    Const ratio As Single = 3
    Const rowCount As Integer = 5

    ' Get the current slide
    Dim slide As slide: Set slide = ActivePresentation.Slides(1)

    ' Add enumeration elements
    Call AddEnumerationElements(slide, startX, startY, drawWidth, drawHeight, spacing, rowCount, ratio)
End Sub

Private Sub AddEnumerationElements(slide As slide, startX As Single, startY As Single, drawWidth As Single, drawHeight As Single, spacing As Single, rowCount As Integer, ratio As Single)
    Dim shapeWidth As Single: shapeWidth = (drawWidth - spacing) / 2
    Dim shapeHeight As Single: shapeHeight = (drawHeight - (rowCount - 1) * spacing) / rowCount

    Dim labelWidth As Single: labelWidth = (drawWidth - spacing) / (1 + ratio)
    Dim elementWidth As Single: elementWidth = (drawWidth - spacing) * ratio / (1 + ratio)

    Dim row As Integer: For row = 0 To rowCount - 1
        With slide.Shapes.AddShape(msoShapeRectangle, startX, startY + (shapeHeight + spacing) * row, labelWidth, shapeHeight)
            .TextFrame.TextRange.Font.Size = 16
        End With
        With slide.Shapes.AddShape(msoShapeRectangle, startX + (labelWidth + spacing), startY + (shapeHeight + spacing) * row, elementWidth, shapeHeight)
            .Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
            .TextFrame.TextRange.Font.Size = 14
        End With
    Next row
End Sub
