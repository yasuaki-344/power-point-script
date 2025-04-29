Type DrawArea
    startX As Single
    startY As Single
    drawWidth As Single
    drawHeight As Single
End Type

Sub InsertEnumeration()
    ' This subroutine inserts a matrix visual into the active sheet.

    ' Parameters for the enumeration visual
    Const cm2pt As Single = 28.35
    Const spacing As Single = 10
    Const ratio As Single = 3
    Const rowCount As Integer = 5
    Dim params As DrawArea
    With params
        .startX = 0.9 * cm2pt
        .startY = 4.8 * cm2pt
        .drawWidth = 32 * cm2pt
        .drawHeight = 12.7 * cm2pt
    End With

    ' Get the current slide
    Dim slide As slide: Set slide = ActivePresentation.Slides(1)

    ' Add enumeration elements
    Call AddEnumerationElements(slide, params, spacing, rowCount, ratio)
End Sub

Private Sub AddEnumerationElements(slide As slide, ByRef params As DrawArea, spacing As Single, rowCount As Integer, ratio As Single)
    Dim shapeWidth As Single: shapeWidth = (params.drawWidth - spacing) / 2
    Dim shapeHeight As Single: shapeHeight = (params.drawHeight - (rowCount - 1) * spacing) / rowCount

    Dim labelWidth As Single: labelWidth = (params.drawWidth - spacing) / (1 + ratio)
    Dim elementWidth As Single: elementWidth = (params.drawWidth - spacing) * ratio / (1 + ratio)

    Dim row As Integer: For row = 0 To rowCount - 1
        With slide.Shapes.AddShape(msoShapeRectangle, params.startX, params.startY + (shapeHeight + spacing) * row, labelWidth, shapeHeight)
            .TextFrame.TextRange.Font.Size = 16
        End With
        With slide.Shapes.AddShape(msoShapeRectangle, params.startX + (labelWidth + spacing), params.startY + (shapeHeight + spacing) * row, elementWidth, shapeHeight)
            .Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
            .TextFrame.TextRange.Font.Size = 14
        End With
    Next row
End Sub
