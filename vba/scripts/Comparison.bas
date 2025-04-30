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
    Const spacing As Single = 15
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
    Call AddComparisonElements(slide, params, spacing)
End Sub


Private Sub AddComparisonElements(slide As slide, ByRef params As DrawArea, spacing As Single)
    Const arrowWidth As Single = 2 * 28.35
    Const arrowHeight As Single = 9 * 28.35
    Const titleHeight As Single = 1.2 * 28.35

    Dim shapeWidth As Single: shapeWidth = (params.drawWidth - 2 * spacing - arrowWidth) / 2

    ' before objects
    With slide.Shapes.AddShape(msoShapeRectangle, params.startX, params.startY, shapeWidth, titleHeight)
        .TextFrame.TextRange.Font.Size = 16
    End With

    With slide.Shapes.AddShape(msoShapeRectangle, params.startX, params.startY + titleHeight, shapeWidth, params.drawHeight - titleHeight)
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
        .TextFrame.TextRange.Font.Size = 14
    End With

    ' arrow object
    Dim arrowPositionX As Single: arrowPositionX = params.startX + shapeWidth + spacing
    Dim arrowStartY As Single: arrowStartY = params.startY + (params.drawHeight - arrowHeight) / 2
    With slide.Shapes.AddShape(msoShapeRightArrow, arrowPositionX, arrowStartY, arrowWidth, arrowHeight)
        .Adjustments.Item(2) = 1#
        .Line.Visible = msoFalse
    End With

    ' after objects
    Dim positionX As Single: positionX = params.startX + shapeWidth + 2 * spacing + arrowWidth
    With slide.Shapes.AddShape(msoShapeRectangle, positionX, params.startY, shapeWidth, titleHeight)
        .TextFrame.TextRange.Font.Size = 16
    End With

    With slide.Shapes.AddShape(msoShapeRectangle, positionX, params.startY + titleHeight, shapeWidth, params.drawHeight - titleHeight)
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
        .TextFrame.TextRange.Font.Size = 14
    End With
End Sub
