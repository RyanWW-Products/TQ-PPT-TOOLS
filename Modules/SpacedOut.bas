Attribute VB_Name = "SpacedOut"
Option Explicit

Dim boxTopLeft As Shape
Dim boxBottomRight As Shape

Sub SpacedOut(control As IRibbonControl)

    Dim slideWidth As Single
    Dim slideHeight As Single
    Dim distanceFromSlide As Single
    Dim activeSlide As slide
    Dim distanceFromSlideScale As String
    
    ' Ask the user for the distance from the slide (on a scale from 10 to 100)
    distanceFromSlideScale = InputBox("Enter the distance from the slide on a scale from 10 to 100:", "Distance from Slide")
    
    ' Convert the scale to points, multiplying by 5 times more
    distanceFromSlide = CInt(distanceFromSlideScale) * 50

    Set activeSlide = ActivePresentation.Slides(ActiveWindow.View.slide.SlideIndex)
    slideWidth = ActivePresentation.PageSetup.slideWidth
    slideHeight = ActivePresentation.PageSetup.slideHeight
    
    ' Add a box to the top left corner
    Set boxTopLeft = activeSlide.Shapes.AddShape(msoShapeRectangle, -distanceFromSlide, -distanceFromSlide, 100, 100)
    
    ' Add a box to the bottom right corner
    Set boxBottomRight = activeSlide.Shapes.AddShape(msoShapeRectangle, slideWidth + distanceFromSlide - 100, slideHeight + distanceFromSlide - 100, 100, 100)
    
End Sub

Sub SpacedOutRemover()

    ' Remove the boxes
    boxTopLeft.Delete
    boxBottomRight.Delete

End Sub
