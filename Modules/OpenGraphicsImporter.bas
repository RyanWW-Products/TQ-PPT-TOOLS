Attribute VB_Name = "OpenGraphicsImporter"
Sub OpenGraphicsImporter(control As IRibbonControl)
    SelectGraphics.Show vbModal
End Sub

Public Sub ImportSlides(filePath As String)
    Dim sourcePresentation As PowerPoint.presentation
    Dim sourceSlide As PowerPoint.slide
    Dim currentPresentation As PowerPoint.presentation

    ' Set the current active presentation
    Set currentPresentation = Application.ActivePresentation

    ' Open the selected presentation
    Set sourcePresentation = Application.Presentations.Open(filePath)

    ' Import slides
    For Each sourceSlide In sourcePresentation.Slides
        sourceSlide.Copy
        currentPresentation.Slides.Paste (currentPresentation.Windows(1).Selection.SlideRange.SlideIndex)
    Next sourceSlide

    ' Close the source presentation without saving changes
    sourcePresentation.Close
End Sub

Public Sub ToggleFullScreen(ByVal SlideIndex As Long)
    Dim wasInSlideShow As Boolean
    On Error Resume Next ' Ignore errors temporarily
    wasInSlideShow = Not (Application.ActivePresentation.SlideShowWindow Is Nothing)
    On Error GoTo 0 ' Re-enable normal error handling

    With Application.ActivePresentation
        ' Start slide show only if not already in slide show view
        If Not wasInSlideShow Then
            .SlideShowSettings.Run
            DoEvents ' Ensure the slide show window has time to open
            .SlideShowWindow.View.Exit
        End If

        ' Return to the original slide if not in slide show initially
        If Not wasInSlideShow And SlideIndex > 0 And Not Application.ActiveWindow Is Nothing Then
            Application.ActiveWindow.View.GotoSlide SlideIndex
        End If
    End With
End Sub
