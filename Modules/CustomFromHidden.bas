Attribute VB_Name = "CustomFromHidden"
Sub CustomFromHidden(control As IRibbonControl)

    Dim sld As slide
    Dim customShowName As String
    Dim slideTitles As Object ' Use Dictionary instead of Collection
    Dim customShowCounter As Integer
    Dim suffix As String
    Dim alphabet As String

    ' Use Dictionary to easily check for existing keys
    Set slideTitles = CreateObject("Scripting.Dictionary")
    alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

    ' Loop through each slide in the presentation
    For Each sld In ActivePresentation.Slides
        If sld.SlideShowTransition.Hidden = msoTrue Then
            ' If slide is hidden, get the title from the highest text box
            customShowName = GetSlideTitle(sld)

            ' Check if the name already exists in the Dictionary
            If slideTitles.exists(customShowName) Then
                customShowCounter = 1
                suffix = Mid(alphabet, customShowCounter, 1)
                While slideTitles.exists(customShowName & " " & suffix) And customShowCounter <= Len(alphabet)
                    customShowCounter = customShowCounter + 1
                    suffix = Mid(alphabet, customShowCounter, 1)
                Wend
                customShowName = customShowName & " " & suffix
            End If

            ' Add to our dictionary
            slideTitles(customShowName) = True

            ' Truncate custom show name to a reasonable length (e.g., 25 characters)
            customShowName = Left(customShowName, 25)

            ' Create a custom slide show for the hidden slide with the determined name
            ActivePresentation.SlideShowSettings.NamedSlideShows.Add Name:=customShowName, safeArrayOfSlideIDs:=Array(sld.SlideID)
        End If
    Next sld

End Sub

Function GetSlideTitle(sld As slide) As String
    Dim shp As Shape
    Dim highestShape As Shape
    Dim highestPosition As Single
    Dim text As String
    Dim pageBreakPos As Integer

    ' Initialize with a very high value
    highestPosition = 99999

    ' Loop through all shapes in the slide
    For Each shp In sld.Shapes
        ' Check if the shape has text and is positioned higher than the previous highest shape
        If shp.HasTextFrame And shp.TextFrame.HasText And shp.Top < highestPosition Then
            ' Extract text from the shape
            text = shp.TextFrame.TextRange.text

            ' Find the position of a page break or newline character
            pageBreakPos = InStr(1, text, Chr(13)) ' Page break or soft return
            If pageBreakPos > 0 Then
                text = Left(text, pageBreakPos - 1)
            End If

            ' Set highest shape and position
            Set highestShape = shp
            highestPosition = shp.Top
        End If
    Next shp

    ' Return the text of the highest shape
    If Not highestShape Is Nothing Then
        GetSlideTitle = text
    Else
        GetSlideTitle = "Unnamed"
    End If
End Function


