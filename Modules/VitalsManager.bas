Attribute VB_Name = "VitalsManager"
Option Explicit

' ============================================================================
' VITALS SCATTERPLOT MACRO SYSTEM
' ============================================================================
' This module provides tools for creating and managing scatter plots
' specifically designed for vitals graphs in clinical trial presentations.
'
' Template Location: %APPDATA%\Microsoft\AddIns\Trial Ex Addin\Scatterplots\Scatterplot1.pptx
'
' Template Shapes (Z-order bottom to top):
'   1. Vital1Tab - Background tab with vital name (e.g., "Blood Pressure")
'   2. ScatterBG1 - Background shape behind the graph
'   3. NormalRange1 - Normal range rectangle (width preset, height adjustable)
'   4. ScatterTemplate1 - The actual scatter graph
' ============================================================================

' --- Windows API for Sleep (avoids 100% CPU in DoEvents loops) ---
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' --- Constants ---
Private Const TEMPLATE_FOLDER As String = "\Microsoft\AddIns\Trial Ex Addin\Scatterplots\"
Private Const TEMPLATE_FILE As String = "Scatterplot1.pptx"
Private Const TAG_VITALS_GRAPH As String = "VITALSGRAPH"
Private Const TAG_GRAPH_NAME As String = "GRAPHNAME"
Private Const TAG_GRAPH_GROUP As String = "GRAPHGROUP"
Private Const TAG_NORMAL_RANGE As String = "NORMALRANGE"
Private Const TAG_NORMAL_YLOW As String = "NORMALYLOW"
Private Const TAG_NORMAL_YHIGH As String = "NORMALYHIGH"
Private Const TAG_COMPONENT_TYPE As String = "COMPONENTTYPE"
Private Const TAG_STYLE_NUMBER As String = "STYLE_NUMBER"
Private Const TAG_SNAP_REL_LEFT As String = "SNAP_REL_LEFT"
Private Const TAG_SNAP_REL_TOP As String = "SNAP_REL_TOP"
Private Const TAG_SNAP_REL_WIDTH As String = "SNAP_REL_WIDTH"
Private Const TAG_SNAP_REL_HEIGHT As String = "SNAP_REL_HEIGHT"
Private Const TAG_SNAP_ROTATION As String = "SNAP_ROTATION"
Private Const TAG_BLOCK_ABOVE_RATIO As String = "BLOCK_ABOVE_RATIO"

' --- Module-level margin shape names (persist during session) ---
Private m_TopMarginShapeName As String
Private m_BottomMarginShapeName As String

' ============================================================================
' PUBLIC ENTRY POINTS
' ============================================================================

Public Sub CreateVitalsScatterplot()
    ' Main entry point to create a new vitals scatterplot from template
    Dim templatePath As String
    Dim templatePres As presentation
    Dim targetSlide As slide
    Dim graphName As String
    Dim xBounds As String
    Dim xMin As Double, xMax As Double
    Dim vitalName As String
    
    On Error GoTo ErrorHandler
    
    ' Build template path
    templatePath = Environ("APPDATA") & TEMPLATE_FOLDER & TEMPLATE_FILE
    
    ' Check template exists
    If Dir(templatePath) = "" Then
        MsgBox "Template file not found:" & vbCrLf & vbCrLf & templatePath & vbCrLf & vbCrLf & _
               "Please ensure the template exists in the correct location.", _
               vbCritical, "Template Not Found"
        Exit Sub
    End If
    
    ' Ensure we have an active presentation and slide
    If ActivePresentation Is Nothing Then
        MsgBox "Please open a presentation first.", vbExclamation, "No Presentation"
        Exit Sub
    End If
    
    If ActiveWindow.Selection.SlideRange.count = 0 Then
        MsgBox "Please select a slide first.", vbExclamation, "No Slide Selected"
        Exit Sub
    End If
    
    Set targetSlide = ActiveWindow.Selection.SlideRange(1)
    
    ' Get graph name from user (also used as the tab label)
    graphName = InputBox("Enter a name for this scatterplot:" & vbCrLf & _
                         "(e.g., Blood Pressure, Heart Rate, Temperature)" & vbCrLf & vbCrLf & _
                         "This will identify the graph and appear on the tab.", _
                         "Scatterplot Name", "Blood Pressure")
    If graphName = "" Then Exit Sub
    vitalName = graphName
    
    ' Check if existing graphs exist and offer to use same bounds
    Dim existingGraphs As Collection
    Dim useExistingBounds As VbMsgBoxResult
    Dim existingGraph As Shape
    
    Set existingGraphs = GetAllVitalsGraphs()
    
    If existingGraphs.count > 0 Then
        ' Get bounds from the first existing graph
        Set existingGraph = existingGraphs(1)
        If existingGraph.HasChart Then
            On Error Resume Next
            xMin = existingGraph.Chart.Axes(xlCategory).MinimumScale
            xMax = existingGraph.Chart.Axes(xlCategory).MaximumScale
            On Error GoTo ErrorHandler
            
            useExistingBounds = MsgBox("An existing vitals graph was found with bounds:" & vbCrLf & vbCrLf & _
                                       GetGraphBounds(existingGraph) & vbCrLf & vbCrLf & _
                                       "Would you like to use the same horizontal bounds?", _
                                       vbYesNo + vbQuestion, "Use Existing Bounds?")
            
            If useExistingBounds = vbYes Then
                ' Skip the bounds input - we already have xMin and xMax
                GoTo SkipBoundsInput
            End If
        End If
    End If
    
    ' Get X-axis bounds (date/time format)
    xBounds = InputBox("Enter horizontal time bounds using one of these formats:" & vbCrLf & vbCrLf & _
                       "A. mm/dd/yyyy hh:mm:ss - mm/dd/yyyy hh:mm:ss" & vbCrLf & _
                       "B. mm/dd/yyyy hh:mm - mm/dd/yyyy hh:mm" & vbCrLf & _
                       "C. mm/dd/yyyy - mm/dd/yyyy" & vbCrLf & _
                       "D. hh:mm:ss - hh:mm:ss" & vbCrLf & _
                       "E. hh:mm - hh:mm" & vbCrLf & vbCrLf & _
                       "Example: 01/01/2024 08:00 - 01/01/2024 20:00", _
                       "Horizontal Time Bounds", "01/01/2024 00:00 - 01/01/2024 23:59")
    If xBounds = "" Then Exit Sub
    If Not ParseDateTimeBounds(xBounds, xMin, xMax) Then
        MsgBox "Invalid date/time format. Please use one of the supported formats.", vbExclamation, "Invalid Format"
        Exit Sub
    End If
    
SkipBoundsInput:
    ' Y-axis bounds will be determined automatically
    
    ' Open template (hidden, read-only)
    Set templatePres = Application.Presentations.Open( _
        fileName:=templatePath, _
        ReadOnly:=msoTrue, _
        WithWindow:=msoFalse)
    
    ' Copy and configure the scatterplot group
    CopyAndConfigureTemplate templatePres, targetSlide, graphName, vitalName, xMin, xMax
    
    ' Close template
    templatePres.Saved = True
    templatePres.Close
    Set templatePres = Nothing
    
    MsgBox "Scatterplot '" & graphName & "' created successfully!" & vbCrLf & vbCrLf & _
           "You can now add your data by double-clicking the graph.", _
           vbInformation, "Success"
    
    Exit Sub
    
ErrorHandler:
    If Not templatePres Is Nothing Then
        templatePres.Saved = True
        templatePres.Close
        Set templatePres = Nothing
    End If
    MsgBox "Error creating scatterplot: " & Err.Description, vbCritical, "Error"
End Sub

Public Sub ManageScatterplots(control As IRibbonControl)
    ' Opens the graph manager form for editing multiple graphs
    frmGraphManager.Show
End Sub

Public Sub SnapAllNormalRanges()
    ' Repositions all normal range rectangles to their linked graphs
    Dim allShapes As New Collection
    Dim shp As Shape
    Dim graphShape As Shape
    Dim graphName As String
    Dim count As Long
    
    On Error Resume Next
    
    CollectAllVitalsShapes allShapes
    For Each shp In allShapes
        If shp.Tags(TAG_NORMAL_RANGE) <> "" Then
            graphName = shp.Tags(TAG_NORMAL_RANGE)
            Set graphShape = FindGraphByName(graphName)
            
            If Not graphShape Is Nothing Then
                RepositionNormalRange shp, graphShape
                count = count + 1
            End If
        End If
    Next shp
    
    On Error GoTo 0
    
    If count > 0 Then
        MsgBox count & " normal range rectangle(s) repositioned.", vbInformation, "Snap Complete"
    Else
        MsgBox "No normal range rectangles found to reposition.", vbInformation, "Snap Complete"
    End If
End Sub

Public Sub SaveAsTemplate()
    ' Saves the selected shapes as a new template
    Dim templatePath As String
    Dim templateFolder As String
    Dim newPres As presentation
    Dim newSlide As slide
    
    On Error GoTo ErrorHandler
    
    If ActiveWindow.Selection.ShapeRange.count = 0 Then
        MsgBox "Please select the shapes to save as template:" & vbCrLf & _
               "- ScatterTemplate1 (graph)" & vbCrLf & _
               "- ScatterBG1 (background)" & vbCrLf & _
               "- NormalRange1 (normal range rectangle)" & vbCrLf & _
               "- Vital1Tab (tab label)", _
               vbExclamation, "No Selection"
        Exit Sub
    End If
    
    templateFolder = Environ("APPDATA") & TEMPLATE_FOLDER
    templatePath = templateFolder & TEMPLATE_FILE
    
    ' Create folder if it doesn't exist
    If Dir(templateFolder, vbDirectory) = "" Then
        MkDir templateFolder
    End If
    
    ' Confirm overwrite if exists
    If Dir(templatePath) <> "" Then
        If MsgBox("Template already exists. Overwrite?", vbYesNo + vbQuestion, "Confirm Overwrite") = vbNo Then
            Exit Sub
        End If
    End If
    
    ' Create new presentation
    Set newPres = Application.Presentations.Add(msoFalse)
    Set newSlide = newPres.Slides.Add(1, ppLayoutBlank)
    
    ' Copy selected shapes
    ActiveWindow.Selection.ShapeRange.Copy
    newSlide.Shapes.Paste
    
    ' Save and close
    newPres.SaveAs templatePath
    newPres.Close
    
    MsgBox "Template saved successfully to:" & vbCrLf & templatePath, vbInformation, "Template Saved"
    
    Exit Sub
    
ErrorHandler:
    If Not newPres Is Nothing Then
        newPres.Close
    End If
    MsgBox "Error saving template: " & Err.Description, vbCritical, "Error"
End Sub

' ============================================================================
' ORIENTATION & MARGINS
' ============================================================================

Public Sub SetOrientMargins()
    ' Lets user pick shapes for top and bottom margin boundaries.
    ' Hides the form, prompts for each margin, then re-shows the form.
    Dim result As VbMsgBoxResult
    Dim sldIdx As Long
    
    sldIdx = ActiveWindow.View.slide.SlideIndex
    
    frmGraphManager.Hide
    DoEvents
    
    ' === TOP MARGIN ===
    result = MsgBox("After clicking OK, click on a shape to define the TOP margin." & vbCrLf & vbCrLf & _
                    "The bottom edge of that shape will be the top boundary" & vbCrLf & _
                    "for graph orientation." & vbCrLf & vbCrLf & _
                    "Click Cancel to use the top of the slide.", _
                    vbOKCancel + vbInformation, "Set Top Margin")
    
    If result = vbOK Then
        ' Deselect everything so we can detect a fresh click
        ActiveWindow.View.GotoSlide sldIdx
        DoEvents
        
        m_TopMarginShapeName = WaitForShapeSelection(10)
        
        If m_TopMarginShapeName <> "" Then
            MsgBox "Top margin set to: " & m_TopMarginShapeName, vbInformation, "Top Margin Set"
        Else
            MsgBox "No shape selected. Top of slide will be used.", vbInformation, "Top Margin"
        End If
    Else
        m_TopMarginShapeName = ""
    End If
    
    ' === BOTTOM MARGIN ===
    result = MsgBox("After clicking OK, click on a shape to define the BOTTOM margin." & vbCrLf & vbCrLf & _
                    "The top edge of that shape will be the bottom boundary" & vbCrLf & _
                    "for graph orientation." & vbCrLf & vbCrLf & _
                    "Click Cancel to use the bottom of the slide.", _
                    vbOKCancel + vbInformation, "Set Bottom Margin")
    
    If result = vbOK Then
        ActiveWindow.View.GotoSlide sldIdx
        DoEvents
        
        m_BottomMarginShapeName = WaitForShapeSelection(10)
        
        If m_BottomMarginShapeName <> "" Then
            MsgBox "Bottom margin set to: " & m_BottomMarginShapeName, vbInformation, "Bottom Margin Set"
        Else
            MsgBox "No shape selected. Bottom of slide will be used.", vbInformation, "Bottom Margin"
        End If
    Else
        m_BottomMarginShapeName = ""
    End If
    
    ' Summary
    Dim topStr As String, bottomStr As String
    If m_TopMarginShapeName <> "" Then topStr = m_TopMarginShapeName Else topStr = "(top of slide)"
    If m_BottomMarginShapeName <> "" Then bottomStr = m_BottomMarginShapeName Else bottomStr = "(bottom of slide)"
    
    MsgBox "Margins configured:" & vbCrLf & vbCrLf & _
           "Top boundary: " & topStr & vbCrLf & _
           "Bottom boundary: " & bottomStr, _
           vbInformation, "Margins Set"
    
    frmGraphManager.Show
End Sub

Private Function WaitForShapeSelection(maxSeconds As Long) As String
    ' Waits up to maxSeconds for the user to click a shape on the slide.
    ' Returns the shape name, or "" if nothing selected in time.
    Dim elapsed As Long
    Dim checkInterval As Long
    checkInterval = 100  ' milliseconds
    
    elapsed = 0
    Do While elapsed < (maxSeconds * 1000)
        DoEvents
        Sleep checkInterval
        elapsed = elapsed + checkInterval
        
        ' Check if user has selected a shape
        On Error Resume Next
        If ActiveWindow.Selection.Type = ppSelectionShapes Then
            WaitForShapeSelection = ActiveWindow.Selection.ShapeRange(1).Name
            On Error GoTo 0
            Exit Function
        End If
        On Error GoTo 0
    Loop
    
    WaitForShapeSelection = ""
End Function

Public Sub OrientSelectedGraphs(graphNames As Collection)
    ' Vertically distributes the selected graphs within the defined margins.
    ' 10% of available height is total margin space, split across N+1 gaps.
    ' Remaining 90% is shared equally among the N graphs.
    
    Dim sld As slide
    Dim topBound As Double
    Dim bottomBound As Double
    Dim availHeight As Double
    Dim numGraphs As Long
    Dim graphHeight As Double
    Dim gapSize As Double
    Dim currentY As Double
    Dim i As Long
    Dim graphShape As Shape
    
    Set sld = ActiveWindow.View.slide
    numGraphs = graphNames.count
    If numGraphs = 0 Then Exit Sub
    
    ' --- Determine vertical bounds ---
    topBound = 0  ' Default: top of slide
    bottomBound = ActivePresentation.PageSetup.slideHeight  ' Default: bottom of slide
    
    ' Check for top margin shape
    If m_TopMarginShapeName <> "" Then
        On Error Resume Next
        Dim topMarginShape As Shape
        Set topMarginShape = sld.Shapes(m_TopMarginShapeName)
        If Not topMarginShape Is Nothing Then
            topBound = topMarginShape.Top + topMarginShape.Height  ' Bottom edge
        End If
        On Error GoTo 0
    End If
    
    ' Check for bottom margin shape
    If m_BottomMarginShapeName <> "" Then
        On Error Resume Next
        Dim bottomMarginShape As Shape
        Set bottomMarginShape = sld.Shapes(m_BottomMarginShapeName)
        If Not bottomMarginShape Is Nothing Then
            bottomBound = bottomMarginShape.Top  ' Top edge
        End If
        On Error GoTo 0
    End If
    
    ' --- Find max above-extension across all graphs ---
    ' Only tabs positioned ABOVE the graph (relTop < 0) count.
    ' Tabs beside/below the graph (e.g. style 2 rotated on right) add no
    ' extra vertical space.  Tab dimensions are FIXED (template size),
    ' so we measure them in absolute points, not ratios.
    Dim maxAboveExt As Double
    Dim groupID As String
    Dim components As Collection
    Dim comp As Shape
    Dim tabIsAbove As Boolean
    maxAboveExt = 0
    
    For i = 1 To numGraphs
        Set graphShape = FindGraphByName(graphNames(i))
        If Not graphShape Is Nothing Then
            groupID = graphShape.Tags(TAG_GRAPH_GROUP)
            Set components = FindGroupComponents(groupID)
            For Each comp In components
                If comp.Tags(TAG_COMPONENT_TYPE) = "TAB" Then
                    tabIsAbove = False
                    If comp.Tags(TAG_SNAP_REL_TOP) <> "" Then
                        tabIsAbove = (CDbl(comp.Tags(TAG_SNAP_REL_TOP)) < 0)
                    Else
                        tabIsAbove = True  ' Legacy fallback: assume above
                    End If
                    If tabIsAbove Then
                        If comp.Height > maxAboveExt Then maxAboveExt = comp.Height
                    End If
                    Exit For
                End If
            Next comp
        End If
    Next i
    
    ' --- Calculate layout ---
    availHeight = bottomBound - topBound
    If availHeight <= 0 Then
        MsgBox "Invalid margin boundaries. Check your margin shapes.", vbExclamation, "Orient Error"
        Exit Sub
    End If
    
    ' Total above space = one fixed-size tab extension per graph
    Dim totalAboveSpace As Double
    totalAboveSpace = numGraphs * maxAboveExt
    
    ' 10% of available height is total gap space, distributed across N+1 gaps
    Dim totalGapSpace As Double
    totalGapSpace = availHeight * 0.1
    gapSize = totalGapSpace / (numGraphs + 1)
    
    ' Remaining space shared equally among graphs
    graphHeight = (availHeight - totalGapSpace - totalAboveSpace) / numGraphs
    
    If graphHeight <= 0 Then
        MsgBox "Not enough vertical space for " & numGraphs & " graph(s).", vbExclamation, "Orient Error"
        Exit Sub
    End If
    
    ' --- Position each graph ---
    ' Each block: [gap] [above ext] [graph] [gap] [above ext] [graph] ...
    currentY = topBound + gapSize
    
    For i = 1 To numGraphs
        Set graphShape = FindGraphByName(graphNames(i))
        If Not graphShape Is Nothing Then
            UngroupVitalsGroup graphShape
            graphShape.Top = currentY + maxAboveExt
            graphShape.Height = graphHeight
            
            ' Snap all group components to the repositioned graph
            SnapGroupToGraph graphShape
            
            currentY = currentY + maxAboveExt + graphHeight + gapSize
        End If
    Next i
    
    MsgBox numGraphs & " graph(s) oriented successfully.", vbInformation, "Orient Complete"
End Sub

Public Function GetTopMarginName() As String
    GetTopMarginName = m_TopMarginShapeName
End Function

Public Function GetBottomMarginName() As String
    GetBottomMarginName = m_BottomMarginShapeName
End Function

' ============================================================================
' TEMPLATE CONFIGURATION
' ============================================================================

Private Sub CopyAndConfigureTemplate(templatePres As presentation, targetSlide As slide, _
                                      graphName As String, vitalName As String, _
                                      xMin As Double, xMax As Double, _
                                      Optional styleNumber As Long = 1)
    Dim templateSlide As slide
    Dim shp As Shape
    Dim pastedShapes As ShapeRange
    Dim graphShape As Shape
    Dim bgShape As Shape
    Dim normalShape As Shape
    Dim tabShape As Shape
    Dim groupID As String
    
    Set templateSlide = templatePres.Slides(1)
    
    ' Generate unique group ID
    groupID = graphName & "_" & Format(Now, "yyyymmddhhnnss")
    
    ' Paste all template shapes and collect individual items.
    ' Handles both grouped templates (group named "VitalsGroup1")
    ' and ungrouped templates (loose shapes) for backward compatibility.
    Dim pastedItems As New Collection
    Dim ungroupedRange As ShapeRange
    Dim k As Long
    
    For Each shp In templateSlide.Shapes
        shp.Copy
        Set pastedShapes = targetSlide.Shapes.Paste
        
        If pastedShapes(1).Type = msoGroup Then
            ' Ungroup to get individual shapes
            Set ungroupedRange = pastedShapes(1).Ungroup
            For k = 1 To ungroupedRange.count
                pastedItems.Add ungroupedRange(k)
            Next k
        Else
            pastedItems.Add pastedShapes(1)
        End If
    Next shp
    
    ' Configure each individual shape by its template name.
    ' Use Like "Name*" matching because PPT may append suffixes
    ' (e.g. "ScatterTemplate1 2") to avoid naming conflicts on paste.
    Dim item As Shape
    For Each item In pastedItems
        Select Case True
            Case item.Name Like "ScatterTemplate1*"
                Set graphShape = item
                graphShape.Name = graphName & "_Graph"
                graphShape.Tags.Add TAG_VITALS_GRAPH, "TRUE"
                graphShape.Tags.Add TAG_GRAPH_NAME, graphName
                graphShape.Tags.Add TAG_GRAPH_GROUP, groupID
                graphShape.Tags.Add TAG_COMPONENT_TYPE, "GRAPH"
                
            Case item.Name Like "ScatterBG1*"
                Set bgShape = item
                bgShape.Name = graphName & "_BG"
                bgShape.Tags.Add TAG_GRAPH_GROUP, groupID
                bgShape.Tags.Add TAG_COMPONENT_TYPE, "BACKGROUND"
                
            Case item.Name Like "NormalRange1*"
                Set normalShape = item
                normalShape.Name = graphName & "_NormalRange"
                normalShape.Tags.Add TAG_GRAPH_GROUP, groupID
                normalShape.Tags.Add TAG_COMPONENT_TYPE, "NORMALRANGE"
                normalShape.Tags.Add TAG_NORMAL_RANGE, graphName
                ' Default Y values (will be updated when user sets normal range)
                normalShape.Tags.Add TAG_NORMAL_YLOW, "0"
                normalShape.Tags.Add TAG_NORMAL_YHIGH, "0"
                ' Hide and vertically center within graph bounds so it
                ' doesn't extend the group bounding box in the template
                normalShape.Visible = msoFalse
                
            Case item.Name Like "Vital1Tab*"
                Set tabShape = item
                tabShape.Name = graphName & "_Tab"
                tabShape.Tags.Add TAG_GRAPH_GROUP, groupID
                tabShape.Tags.Add TAG_COMPONENT_TYPE, "TAB"
                ' Set the vital name text
                If tabShape.HasTextFrame Then
                    tabShape.TextFrame.TextRange.text = vitalName
                End If
        End Select
    Next item
    
    ' Reset all axis bounds to auto first (clears any template presets)
    If Not graphShape Is Nothing Then
        If graphShape.HasChart Then
            On Error Resume Next
            graphShape.Chart.Axes(xlCategory).MinimumScaleIsAuto = True
            graphShape.Chart.Axes(xlCategory).MaximumScaleIsAuto = True
            graphShape.Chart.Axes(xlValue).MinimumScaleIsAuto = True
            graphShape.Chart.Axes(xlValue).MaximumScaleIsAuto = True
            On Error GoTo 0
        End If
    End If
    
    ' Configure graph axis bounds (X only, Y is automatic)
    If Not graphShape Is Nothing Then
        ConfigureGraphXBounds graphShape, xMin, xMax
    End If
    
    ' Store style number
    If Not graphShape Is Nothing Then
        graphShape.Tags.Add TAG_STYLE_NUMBER, CStr(styleNumber)
    End If
    
    ' Calculate and store relative snap positions from template layout.
    ' At this point, pasted shapes retain their template-relative positions,
    ' so we can capture how each component relates to the plot area.
    If Not graphShape Is Nothing Then
        Dim tpLeft As Double, tpTop As Double, tpWidth As Double, tpHeight As Double
        GetPlotAreaSlideCoords graphShape, tpLeft, tpTop, tpWidth, tpHeight
        
        If tpWidth > 0 And tpHeight > 0 Then
            If Not bgShape Is Nothing Then
                StoreSnapTags bgShape, tpLeft, tpTop, tpWidth, tpHeight
            End If
            If Not tabShape Is Nothing Then
                StoreSnapTags tabShape, tpLeft, tpTop, tpWidth, tpHeight
            End If
            
            ' Calculate how much extra space extends above the graph shape
            ' (used by OrientSelectedGraphs for layout)
            Dim blockAboveRatio As Double
            blockAboveRatio = 0
            If Not tabShape Is Nothing Then
                If tabShape.Top < graphShape.Top Then
                    blockAboveRatio = (graphShape.Top - tabShape.Top) / graphShape.Height
                End If
            End If
            If Not bgShape Is Nothing Then
                If bgShape.Top < graphShape.Top Then
                    Dim bgAbove As Double
                    bgAbove = (graphShape.Top - bgShape.Top) / graphShape.Height
                    If bgAbove > blockAboveRatio Then blockAboveRatio = bgAbove
                End If
            End If
            graphShape.Tags.Add TAG_BLOCK_ABOVE_RATIO, CStr(blockAboveRatio)
        End If
    End If
    
    ' Arrange Z-order (bottom to top: Tab, BG, Normal, Graph)
    If Not tabShape Is Nothing Then tabShape.ZOrder msoSendToBack
    If Not bgShape Is Nothing Then bgShape.ZOrder msoSendToBack: bgShape.ZOrder msoBringForward
    If Not normalShape Is Nothing Then normalShape.ZOrder msoBringForward
    If Not graphShape Is Nothing Then graphShape.ZOrder msoBringToFront
    
    ' Position graph just below the slide (offscreen) and snap components
    ' SnapGroupToGraph will also handle regrouping into a PPT group
    If Not graphShape Is Nothing Then
        graphShape.Top = ActivePresentation.PageSetup.slideHeight + 10
        SnapGroupToGraph graphShape
    End If
End Sub

Private Sub StoreSnapTags(shp As Shape, plotLeft As Double, plotTop As Double, _
                           plotWidth As Double, plotHeight As Double)
    ' Stores relative snap position tags on a shape.
    ' These define how the shape relates to the graph's plot area,
    ' enabling style-aware positioning when the graph moves/resizes.
    shp.Tags.Add TAG_SNAP_REL_LEFT, CStr((shp.Left - plotLeft) / plotWidth)
    shp.Tags.Add TAG_SNAP_REL_TOP, CStr((shp.Top - plotTop) / plotHeight)
    shp.Tags.Add TAG_SNAP_REL_WIDTH, CStr(shp.Width / plotWidth)
    shp.Tags.Add TAG_SNAP_REL_HEIGHT, CStr(shp.Height / plotHeight)
    shp.Tags.Add TAG_SNAP_ROTATION, CStr(shp.Rotation)
End Sub

' ============================================================================
' GRAPH AXIS CONFIGURATION
' ============================================================================

Public Sub ConfigureGraphXBounds(graphShape As Shape, xMin As Double, xMax As Double)
    ' Configures only X-axis bounds, Y-axis remains automatic
    Dim cht As Chart
    
    If Not graphShape.HasChart Then Exit Sub
    
    Set cht = graphShape.Chart
    
    On Error Resume Next
    
    With cht.Axes(xlCategory)
        .MinimumScaleIsAuto = False
        .MaximumScaleIsAuto = False
        .MinimumScale = xMin
        .MaximumScale = xMax
    End With
    
    With cht.Axes(xlValue)
        .MinimumScaleIsAuto = True
        .MaximumScaleIsAuto = True
    End With
    
    On Error GoTo 0
End Sub

Public Sub ConfigureGraphBounds(graphShape As Shape, xMin As Double, xMax As Double, _
                                 yMin As Double, yMax As Double)
    ' Configures both X and Y axis bounds (used by Edit Bounds in manager)
    Dim cht As Chart
    
    If Not graphShape.HasChart Then Exit Sub
    
    Set cht = graphShape.Chart
    
    On Error Resume Next
    
    With cht.Axes(xlCategory)
        .MinimumScaleIsAuto = False
        .MaximumScaleIsAuto = False
        .MinimumScale = xMin
        .MaximumScale = xMax
    End With
    
    With cht.Axes(xlValue)
        .MinimumScaleIsAuto = False
        .MaximumScaleIsAuto = False
        .MinimumScale = yMin
        .MaximumScale = yMax
    End With
    
    On Error GoTo 0
End Sub

' ============================================================================
' GRAPH LOOKUP & COMPONENT FUNCTIONS
' ============================================================================

Private Sub CollectAllVitalsShapes(ByRef allShapes As Collection)
    ' Collects all shapes in the presentation, including children of PPT groups.
    ' This allows tag-based lookups to find shapes inside grouped vitals.
    Dim sld As slide
    Dim shp As Shape
    Dim gi As Long
    
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If shp.Type = msoGroup Then
                For gi = 1 To shp.GroupItems.count
                    allShapes.Add shp.GroupItems(gi)
                Next gi
            Else
                allShapes.Add shp
            End If
        Next shp
    Next sld
End Sub

Public Function GetAllVitalsGraphs() As Collection
    ' Returns a collection of all vitals graph shapes in the presentation
    Dim graphs As New Collection
    Dim allShapes As New Collection
    Dim shp As Shape
    
    CollectAllVitalsShapes allShapes
    For Each shp In allShapes
        If shp.Tags(TAG_VITALS_GRAPH) = "TRUE" Then
            graphs.Add shp
        End If
    Next shp
    
    Set GetAllVitalsGraphs = graphs
End Function

Public Function FindGraphByName(graphName As String) As Shape
    ' Finds a vitals graph by its name tag (searches inside PPT groups too)
    Dim allShapes As New Collection
    Dim shp As Shape
    
    CollectAllVitalsShapes allShapes
    For Each shp In allShapes
        If shp.Tags(TAG_GRAPH_NAME) = graphName Then
            Set FindGraphByName = shp
            Exit Function
        End If
    Next shp
    
    Set FindGraphByName = Nothing
End Function

Public Function FindGroupComponents(groupID As String) As Collection
    ' Finds all shapes belonging to a graph group (searches inside PPT groups too)
    Dim components As New Collection
    Dim allShapes As New Collection
    Dim shp As Shape
    
    CollectAllVitalsShapes allShapes
    For Each shp In allShapes
        If shp.Tags(TAG_GRAPH_GROUP) = groupID Then
            components.Add shp
        End If
    Next shp
    
    Set FindGroupComponents = components
End Function

Public Function FindPPTGroupForGraph(graphShape As Shape) As Shape
    ' Returns the PPT group shape that contains the given graph shape,
    ' or Nothing if the graph is not inside a PPT group.
    Dim parent As Object
    Dim sld As slide
    Dim shp As Shape
    Dim gi As Long
    
    ' Get the slide — if shape is inside a group, .Parent is the group shape
    Set parent = graphShape.parent
    If TypeOf parent Is Shape Then
        ' Already inside a group — return the parent group
        Set FindPPTGroupForGraph = parent
        Exit Function
    End If
    
    ' Parent is the slide — search for a group containing this shape
    Set sld = parent
    For Each shp In sld.Shapes
        If shp.Type = msoGroup Then
            For gi = 1 To shp.GroupItems.count
                If shp.GroupItems(gi).Name = graphShape.Name Then
                    Set FindPPTGroupForGraph = shp
                    Exit Function
                End If
            Next gi
        End If
    Next shp
    
    Set FindPPTGroupForGraph = Nothing
End Function

Public Sub UngroupVitalsGroup(graphShape As Shape)
    ' If the graph is inside a PPT group, ungroups it so individual
    ' shapes can be repositioned. Call RegroupVitalsShapes after.
    Dim pptGroup As Shape
    Set pptGroup = FindPPTGroupForGraph(graphShape)
    If Not pptGroup Is Nothing Then
        pptGroup.Ungroup
    End If
End Sub

Public Sub RegroupVitalsShapes(graphShape As Shape)
    ' Regroups all vitals components for a graph into a named PPT group.
    ' Named "VITALS - {graphName}".
    Dim groupID As String
    Dim components As Collection
    Dim comp As Shape
    Dim sld As slide
    Dim shapeNames() As String
    Dim idx As Long
    Dim pptGroup As Shape
    Dim graphName As String
    
    groupID = graphShape.Tags(TAG_GRAPH_GROUP)
    If groupID = "" Then Exit Sub
    
    Set sld = graphShape.parent
    Set components = FindGroupComponents(groupID)
    
    If components.count < 2 Then Exit Sub
    
    ' Collect shape names on this slide
    idx = 0
    ReDim shapeNames(1 To components.count)
    For Each comp In components
        idx = idx + 1
        shapeNames(idx) = comp.Name
    Next comp
    
    ' Group them
    Dim rng As ShapeRange
    Set rng = sld.Shapes.range(shapeNames)
    Set pptGroup = rng.Group
    
    graphName = graphShape.Tags(TAG_GRAPH_NAME)
    pptGroup.Name = "VITALS - " & graphName
End Sub

' ============================================================================
' COORDINATE CONVERSION
' ============================================================================

Public Sub GetPlotAreaSlideCoords(graphShape As Shape, ByRef pLeft As Double, ByRef pTop As Double, _
                                   ByRef pWidth As Double, ByRef pHeight As Double)
    ' Calculates the plot area's INNER bounds (data region only) in slide coordinates.
    ' PlotArea properties are in the graph's internal coordinate system, which may be
    ' scaled differently from slide points. We apply a scaling factor using
    ' ChartArea dimensions vs the shape's actual dimensions on the slide.
    '
    ' Coordinate chain:  Shape edge -> ChartArea offset -> PlotArea offset -> data region
    ' When a graph has a drop shadow or other effects, the shape bounding box may differ
    ' from the ChartArea. ChartArea.Left/Top gives the offset from the shape edge to
    ' where the actual graph content begins.
    Dim cht As Chart
    Dim scaleX As Double, scaleY As Double
    
    Set cht = graphShape.Chart
    
    ' Scaling factors: graph internal coords -> slide points
    scaleX = graphShape.Width / cht.ChartArea.Width
    scaleY = graphShape.Height / cht.ChartArea.Height
    
    ' Full offset = ChartArea offset (shadow/effect padding) + PlotArea inside offset
    ' Both are in graph-internal coords so both get scaled
    pLeft = graphShape.Left + ((cht.ChartArea.Left + cht.PlotArea.InsideLeft) * scaleX)
    pTop = graphShape.Top + ((cht.ChartArea.Top + cht.PlotArea.InsideTop) * scaleY)
    pWidth = cht.PlotArea.InsideWidth * scaleX
    pHeight = cht.PlotArea.InsideHeight * scaleY
End Sub

Public Function GraphYToSlideY(graphShape As Shape, yValue As Double) As Double
    ' Converts a graph Y-axis value to slide Y coordinate (in points)
    Dim pLeft As Double, pTop As Double, pWidth As Double, pHeight As Double
    Dim yMin As Double, yMax As Double
    Dim relativePosition As Double
    
    If Not graphShape.HasChart Then
        GraphYToSlideY = 0
        Exit Function
    End If
    
    ' Get scaled plot area bounds in slide coordinates
    GetPlotAreaSlideCoords graphShape, pLeft, pTop, pWidth, pHeight
    
    yMin = graphShape.Chart.Axes(xlValue).MinimumScale
    yMax = graphShape.Chart.Axes(xlValue).MaximumScale
    
    ' Calculate relative position (0 = top of plot, 1 = bottom)
    ' Y-axis is inverted in screen coordinates (higher Y value = lower on screen)
    relativePosition = (yMax - yValue) / (yMax - yMin)
    
    ' Convert to slide coordinates
    GraphYToSlideY = pTop + (relativePosition * pHeight)
End Function

Public Function GraphXToSlideX(graphShape As Shape, xValue As Double) As Double
    ' Converts a graph X-axis value to slide X coordinate (in points)
    Dim pLeft As Double, pTop As Double, pWidth As Double, pHeight As Double
    Dim xMin As Double, xMax As Double
    Dim relativePosition As Double
    
    If Not graphShape.HasChart Then
        GraphXToSlideX = 0
        Exit Function
    End If
    
    ' Get scaled plot area bounds in slide coordinates
    GetPlotAreaSlideCoords graphShape, pLeft, pTop, pWidth, pHeight
    
    xMin = graphShape.Chart.Axes(xlCategory).MinimumScale
    xMax = graphShape.Chart.Axes(xlCategory).MaximumScale
    
    ' Calculate relative position (0 = left, 1 = right)
    relativePosition = (xValue - xMin) / (xMax - xMin)
    
    ' Convert to slide coordinates
    GraphXToSlideX = pLeft + (relativePosition * pWidth)
End Function

' ============================================================================
' NORMAL RANGE MANAGEMENT
' ============================================================================

Public Sub AddNormalRangeToGraph(graphShape As Shape, yLow As Double, yHigh As Double)
    ' Adds or updates the normal range rectangle for a graph
    Dim normalShape As Shape
    Dim groupID As String
    Dim components As Collection
    Dim comp As Shape
    
    ' Ungroup so we can modify shapes
    UngroupVitalsGroup graphShape
    
    groupID = graphShape.Tags(TAG_GRAPH_GROUP)
    Set components = FindGroupComponents(groupID)
    
    ' Find existing normal range shape
    For Each comp In components
        If comp.Tags(TAG_COMPONENT_TYPE) = "NORMALRANGE" Then
            Set normalShape = comp
            Exit For
        End If
    Next comp
    
    If normalShape Is Nothing Then
        ' Create new normal range shape if not found
        Set normalShape = CreateNormalRangeShape(graphShape, yLow, yHigh)
    Else
        ' Update existing shape
        normalShape.Tags.Delete TAG_NORMAL_YLOW
        normalShape.Tags.Delete TAG_NORMAL_YHIGH
        normalShape.Tags.Add TAG_NORMAL_YLOW, CStr(yLow)
        normalShape.Tags.Add TAG_NORMAL_YHIGH, CStr(yHigh)
        normalShape.Visible = msoTrue
        RepositionNormalRange normalShape, graphShape
    End If
    
    ' Regroup
    RegroupVitalsShapes graphShape
End Sub

Private Function CreateNormalRangeShape(graphShape As Shape, yLow As Double, yHigh As Double) As Shape
    ' Creates a new normal range rectangle shape
    Dim sld As slide
    Dim rect As Shape
    Dim topY As Double, bottomY As Double
    Dim plotLeft As Double, plotTop As Double, plotWidth As Double, plotHeight As Double
    Dim graphName As String
    Dim groupID As String
    
    Set sld = graphShape.parent
    
    graphName = graphShape.Tags(TAG_GRAPH_NAME)
    groupID = graphShape.Tags(TAG_GRAPH_GROUP)
    
    ' Get properly scaled plot area bounds in slide coordinates
    GetPlotAreaSlideCoords graphShape, plotLeft, plotTop, plotWidth, plotHeight
    topY = GraphYToSlideY(graphShape, yHigh)
    bottomY = GraphYToSlideY(graphShape, yLow)
    
    ' Create rectangle
    Set rect = sld.Shapes.AddShape(msoShapeRectangle, _
        plotLeft, topY, plotWidth, bottomY - topY)
    
    ' Style it (green, 70% opacity = 30% transparency)
    With rect
        .Name = graphName & "_NormalRange"
        .Fill.Solid
        .Fill.ForeColor.RGB = RGB(144, 238, 144)  ' Light green
        .Fill.Transparency = 0.3  ' 70% opacity
        .line.Visible = msoFalse
        
        ' Tag it
        .Tags.Add TAG_GRAPH_GROUP, groupID
        .Tags.Add TAG_COMPONENT_TYPE, "NORMALRANGE"
        .Tags.Add TAG_NORMAL_RANGE, graphName
        .Tags.Add TAG_NORMAL_YLOW, CStr(yLow)
        .Tags.Add TAG_NORMAL_YHIGH, CStr(yHigh)
    End With
    
    ' Position in Z-order (behind graph, in front of background)
    rect.ZOrder msoSendToBack
    rect.ZOrder msoBringForward
    rect.ZOrder msoBringForward
    
    Set CreateNormalRangeShape = rect
End Function

Public Sub RepositionNormalRange(normalShape As Shape, graphShape As Shape)
    ' Repositions a normal range rectangle based on stored Y values
    Dim yLow As Double, yHigh As Double
    Dim topY As Double, bottomY As Double
    Dim plotLeft As Double, plotTop As Double, plotWidth As Double, plotHeight As Double
    
    If Not graphShape.HasChart Then Exit Sub
    
    ' Get stored Y values
    On Error Resume Next
    yLow = CDbl(normalShape.Tags(TAG_NORMAL_YLOW))
    yHigh = CDbl(normalShape.Tags(TAG_NORMAL_YHIGH))
    On Error GoTo 0
    
    If yLow = 0 And yHigh = 0 Then Exit Sub  ' Not configured
    
    ' Get properly scaled plot area bounds in slide coordinates
    GetPlotAreaSlideCoords graphShape, plotLeft, plotTop, plotWidth, plotHeight
    topY = GraphYToSlideY(graphShape, yHigh)
    bottomY = GraphYToSlideY(graphShape, yLow)
    
    ' Update shape position
    With normalShape
        .Left = plotLeft
        .Top = topY
        .Width = plotWidth
        .Height = bottomY - topY
    End With
End Sub

Public Sub UpdateLinkedNormalRanges(graphShape As Shape)
    ' Updates all normal range rectangles linked to a specific graph
    Dim graphName As String
    Dim allShapes As New Collection
    Dim shp As Shape
    
    graphName = graphShape.Tags(TAG_GRAPH_NAME)
    If graphName = "" Then Exit Sub
    
    CollectAllVitalsShapes allShapes
    For Each shp In allShapes
        If shp.Tags(TAG_NORMAL_RANGE) = graphName Then
            RepositionNormalRange shp, graphShape
        End If
    Next shp
End Sub

' ============================================================================
' SNAP GROUP TO GRAPH
' ============================================================================

Public Sub SnapGroupToGraph(graphShape As Shape)
    ' Repositions all group components (BG, Tab, Normal) relative to the graph.
    ' Uses stored relative snap tags from the template for style-aware positioning.
    ' Ungroups first if inside a PPT group, then always regroups after.
    Dim groupID As String
    Dim components As Collection
    Dim comp As Shape
    Dim relLeft As Double, relTop As Double
    Dim relWidth As Double, relHeight As Double
    Dim snapRot As Double
    
    ' Ungroup if needed so individual shapes can be repositioned
    UngroupVitalsGroup graphShape
    
    groupID = graphShape.Tags(TAG_GRAPH_GROUP)
    If groupID = "" Then Exit Sub
    
    Set components = FindGroupComponents(groupID)
    
    ' Get the plot area's inner data region in slide coordinates
    Dim plotLeft As Double, plotTop As Double
    Dim plotWidth As Double, plotHeight As Double
    
    GetPlotAreaSlideCoords graphShape, plotLeft, plotTop, plotWidth, plotHeight
    
    ' First pass: position BG and TAB so their bounds are final
    Dim bgComp As Shape
    For Each comp In components
        Select Case comp.Tags(TAG_COMPONENT_TYPE)
            Case "BACKGROUND"
                Set bgComp = comp
                If comp.Tags(TAG_SNAP_REL_LEFT) <> "" Then
                    ' BG scales fully with plot area
                    relLeft = CDbl(comp.Tags(TAG_SNAP_REL_LEFT))
                    relTop = CDbl(comp.Tags(TAG_SNAP_REL_TOP))
                    relWidth = CDbl(comp.Tags(TAG_SNAP_REL_WIDTH))
                    relHeight = CDbl(comp.Tags(TAG_SNAP_REL_HEIGHT))
                    
                    comp.Width = relWidth * plotWidth
                    comp.Height = relHeight * plotHeight
                    comp.Left = plotLeft + (relLeft * plotWidth)
                    comp.Top = plotTop + (relTop * plotHeight)
                Else
                    comp.Left = plotLeft
                    comp.Top = plotTop
                    comp.Width = plotWidth
                    comp.Height = plotHeight
                End If
                
            Case "TAB"
                ' TAB: position and rotation always adjust.
                ' Sizing depends on orientation:
                '   - Rotated sideways (90/270): Width is visual height,
                '     so scale it to match plotHeight. Height (thickness) stays fixed.
                '   - Not rotated (0/180): keep template dimensions.
                '
                ' For rotated shapes, PPT's .Left/.Top refer to the UNROTATED
                ' frame; the shape renders rotated around its center. We must
                ' solve for .Left/.Top that produce the desired VISUAL position.
                '   Visual left  = Left + Width/2 - Height/2
                '   Visual top   = Top  + Height/2 - Width/2
                If comp.Tags(TAG_SNAP_REL_LEFT) <> "" Then
                    relLeft = CDbl(comp.Tags(TAG_SNAP_REL_LEFT))
                    relTop = CDbl(comp.Tags(TAG_SNAP_REL_TOP))
                    snapRot = CDbl(comp.Tags(TAG_SNAP_ROTATION))
                    
                    comp.Rotation = snapRot
                    
                    ' Check if tab is rotated sideways
                    If (snapRot > 45 And snapRot < 135) Or _
                       (snapRot > 225 And snapRot < 315) Then
                        ' Rotated sideways: Width = visual height ? match plot height
                        comp.Width = plotHeight
                        
                        ' Anchor to the correct edge based on template position
                        If relLeft > 0.5 Then
                            ' Tab on the RIGHT side of the plot area
                            ' Desired visual left = plotLeft + plotWidth (flush right)
                            ' Desired visual top  = plotTop (aligned with top)
                            comp.Left = (plotLeft + plotWidth) - comp.Width / 2 + comp.Height / 2
                            comp.Top = plotTop - comp.Height / 2 + comp.Width / 2
                        Else
                            ' Tab on the LEFT side of the plot area
                            ' Desired visual right = plotLeft (flush left)
                            comp.Left = plotLeft - comp.Width / 2 - comp.Height / 2
                            comp.Top = plotTop - comp.Height / 2 + comp.Width / 2
                        End If
                    Else
                        ' Not rotated: keep dimensions, position relative to plot
                        comp.Left = plotLeft + (relLeft * plotWidth)
                        If relTop < 0 Then
                            comp.Top = plotTop - comp.Height
                        Else
                            comp.Top = plotTop + (relTop * plotHeight)
                        End If
                    End If
                Else
                    comp.Left = plotLeft + (plotWidth * 0.05)
                    comp.Top = plotTop - comp.Height
                End If
        End Select
    Next comp
    
    ' Second pass: position NormalRange (needs BG already positioned)
    For Each comp In components
        If comp.Tags(TAG_COMPONENT_TYPE) = "NORMALRANGE" Then
            RepositionNormalRange comp, graphShape
            ' If not configured, center on BG so it doesn't skew group bounds
            Dim nYLow As Double, nYHigh As Double
            On Error Resume Next
            nYLow = CDbl(comp.Tags(TAG_NORMAL_YLOW))
            nYHigh = CDbl(comp.Tags(TAG_NORMAL_YHIGH))
            On Error GoTo 0
            If nYLow = 0 And nYHigh = 0 Then
                ' Shrink to a small size and center on BG so it can't
                ' extend beyond the group bounding box
                If Not bgComp Is Nothing Then
                    comp.Width = 10
                    comp.Height = 10
                    comp.Left = bgComp.Left + (bgComp.Width - comp.Width) / 2
                    comp.Top = bgComp.Top + (bgComp.Height - comp.Height) / 2
                Else
                    comp.Width = 10
                    comp.Height = 10
                    comp.Left = graphShape.Left + (graphShape.Width - comp.Width) / 2
                    comp.Top = graphShape.Top + (graphShape.Height - comp.Height) / 2
                End If
            End If
        End If
    Next comp
    
    ' Always regroup after snapping
    RegroupVitalsShapes graphShape
End Sub

' ============================================================================
' SHOW/HIDE EXTRAS
' ============================================================================

Public Sub HideVitalsExtras()
    ' Hides all vitals-related shapes EXCEPT the graph shapes themselves.
    ' This hides backgrounds, tabs, and normal range shapes.
    Dim allShapes As New Collection
    Dim shp As Shape
    Dim compType As String
    Dim count As Long
    
    CollectAllVitalsShapes allShapes
    For Each shp In allShapes
        If shp.Tags(TAG_GRAPH_GROUP) <> "" Then
            compType = shp.Tags(TAG_COMPONENT_TYPE)
            Select Case compType
                Case "BACKGROUND", "TAB", "NORMALRANGE"
                    shp.Visible = msoFalse
                    count = count + 1
            End Select
        End If
    Next shp
End Sub

Public Sub ShowVitalsExtras()
    ' Shows all vitals-related shapes EXCEPT normal ranges that were hidden by the user.
    ' This re-shows backgrounds and tabs.
    Dim allShapes As New Collection
    Dim shp As Shape
    Dim compType As String
    
    CollectAllVitalsShapes allShapes
    For Each shp In allShapes
        If shp.Tags(TAG_GRAPH_GROUP) <> "" Then
            compType = shp.Tags(TAG_COMPONENT_TYPE)
            Select Case compType
                Case "BACKGROUND", "TAB"
                    shp.Visible = msoTrue
                Case "NORMALRANGE"
                    Dim yLow As Double, yHigh As Double
                    On Error Resume Next
                    yLow = CDbl(shp.Tags(TAG_NORMAL_YLOW))
                    yHigh = CDbl(shp.Tags(TAG_NORMAL_YHIGH))
                    On Error GoTo 0
                    If yLow <> 0 Or yHigh <> 0 Then
                        shp.Visible = msoTrue
                    End If
            End Select
        End If
    Next shp
End Sub

' ============================================================================
' STYLE SWAP
' ============================================================================

Public Sub SwapGraphStyle(graphName As String)
    ' Cycles the graph to the next available style template,
    ' preserving position, bounds, chart data, and normal range settings.
    Dim graphShape As Shape
    Dim currentStyle As Long, nextStyle As Long
    Dim groupID As String
    Dim components As Collection
    Dim comp As Shape
    Dim templatePres As presentation
    Dim templatePath As String
    Dim savedLeft As Double, savedTop As Double
    Dim savedWidth As Double, savedHeight As Double
    Dim savedXMin As Double, savedXMax As Double
    Dim savedYMin As Double, savedYMax As Double
    Dim savedYAuto As Boolean
    Dim savedVitalName As String
    Dim savedNormalYLow As Double, savedNormalYHigh As Double
    Dim savedNormalVisible As Boolean, savedNormalConfigured As Boolean
    Dim targetSlide As slide
    Dim shapesToDelete As New Collection
    Dim seriesCount As Long, s As Long
    Dim seriesNames() As String
    Dim seriesXValues() As Variant
    Dim seriesYValues() As Variant
    Dim newSeries As Object
    
    On Error GoTo SwapError
    
    ' Find the graph
    Set graphShape = FindGraphByName(graphName)
    If graphShape Is Nothing Then
        MsgBox "Graph '" & graphName & "' not found.", vbExclamation, "Swap Style"
        Exit Sub
    End If
    
    ' Get current style and find next available
    currentStyle = 1
    If graphShape.Tags(TAG_STYLE_NUMBER) <> "" Then
        currentStyle = CLng(graphShape.Tags(TAG_STYLE_NUMBER))
    End If
    
    nextStyle = GetNextStyleNumber(currentStyle)
    If nextStyle = 0 Then
        MsgBox "Only one style template found." & vbCrLf & _
               "Add more Scatterplot#.pptx files to:" & vbCrLf & _
               Environ("APPDATA") & TEMPLATE_FOLDER, _
               vbInformation, "Swap Style"
        Exit Sub
    End If
    
    ' === Save current settings ===
    savedLeft = graphShape.Left
    savedTop = graphShape.Top
    savedWidth = graphShape.Width
    savedHeight = graphShape.Height
    
    ' Save axis bounds
    If graphShape.HasChart Then
        On Error Resume Next
        savedXMin = graphShape.Chart.Axes(xlCategory).MinimumScale
        savedXMax = graphShape.Chart.Axes(xlCategory).MaximumScale
        savedYAuto = graphShape.Chart.Axes(xlValue).MinimumScaleIsAuto
        savedYMin = graphShape.Chart.Axes(xlValue).MinimumScale
        savedYMax = graphShape.Chart.Axes(xlValue).MaximumScale
        On Error GoTo SwapError
    End If
    
    ' Save chart data (series)
    seriesCount = 0
    If graphShape.HasChart Then
        On Error Resume Next
        seriesCount = graphShape.Chart.SeriesCollection.count
        On Error GoTo SwapError
        
        If seriesCount > 0 Then
            ReDim seriesNames(1 To seriesCount)
            ReDim seriesXValues(1 To seriesCount)
            ReDim seriesYValues(1 To seriesCount)
            
            For s = 1 To seriesCount
                On Error Resume Next
                seriesNames(s) = graphShape.Chart.SeriesCollection(s).Name
                seriesXValues(s) = graphShape.Chart.SeriesCollection(s).XValues
                seriesYValues(s) = graphShape.Chart.SeriesCollection(s).Values
                On Error GoTo SwapError
            Next s
        End If
    End If
    
    ' Save vital name and normal range from group components
    groupID = graphShape.Tags(TAG_GRAPH_GROUP)
    Set components = FindGroupComponents(groupID)
    savedVitalName = graphName
    savedNormalConfigured = False
    
    For Each comp In components
        Select Case comp.Tags(TAG_COMPONENT_TYPE)
            Case "TAB"
                If comp.HasTextFrame Then
                    savedVitalName = comp.TextFrame.TextRange.text
                End If
            Case "NORMALRANGE"
                On Error Resume Next
                savedNormalYLow = CDbl(comp.Tags(TAG_NORMAL_YLOW))
                savedNormalYHigh = CDbl(comp.Tags(TAG_NORMAL_YHIGH))
                savedNormalVisible = (comp.Visible = msoTrue)
                savedNormalConfigured = (savedNormalYLow <> 0 Or savedNormalYHigh <> 0)
                On Error GoTo SwapError
        End Select
    Next comp
    
    ' Save target slide (check if inside a PPT group)
    Dim pptGroup As Shape
    Set pptGroup = FindPPTGroupForGraph(graphShape)
    If Not pptGroup Is Nothing Then
        Set targetSlide = pptGroup.parent
    Else
        Set targetSlide = graphShape.parent
    End If
    
    ' === Delete all current group shapes ===
    ' If inside a PPT group, delete the whole PPT group shape
    If Not pptGroup Is Nothing Then
        pptGroup.Delete
    Else
        Set components = FindGroupComponents(groupID)
        For Each comp In components
            shapesToDelete.Add comp
        Next comp
        For Each comp In shapesToDelete
            comp.Delete
        Next comp
    End If
    
    ' === Create from new template ===
    templatePath = Environ("APPDATA") & TEMPLATE_FOLDER & "Scatterplot" & nextStyle & ".pptx"
    
    Set templatePres = Application.Presentations.Open( _
        fileName:=templatePath, _
        ReadOnly:=msoTrue, _
        WithWindow:=msoFalse)
    
    CopyAndConfigureTemplate templatePres, targetSlide, graphName, savedVitalName, _
                             savedXMin, savedXMax, nextStyle
    
    templatePres.Saved = True
    templatePres.Close
    Set templatePres = Nothing
    
    ' === Restore settings on the new graph ===
    Set graphShape = FindGraphByName(graphName)
    If graphShape Is Nothing Then
        MsgBox "Error: newly created graph not found.", vbCritical, "Swap Error"
        Exit Sub
    End If
    
    ' Ungroup so we can modify individual shape properties
    UngroupVitalsGroup graphShape
    
    ' Restore position and size
    graphShape.Left = savedLeft
    graphShape.Top = savedTop
    graphShape.Width = savedWidth
    graphShape.Height = savedHeight
    
    ' Restore Y bounds if manually set
    If Not savedYAuto And graphShape.HasChart Then
        On Error Resume Next
        With graphShape.Chart.Axes(xlValue)
            .MinimumScaleIsAuto = False
            .MaximumScaleIsAuto = False
            .MinimumScale = savedYMin
            .MaximumScale = savedYMax
        End With
        On Error GoTo SwapError
    End If
    
    ' Restore chart data — update existing template series to preserve formatting,
    ' only add/remove series if counts differ
    If seriesCount > 0 And graphShape.HasChart Then
        Dim templateSeriesCount As Long
        On Error Resume Next
        templateSeriesCount = graphShape.Chart.SeriesCollection.count
        On Error GoTo SwapError
        
        ' Remove extra template series if template has more than saved data
        On Error Resume Next
        Do While graphShape.Chart.SeriesCollection.count > seriesCount
            graphShape.Chart.SeriesCollection(graphShape.Chart.SeriesCollection.count).Delete
        Loop
        On Error GoTo SwapError
        
        ' Update existing series (preserves template marker/line formatting)
        Dim existingCount As Long
        On Error Resume Next
        existingCount = graphShape.Chart.SeriesCollection.count
        On Error GoTo SwapError
        
        For s = 1 To existingCount
            On Error Resume Next
            graphShape.Chart.SeriesCollection(s).Name = seriesNames(s)
            graphShape.Chart.SeriesCollection(s).XValues = seriesXValues(s)
            graphShape.Chart.SeriesCollection(s).Values = seriesYValues(s)
            On Error GoTo SwapError
        Next s
        
        ' Add new series if saved data has more than template provided
        If seriesCount > existingCount Then
            For s = existingCount + 1 To seriesCount
                On Error Resume Next
                Set newSeries = graphShape.Chart.SeriesCollection.newSeries
                newSeries.Name = seriesNames(s)
                newSeries.XValues = seriesXValues(s)
                newSeries.Values = seriesYValues(s)
                On Error GoTo SwapError
            Next s
        End If
    End If
    
    ' Snap components to restored graph position
    SnapGroupToGraph graphShape
    
    ' Restore normal range
    If savedNormalConfigured Then
        AddNormalRangeToGraph graphShape, savedNormalYLow, savedNormalYHigh
        If Not savedNormalVisible Then
            groupID = graphShape.Tags(TAG_GRAPH_GROUP)
            Set components = FindGroupComponents(groupID)
            For Each comp In components
                If comp.Tags(TAG_COMPONENT_TYPE) = "NORMALRANGE" Then
                    comp.Visible = msoFalse
                    Exit For
                End If
            Next comp
        End If
    End If
    
    MsgBox "Swapped to Style " & nextStyle & " of " & CountAvailableStyles() & ".", _
           vbInformation, "Style Swapped"
    Exit Sub
    
SwapError:
    If Not templatePres Is Nothing Then
        templatePres.Saved = True
        templatePres.Close
    End If
    MsgBox "Error swapping style: " & Err.Description, vbCritical, "Swap Error"
End Sub

Public Function GetNextStyleNumber(currentStyle As Long) As Long
    ' Returns the next available style number (cycling), or 0 if only one exists.
    Dim templateFolder As String
    Dim fileName As String
    Dim numStr As String
    Dim styleNum As Long
    Dim nums() As Long
    Dim numCount As Long
    Dim i As Long, j As Long
    Dim temp As Long
    
    templateFolder = Environ("APPDATA") & TEMPLATE_FOLDER
    numCount = 0
    
    ' Collect all style numbers from Scatterplot#.pptx files
    fileName = Dir(templateFolder & "Scatterplot*.pptx")
    Do While fileName <> ""
        numStr = Mid(fileName, Len("Scatterplot") + 1)
        numStr = Left(numStr, Len(numStr) - Len(".pptx"))
        On Error Resume Next
        styleNum = CLng(numStr)
        If Err.Number = 0 And styleNum > 0 Then
            numCount = numCount + 1
            ReDim Preserve nums(1 To numCount)
            nums(numCount) = styleNum
        End If
        Err.Clear
        On Error GoTo 0
        fileName = Dir()
    Loop
    
    If numCount <= 1 Then
        GetNextStyleNumber = 0
        Exit Function
    End If
    
    ' Sort ascending (simple bubble sort for small array)
    For i = 1 To numCount - 1
        For j = i + 1 To numCount
            If nums(j) < nums(i) Then
                temp = nums(i)
                nums(i) = nums(j)
                nums(j) = temp
            End If
        Next j
    Next i
    
    ' Find current style and return the next one (cycling)
    For i = 1 To numCount
        If nums(i) = currentStyle Then
            If i < numCount Then
                GetNextStyleNumber = nums(i + 1)
            Else
                GetNextStyleNumber = nums(1)
            End If
            Exit Function
        End If
    Next i
    
    ' Current style not found, return first available
    GetNextStyleNumber = nums(1)
End Function

Public Function CountAvailableStyles() As Long
    ' Counts Scatterplot#.pptx files in the template folder
    Dim templateFolder As String
    Dim fileName As String
    Dim count As Long
    
    templateFolder = Environ("APPDATA") & TEMPLATE_FOLDER
    count = 0
    
    fileName = Dir(templateFolder & "Scatterplot*.pptx")
    Do While fileName <> ""
        count = count + 1
        fileName = Dir()
    Loop
    
    CountAvailableStyles = count
End Function

' ============================================================================
' PARSING & FORMATTING
' ============================================================================

Public Function ParseDateTimeBounds(boundsStr As String, ByRef minVal As Double, ByRef maxVal As Double) As Boolean
    ' Parses a date/time bounds string and converts to Excel serial date numbers
    ' Supported formats:
    '   A. mm/dd/yyyy hh:mm:ss - mm/dd/yyyy hh:mm:ss
    '   B. mm/dd/yyyy hh:mm - mm/dd/yyyy hh:mm
    '   C. mm/dd/yyyy - mm/dd/yyyy
    '   D. hh:mm:ss - hh:mm:ss
    '   E. hh:mm - hh:mm
    
    Dim startStr As String, endStr As String
    Dim separatorPos As Long
    
    On Error GoTo ParseError
    
    ' Find the separator " - " (space-dash-space)
    separatorPos = InStr(boundsStr, " - ")
    
    If separatorPos = 0 Then
        ParseDateTimeBounds = False
        Exit Function
    End If
    
    startStr = Trim(Left(boundsStr, separatorPos - 1))
    endStr = Trim(Mid(boundsStr, separatorPos + 3))
    
    ' Try to parse as date/time - VBA's CDate handles multiple formats
    minVal = CDate(startStr)
    maxVal = CDate(endStr)
    
    If minVal >= maxVal Then
        ParseDateTimeBounds = False
        Exit Function
    End If
    
    ParseDateTimeBounds = True
    Exit Function
    
ParseError:
    ParseDateTimeBounds = False
End Function

Public Function GetGraphBounds(graphShape As Shape) As String
    ' Returns a formatted string of the graph's current axis bounds
    Dim cht As Chart
    Dim xMin As Double, xMax As Double, yMin As Double, yMax As Double
    
    If Not graphShape.HasChart Then
        GetGraphBounds = "N/A"
        Exit Function
    End If
    
    Set cht = graphShape.Chart
    
    On Error Resume Next
    xMin = cht.Axes(xlCategory).MinimumScale
    xMax = cht.Axes(xlCategory).MaximumScale
    yMin = cht.Axes(xlValue).MinimumScale
    yMax = cht.Axes(xlValue).MaximumScale
    On Error GoTo 0
    
    ' Format X bounds as date/time if they look like serial dates (> 1)
    If xMin > 1 And xMax > 1 Then
        GetGraphBounds = Format(xMin, "mm/dd/yyyy hh:mm") & " - " & Format(xMax, "mm/dd/yyyy hh:mm")
    Else
        GetGraphBounds = "X: " & xMin & "-" & xMax
    End If
End Function

Public Function GetSlideIndexForShape(shp As Shape) As Long
    ' Returns the slide index for a shape (handles shapes inside PPT groups)
    Dim parent As Object
    Set parent = shp.parent
    ' If parent is a Group shape, go up one more level to get the Slide
    If TypeOf parent Is Shape Then
        Set parent = parent.parent
    End If
    GetSlideIndexForShape = parent.SlideIndex
End Function

' ============================================================================
' MISCELLANEOUS OPTIONS
' ============================================================================

Public Sub MiscSetAxisUnits(graphName As String)
    ' Sets major and/or minor units on X and/or Y axes
    Dim graphShape As Shape
    Dim cht As Chart
    Dim choice As String
    Dim unitStr As String
    Dim unitVal As Double
    
    Set graphShape = FindGraphByName(graphName)
    If graphShape Is Nothing Then Exit Sub
    If Not graphShape.HasChart Then Exit Sub
    Set cht = graphShape.Chart
    
    choice = InputBox("Which axis units to set?" & vbCrLf & vbCrLf & _
                      "1 = X Major" & vbCrLf & _
                      "2 = X Minor" & vbCrLf & _
                      "3 = Y Major" & vbCrLf & _
                      "4 = Y Minor" & vbCrLf & _
                      "5 = Y Major + Minor", _
                      "Set Axis Units")
    If choice = "" Then Exit Sub
    
    On Error GoTo UnitError
    
    Select Case choice
        Case "1"
            unitStr = InputBox("Enter X major unit value:" & vbCrLf & _
                               "(For date axis, 1 = one day, 0.0417 = one hour)", _
                               "X Major Unit")
            If unitStr = "" Then Exit Sub
            unitVal = CDbl(unitStr)
            cht.Axes(xlCategory).MajorUnit = unitVal
            
        Case "2"
            unitStr = InputBox("Enter X minor unit value:", "X Minor Unit")
            If unitStr = "" Then Exit Sub
            unitVal = CDbl(unitStr)
            cht.Axes(xlCategory).MinorUnit = unitVal
            
        Case "3"
            unitStr = InputBox("Enter Y major unit value:", "Y Major Unit")
            If unitStr = "" Then Exit Sub
            unitVal = CDbl(unitStr)
            cht.Axes(xlValue).MajorUnit = unitVal
            
        Case "4"
            unitStr = InputBox("Enter Y minor unit value:", "Y Minor Unit")
            If unitStr = "" Then Exit Sub
            unitVal = CDbl(unitStr)
            cht.Axes(xlValue).MinorUnit = unitVal
            
        Case "5"
            unitStr = InputBox("Enter Y major unit value:", "Y Major Unit")
            If unitStr = "" Then Exit Sub
            unitVal = CDbl(unitStr)
            cht.Axes(xlValue).MajorUnit = unitVal
            
            unitStr = InputBox("Enter Y minor unit value:", "Y Minor Unit")
            If unitStr = "" Then Exit Sub
            unitVal = CDbl(unitStr)
            cht.Axes(xlValue).MinorUnit = unitVal
            
        Case Else
            MsgBox "Invalid choice.", vbExclamation, "Error"
            Exit Sub
    End Select
    
    SnapGroupToGraph graphShape
    MsgBox "Axis units updated.", vbInformation, "Units Set"
    Exit Sub
    
UnitError:
    MsgBox "Error setting axis units: " & Err.Description, vbCritical, "Error"
End Sub

Public Sub MiscSetYBounds(graphName As String)
    ' Manually set Y-axis min and max
    Dim graphShape As Shape
    Dim boundsStr As String
    Dim parts() As String
    Dim yMin As Double, yMax As Double
    
    Set graphShape = FindGraphByName(graphName)
    If graphShape Is Nothing Then Exit Sub
    If Not graphShape.HasChart Then Exit Sub
    
    boundsStr = InputBox("Enter Y-axis bounds (min - max):" & vbCrLf & _
                         "Example: 40 - 200" & vbCrLf & vbCrLf & _
                         "Enter 'auto' to reset to automatic.", _
                         "Set Y-Axis Bounds")
    If boundsStr = "" Then Exit Sub
    
    If LCase(Trim(boundsStr)) = "auto" Then
        On Error Resume Next
        graphShape.Chart.Axes(xlValue).MinimumScaleIsAuto = True
        graphShape.Chart.Axes(xlValue).MaximumScaleIsAuto = True
        On Error GoTo 0
        SnapGroupToGraph graphShape
        MsgBox "Y-axis set to automatic.", vbInformation, "Y-Axis Auto"
        Exit Sub
    End If
    
    parts = Split(boundsStr, " - ")
    If UBound(parts) <> 1 Then
        MsgBox "Invalid format. Use: min - max", vbExclamation, "Error"
        Exit Sub
    End If
    
    On Error Resume Next
    yMin = CDbl(Trim(parts(0)))
    yMax = CDbl(Trim(parts(1)))
    On Error GoTo 0
    
    If yMin >= yMax Then
        MsgBox "Min must be less than max.", vbExclamation, "Error"
        Exit Sub
    End If
    
    On Error Resume Next
    With graphShape.Chart.Axes(xlValue)
        .MinimumScaleIsAuto = False
        .MaximumScaleIsAuto = False
        .MinimumScale = yMin
        .MaximumScale = yMax
    End With
    On Error GoTo 0
    
    SnapGroupToGraph graphShape
    MsgBox "Y-axis bounds set to " & yMin & " - " & yMax, vbInformation, "Y Bounds Set"
End Sub

Public Sub MiscRenameGraph(graphName As String)
    ' Renames a graph and updates all component names and tags
    Dim graphShape As Shape
    Dim newName As String
    Dim groupID As String
    Dim components As Collection
    Dim comp As Shape
    Dim pptGroup As Shape
    
    Set graphShape = FindGraphByName(graphName)
    If graphShape Is Nothing Then Exit Sub
    
    newName = InputBox("Enter new name for '" & graphName & "':", _
                       "Rename Graph", graphName)
    If newName = "" Or newName = graphName Then Exit Sub
    
    ' Check for duplicate name
    If Not FindGraphByName(newName) Is Nothing Then
        MsgBox "A graph named '" & newName & "' already exists.", vbExclamation, "Duplicate Name"
        Exit Sub
    End If
    
    ' Ungroup if inside a PPT group
    Set pptGroup = FindPPTGroupForGraph(graphShape)
    If Not pptGroup Is Nothing Then
        pptGroup.Ungroup
    End If
    
    groupID = graphShape.Tags(TAG_GRAPH_GROUP)
    Set components = FindGroupComponents(groupID)
    
    For Each comp In components
        Select Case comp.Tags(TAG_COMPONENT_TYPE)
            Case "GRAPH"
                comp.Name = newName & "_Graph"
                comp.Tags.Delete TAG_GRAPH_NAME
                comp.Tags.Add TAG_GRAPH_NAME, newName
            Case "BACKGROUND"
                comp.Name = newName & "_BG"
            Case "NORMALRANGE"
                comp.Name = newName & "_NormalRange"
                comp.Tags.Delete TAG_NORMAL_RANGE
                comp.Tags.Add TAG_NORMAL_RANGE, newName
            Case "TAB"
                comp.Name = newName & "_Tab"
                If comp.HasTextFrame Then
                    comp.TextFrame.TextRange.text = newName
                End If
        End Select
    Next comp
    
    ' Regroup with new name
    RegroupVitalsShapes graphShape
    
    MsgBox "Graph renamed to '" & newName & "'.", vbInformation, "Renamed"
End Sub

Public Sub MiscDeleteGraph(graphName As String)
    ' Deletes a graph and all its group components (or the PPT group)
    Dim graphShape As Shape
    Dim groupID As String
    Dim pptGroup As Shape
    Dim components As Collection
    Dim comp As Shape
    Dim shapesToDelete As New Collection
    
    Set graphShape = FindGraphByName(graphName)
    If graphShape Is Nothing Then Exit Sub
    
    If MsgBox("Delete graph '" & graphName & "' and all its components?" & vbCrLf & _
              "This cannot be undone.", _
              vbYesNo + vbExclamation, "Confirm Delete") = vbNo Then
        Exit Sub
    End If
    
    Set pptGroup = FindPPTGroupForGraph(graphShape)
    If Not pptGroup Is Nothing Then
        pptGroup.Delete
    Else
        groupID = graphShape.Tags(TAG_GRAPH_GROUP)
        Set components = FindGroupComponents(groupID)
        For Each comp In components
            shapesToDelete.Add comp
        Next comp
        For Each comp In shapesToDelete
            comp.Delete
        Next comp
    End If
    
    MsgBox "Graph '" & graphName & "' deleted.", vbInformation, "Deleted"
End Sub

Public Sub MiscMatchSizes(graphNames As Collection)
    ' Makes all selected graphs the same size as the first selected graph
    Dim firstGraph As Shape
    Dim graphShape As Shape
    Dim i As Long
    
    If graphNames.count < 2 Then
        MsgBox "Select at least 2 graphs to match sizes.", vbExclamation, "Match Sizes"
        Exit Sub
    End If
    
    Set firstGraph = FindGraphByName(graphNames(1))
    If firstGraph Is Nothing Then Exit Sub
    
    For i = 2 To graphNames.count
        Set graphShape = FindGraphByName(graphNames(i))
        If Not graphShape Is Nothing Then
            UngroupVitalsGroup graphShape
            graphShape.Width = firstGraph.Width
            graphShape.Height = firstGraph.Height
            SnapGroupToGraph graphShape
        End If
    Next i
    
    MsgBox graphNames.count - 1 & " graph(s) resized to match '" & graphNames(1) & "'.", _
           vbInformation, "Sizes Matched"
End Sub

Public Sub MiscAlignLeft(graphNames As Collection)
    ' Aligns left edges of all selected graphs to the first selected graph
    Dim firstGraph As Shape
    Dim graphShape As Shape
    Dim i As Long
    
    If graphNames.count < 2 Then
        MsgBox "Select at least 2 graphs to align.", vbExclamation, "Align Left"
        Exit Sub
    End If
    
    Set firstGraph = FindGraphByName(graphNames(1))
    If firstGraph Is Nothing Then Exit Sub
    
    For i = 2 To graphNames.count
        Set graphShape = FindGraphByName(graphNames(i))
        If Not graphShape Is Nothing Then
            UngroupVitalsGroup graphShape
            graphShape.Left = firstGraph.Left
            SnapGroupToGraph graphShape
        End If
    Next i
    
    MsgBox graphNames.count - 1 & " graph(s) aligned to '" & graphNames(1) & "'.", _
           vbInformation, "Aligned"
End Sub

Public Sub MiscDuplicateGraph(graphName As String)
    ' Duplicates a graph and all components, offset slightly
    Dim graphShape As Shape
    Dim groupID As String
    Dim components As Collection
    Dim comp As Shape
    Dim newName As String
    Dim newGroupID As String
    Dim dupShape As Shape
    Dim newGraphShape As Shape
    
    Set graphShape = FindGraphByName(graphName)
    If graphShape Is Nothing Then Exit Sub
    
    newName = InputBox("Enter name for the duplicate graph:", _
                       "Duplicate Graph", graphName & " Copy")
    If newName = "" Then Exit Sub
    
    If Not FindGraphByName(newName) Is Nothing Then
        MsgBox "A graph named '" & newName & "' already exists.", vbExclamation, "Duplicate Name"
        Exit Sub
    End If
    
    ' Ungroup if inside a PPT group so we can access individual shapes
    UngroupVitalsGroup graphShape
    
    groupID = graphShape.Tags(TAG_GRAPH_GROUP)
    newGroupID = newName & "_" & Format(Now, "yyyymmddhhnnss")
    Set components = FindGroupComponents(groupID)
    
    Dim sld As slide
    Set sld = graphShape.parent
    
    For Each comp In components
        comp.Copy
        Set dupShape = sld.Shapes.Paste(1)
        dupShape.Left = comp.Left + 20
        dupShape.Top = comp.Top + 20
        
        ' Transfer all tags with updated group/name references
        dupShape.Tags.Add TAG_GRAPH_GROUP, newGroupID
        
        Select Case comp.Tags(TAG_COMPONENT_TYPE)
            Case "GRAPH"
                dupShape.Name = newName & "_Graph"
                dupShape.Tags.Add TAG_VITALS_GRAPH, "TRUE"
                dupShape.Tags.Add TAG_GRAPH_NAME, newName
                dupShape.Tags.Add TAG_COMPONENT_TYPE, "GRAPH"
                If comp.Tags(TAG_STYLE_NUMBER) <> "" Then
                    dupShape.Tags.Add TAG_STYLE_NUMBER, comp.Tags(TAG_STYLE_NUMBER)
                End If
                If comp.Tags(TAG_BLOCK_ABOVE_RATIO) <> "" Then
                    dupShape.Tags.Add TAG_BLOCK_ABOVE_RATIO, comp.Tags(TAG_BLOCK_ABOVE_RATIO)
                End If
                Set newGraphShape = dupShape
            Case "BACKGROUND"
                dupShape.Name = newName & "_BG"
                dupShape.Tags.Add TAG_COMPONENT_TYPE, "BACKGROUND"
            Case "NORMALRANGE"
                dupShape.Name = newName & "_NormalRange"
                dupShape.Tags.Add TAG_COMPONENT_TYPE, "NORMALRANGE"
                dupShape.Tags.Add TAG_NORMAL_RANGE, newName
                If comp.Tags(TAG_NORMAL_YLOW) <> "" Then dupShape.Tags.Add TAG_NORMAL_YLOW, comp.Tags(TAG_NORMAL_YLOW)
                If comp.Tags(TAG_NORMAL_YHIGH) <> "" Then dupShape.Tags.Add TAG_NORMAL_YHIGH, comp.Tags(TAG_NORMAL_YHIGH)
            Case "TAB"
                dupShape.Name = newName & "_Tab"
                dupShape.Tags.Add TAG_COMPONENT_TYPE, "TAB"
                If dupShape.HasTextFrame Then dupShape.TextFrame.TextRange.text = newName
        End Select
        
        ' Copy snap tags
        If comp.Tags(TAG_SNAP_REL_LEFT) <> "" Then dupShape.Tags.Add TAG_SNAP_REL_LEFT, comp.Tags(TAG_SNAP_REL_LEFT)
        If comp.Tags(TAG_SNAP_REL_TOP) <> "" Then dupShape.Tags.Add TAG_SNAP_REL_TOP, comp.Tags(TAG_SNAP_REL_TOP)
        If comp.Tags(TAG_SNAP_REL_WIDTH) <> "" Then dupShape.Tags.Add TAG_SNAP_REL_WIDTH, comp.Tags(TAG_SNAP_REL_WIDTH)
        If comp.Tags(TAG_SNAP_REL_HEIGHT) <> "" Then dupShape.Tags.Add TAG_SNAP_REL_HEIGHT, comp.Tags(TAG_SNAP_REL_HEIGHT)
        If comp.Tags(TAG_SNAP_ROTATION) <> "" Then dupShape.Tags.Add TAG_SNAP_ROTATION, comp.Tags(TAG_SNAP_ROTATION)
    Next comp
    
    ' Regroup the original
    RegroupVitalsShapes graphShape
    
    ' Group the duplicate
    If Not newGraphShape Is Nothing Then
        SnapGroupToGraph newGraphShape
    End If
    
    MsgBox "Graph duplicated as '" & newName & "'.", vbInformation, "Duplicated"
End Sub

Public Sub MiscResetBounds(graphName As String)
    ' Reset axis bounds to auto (X, Y, or both)
    Dim graphShape As Shape
    Dim choice As String
    
    Set graphShape = FindGraphByName(graphName)
    If graphShape Is Nothing Then Exit Sub
    If Not graphShape.HasChart Then Exit Sub
    
    choice = InputBox("Reset which axis to auto?" & vbCrLf & vbCrLf & _
                      "1 = X-Axis (horizontal)" & vbCrLf & _
                      "2 = Y-Axis (vertical)" & vbCrLf & _
                      "3 = Both", _
                      "Reset Bounds to Auto")
    If choice = "" Then Exit Sub
    
    On Error Resume Next
    Select Case choice
        Case "1"
            graphShape.Chart.Axes(xlCategory).MinimumScaleIsAuto = True
            graphShape.Chart.Axes(xlCategory).MaximumScaleIsAuto = True
            On Error GoTo 0
            SnapGroupToGraph graphShape
            MsgBox "X-axis bounds reset to auto.", vbInformation, "Bounds Reset"
            
        Case "2"
            graphShape.Chart.Axes(xlValue).MinimumScaleIsAuto = True
            graphShape.Chart.Axes(xlValue).MaximumScaleIsAuto = True
            On Error GoTo 0
            SnapGroupToGraph graphShape
            MsgBox "Y-axis bounds reset to auto.", vbInformation, "Bounds Reset"
            
        Case "3"
            graphShape.Chart.Axes(xlCategory).MinimumScaleIsAuto = True
            graphShape.Chart.Axes(xlCategory).MaximumScaleIsAuto = True
            graphShape.Chart.Axes(xlValue).MinimumScaleIsAuto = True
            graphShape.Chart.Axes(xlValue).MaximumScaleIsAuto = True
            On Error GoTo 0
            SnapGroupToGraph graphShape
            MsgBox "Both axes reset to auto.", vbInformation, "Bounds Reset"
            
        Case Else
            On Error GoTo 0
            MsgBox "Invalid option.", vbExclamation, "Error"
    End Select
End Sub




