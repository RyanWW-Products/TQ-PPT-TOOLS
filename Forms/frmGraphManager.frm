VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGraphManager 
   Caption         =   "Vitals Graph Manager"
   ClientHeight    =   7545
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4455
   OleObjectBlob   =   "frmGraphManager.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmGraphManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' ============================================================================
' frmChartManager - Vitals Scatterplot Manager UserForm Code
' ============================================================================
' Buttons on the form:
'   VitalsSelectionBox  - ListBox (MultiSelect = fmMultiSelectMulti)
'   btnNewGraph         - Create a new vitals graph
'   btnEditBounds       - Edit bounds on selected graph
'   btnAddNormal        - Add/edit normal range on selected graph
'   btnEditNormalRange  - Edit existing normal range values
'   btnToggleNormal     - Toggle normal range visibility
'   btnSnapGroup        - Snap components to graph
'   btnOrientGraph      - Auto-orient selected graphs vertically
'   btnSetMargins       - Set top/bottom margin shapes for orientation
'   btnHideExtras       - Toggle visibility of extras (BG, tab, normal range)
'   btnSwapStyle        - Cycle through available template styles
'   btnMiscOptions      - Misc options (axis units, Y bounds, rename, delete, etc.)
'   btnClose            - Close the form
' ============================================================================

Private Sub UserForm_Initialize()
    RefreshGraphList
End Sub

Private Sub RefreshGraphList()
    Dim graphs As Collection
    Dim shp As Shape
    Dim i As Long
    
    VitalsSelectionBox.Clear
    Set graphs = GetAllVitalsGraphs()
    
    For i = 1 To graphs.count
        Set shp = graphs(i)
        VitalsSelectionBox.AddItem shp.Tags("GRAPHNAME") & " (Slide " & GetSlideIndexForShape(shp) & ")"
    Next i
End Sub

Private Function GetSelectedGraphName() As String
    ' Returns the graph name from the first selected item in the list
    Dim i As Long
    Dim itemText As String
    
    For i = 0 To VitalsSelectionBox.ListCount - 1
        If VitalsSelectionBox.Selected(i) Then
            itemText = VitalsSelectionBox.List(i)
            ' Strip the " (Slide X)" suffix
            GetSelectedGraphName = Left(itemText, InStr(itemText, " (Slide") - 1)
            Exit Function
        End If
    Next i
    
    GetSelectedGraphName = ""
End Function

Private Function GetSelectedGraphNames() As Collection
    ' Returns a collection of all selected graph names
    Dim names As New Collection
    Dim i As Long
    Dim itemText As String
    
    For i = 0 To VitalsSelectionBox.ListCount - 1
        If VitalsSelectionBox.Selected(i) Then
            itemText = VitalsSelectionBox.List(i)
            names.Add Left(itemText, InStr(itemText, " (Slide") - 1)
        End If
    Next i
    
    Set GetSelectedGraphNames = names
End Function

' --- Button Handlers ---

Private Sub btnNewGraph_Click()
    Me.Hide
    CreateVitalsScatterplot
    Me.Show
    RefreshGraphList
End Sub

Private Sub btnEditBounds_Click()
    Dim graphName As String
    Dim graphShape As Shape
    Dim boundsStr As String
    Dim xMin As Double, xMax As Double
    
    graphName = GetSelectedGraphName()
    If graphName = "" Then
        MsgBox "Please select a graph first.", vbExclamation, "No Selection"
        Exit Sub
    End If
    
    Set graphShape = FindGraphByName(graphName)
    If graphShape Is Nothing Then Exit Sub
    
    ' Show current bounds
    MsgBox "Current bounds:" & vbCrLf & vbCrLf & GetGraphBounds(graphShape), _
           vbInformation, "Current Bounds"
    
    ' Get new X bounds
    boundsStr = InputBox("Enter new horizontal time bounds:" & vbCrLf & vbCrLf & _
                         "Format: mm/dd/yyyy hh:mm - mm/dd/yyyy hh:mm" & vbCrLf & vbCrLf & _
                         "Leave empty to keep current bounds.", _
                         "Edit Horizontal Bounds")
    
    If boundsStr <> "" Then
        If ParseDateTimeBounds(boundsStr, xMin, xMax) Then
            ConfigureGraphXBounds graphShape, xMin, xMax
            SnapGroupToGraph graphShape
            MsgBox "Bounds updated successfully.", vbInformation, "Bounds Updated"
        Else
            MsgBox "Invalid format.", vbExclamation, "Error"
        End If
    End If
End Sub

Private Sub btnAddNormal_Click()
    Dim graphName As String
    Dim graphShape As Shape
    Dim rangeStr As String
    Dim parts() As String
    Dim yLow As Double, yHigh As Double
    
    graphName = GetSelectedGraphName()
    If graphName = "" Then
        MsgBox "Please select a graph first.", vbExclamation, "No Selection"
        Exit Sub
    End If
    
    Set graphShape = FindGraphByName(graphName)
    If graphShape Is Nothing Then Exit Sub
    
    rangeStr = InputBox("Enter the normal range (low - high):" & vbCrLf & vbCrLf & _
                        "Example: 60 - 100", _
                        "Normal Range", "60 - 100")
    If rangeStr = "" Then Exit Sub
    
    parts = Split(rangeStr, " - ")
    If UBound(parts) <> 1 Then
        MsgBox "Invalid format. Use: low - high", vbExclamation, "Error"
        Exit Sub
    End If
    
    On Error Resume Next
    yLow = CDbl(Trim(parts(0)))
    yHigh = CDbl(Trim(parts(1)))
    On Error GoTo 0
    
    If yLow >= yHigh Then
        MsgBox "Low value must be less than high value.", vbExclamation, "Error"
        Exit Sub
    End If
    
    AddNormalRangeToGraph graphShape, yLow, yHigh
    MsgBox "Normal range set to " & yLow & " - " & yHigh, vbInformation, "Normal Range Set"
End Sub

Private Sub btnEditNormalRange_Click()
    Dim graphName As String
    Dim graphShape As Shape
    Dim groupID As String
    Dim components As Collection
    Dim comp As Shape
    Dim currentYLow As Double, currentYHigh As Double
    Dim hasNormal As Boolean
    Dim rangeStr As String
    Dim parts() As String
    Dim yLow As Double, yHigh As Double
    Dim defaultStr As String
    
    graphName = GetSelectedGraphName()
    If graphName = "" Then
        MsgBox "Please select a graph first.", vbExclamation, "No Selection"
        Exit Sub
    End If
    
    Set graphShape = FindGraphByName(graphName)
    If graphShape Is Nothing Then Exit Sub
    
    ' Find current normal range values
    groupID = graphShape.Tags("GRAPHGROUP")
    Set components = FindGroupComponents(groupID)
    hasNormal = False
    
    For Each comp In components
        If comp.Tags("COMPONENTTYPE") = "NORMALRANGE" Then
            On Error Resume Next
            currentYLow = CDbl(comp.Tags("NORMALYLOW"))
            currentYHigh = CDbl(comp.Tags("NORMALYHIGH"))
            On Error GoTo 0
            If currentYLow <> 0 Or currentYHigh <> 0 Then
                hasNormal = True
            End If
            Exit For
        End If
    Next comp
    
    If Not hasNormal Then
        MsgBox "No normal range has been set for this graph." & vbCrLf & _
               "Use 'Add Normal' to create one first.", vbExclamation, "No Normal Range"
        Exit Sub
    End If
    
    defaultStr = CStr(currentYLow) & " - " & CStr(currentYHigh)
    
    rangeStr = InputBox("Current normal range: " & defaultStr & vbCrLf & vbCrLf & _
                        "Enter new normal range (low - high):" & vbCrLf & _
                        "Example: 60 - 100", _
                        "Edit Normal Range", defaultStr)
    If rangeStr = "" Then Exit Sub
    
    parts = Split(rangeStr, " - ")
    If UBound(parts) <> 1 Then
        MsgBox "Invalid format. Use: low - high", vbExclamation, "Error"
        Exit Sub
    End If
    
    On Error Resume Next
    yLow = CDbl(Trim(parts(0)))
    yHigh = CDbl(Trim(parts(1)))
    On Error GoTo 0
    
    If yLow >= yHigh Then
        MsgBox "Low value must be less than high value.", vbExclamation, "Error"
        Exit Sub
    End If
    
    AddNormalRangeToGraph graphShape, yLow, yHigh
    MsgBox "Normal range updated to " & yLow & " - " & yHigh, vbInformation, "Normal Range Updated"
End Sub

Private Sub btnToggleNormal_Click()
    Dim graphName As String
    Dim graphShape As Shape
    Dim groupID As String
    Dim components As Collection
    Dim comp As Shape
    
    graphName = GetSelectedGraphName()
    If graphName = "" Then
        MsgBox "Please select a graph first.", vbExclamation, "No Selection"
        Exit Sub
    End If
    
    Set graphShape = FindGraphByName(graphName)
    If graphShape Is Nothing Then Exit Sub
    
    groupID = graphShape.Tags("GRAPHGROUP")
    Set components = FindGroupComponents(groupID)
    
    For Each comp In components
        If comp.Tags("COMPONENTTYPE") = "NORMALRANGE" Then
            If comp.Visible = msoTrue Then
                comp.Visible = msoFalse
            Else
                comp.Visible = msoTrue
            End If
            Exit For
        End If
    Next comp
End Sub

Private Sub btnSnapGroup_Click()
    Dim graphName As String
    Dim graphShape As Shape
    
    graphName = GetSelectedGraphName()
    If graphName = "" Then
        MsgBox "Please select a graph first.", vbExclamation, "No Selection"
        Exit Sub
    End If
    
    Set graphShape = FindGraphByName(graphName)
    If graphShape Is Nothing Then Exit Sub
    
    SnapGroupToGraph graphShape
    MsgBox "Components snapped to graph.", vbInformation, "Snap Complete"
End Sub

Private Sub btnOrientGraph_Click()
    Dim selectedGraphs As Collection
    Set selectedGraphs = GetSelectedGraphNames()
    
    If selectedGraphs.count = 0 Then
        MsgBox "Please select one or more graphs to orient.", vbExclamation, "No Selection"
        Exit Sub
    End If
    
    OrientSelectedGraphs selectedGraphs
End Sub

Private Sub btnSetMargins_Click()
    SetOrientMargins
    RefreshGraphList
End Sub

Private Sub btnSwapStyle_Click()
    Dim graphName As String
    
    graphName = GetSelectedGraphName()
    If graphName = "" Then
        MsgBox "Please select a graph first.", vbExclamation, "No Selection"
        Exit Sub
    End If
    
    SwapGraphStyle graphName
    RefreshGraphList
End Sub

Private Sub btnHideExtras_Click()
    ' Toggle: if any extras are currently visible, hide all; otherwise show all
    Dim graphs As Collection
    Dim graphShape As Shape
    Dim groupID As String
    Dim components As Collection
    Dim comp As Shape
    Dim anyVisible As Boolean
    
    anyVisible = False
    Set graphs = GetAllVitalsGraphs()
    
    Dim i As Long
    For i = 1 To graphs.count
        Set graphShape = graphs(i)
        groupID = graphShape.Tags("GRAPHGROUP")
        Set components = FindGroupComponents(groupID)
        For Each comp In components
            Select Case comp.Tags("COMPONENTTYPE")
                Case "BACKGROUND", "TAB"
                    If comp.Visible = msoTrue Then
                        anyVisible = True
                        Exit For
                    End If
            End Select
        Next comp
        If anyVisible Then Exit For
    Next i
    
    If anyVisible Then
        HideVitalsExtras
        btnHideExtras.Caption = "Show Extras"
    Else
        ShowVitalsExtras
        btnHideExtras.Caption = "Hide Extras"
    End If
End Sub

Private Sub btnMiscOptions_Click()
    Dim choice As String
    Dim graphName As String
    Dim selectedGraphs As Collection
    
    choice = InputBox("Select an option:" & vbCrLf & vbCrLf & _
                      "1 = Set Axis Units (major/minor)" & vbCrLf & _
                      "2 = Set Y-Axis Bounds" & vbCrLf & _
                      "3 = Rename Graph" & vbCrLf & _
                      "4 = Delete Graph" & vbCrLf & _
                      "5 = Duplicate Graph" & vbCrLf & _
                      "6 = Match Sizes (multi-select)" & vbCrLf & _
                      "7 = Align Left Edges (multi-select)" & vbCrLf & _
                      "8 = Reset Bounds to Auto", _
                      "Misc Options")
    If choice = "" Then Exit Sub
    
    Select Case choice
        Case "1"
            graphName = GetSelectedGraphName()
            If graphName = "" Then
                MsgBox "Please select a graph first.", vbExclamation, "No Selection"
                Exit Sub
            End If
            MiscSetAxisUnits graphName
            
        Case "2"
            graphName = GetSelectedGraphName()
            If graphName = "" Then
                MsgBox "Please select a graph first.", vbExclamation, "No Selection"
                Exit Sub
            End If
            MiscSetYBounds graphName
            
        Case "3"
            graphName = GetSelectedGraphName()
            If graphName = "" Then
                MsgBox "Please select a graph first.", vbExclamation, "No Selection"
                Exit Sub
            End If
            MiscRenameGraph graphName
            RefreshGraphList
            
        Case "4"
            graphName = GetSelectedGraphName()
            If graphName = "" Then
                MsgBox "Please select a graph first.", vbExclamation, "No Selection"
                Exit Sub
            End If
            MiscDeleteGraph graphName
            RefreshGraphList
            
        Case "5"
            graphName = GetSelectedGraphName()
            If graphName = "" Then
                MsgBox "Please select a graph first.", vbExclamation, "No Selection"
                Exit Sub
            End If
            MiscDuplicateGraph graphName
            RefreshGraphList
            
        Case "6"
            Set selectedGraphs = GetSelectedGraphNames()
            If selectedGraphs.count < 2 Then
                MsgBox "Select at least 2 graphs to match sizes.", vbExclamation, "No Selection"
                Exit Sub
            End If
            MiscMatchSizes selectedGraphs
            
        Case "7"
            Set selectedGraphs = GetSelectedGraphNames()
            If selectedGraphs.count < 2 Then
                MsgBox "Select at least 2 graphs to align.", vbExclamation, "No Selection"
                Exit Sub
            End If
            MiscAlignLeft selectedGraphs
            
        Case "8"
            graphName = GetSelectedGraphName()
            If graphName = "" Then
                MsgBox "Please select a graph first.", vbExclamation, "No Selection"
                Exit Sub
            End If
            MiscResetBounds graphName
            
        Case Else
            MsgBox "Invalid option.", vbExclamation, "Error"
    End Select
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub


