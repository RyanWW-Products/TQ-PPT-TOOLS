VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "EZHighlightEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mApp As Application
Attribute mApp.VB_VarHelpID = -1
Private mWindow As DocumentWindow
Private mSlide As Slide
Private mExisting As Object
Private mArmed As Boolean
Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal key As Long) As Integer

Public Sub Arm(ByVal app As Application, ByVal window As DocumentWindow, ByVal sld As Slide)
    Dim shape As Shape, ignored As Integer
    Set mWindow = window: Set mSlide = sld
    Set mExisting = CreateObject("Scripting.Dictionary")
    For Each shape In sld.Shapes
        mExisting.Add CStr(shape.Id), True
    Next
    ignored = GetAsyncKeyState(vbKeyEscape)
    Set mApp = app
    mArmed = True
End Sub

Public Sub Disarm()
    mArmed = False
    Set mApp = Nothing
    Set mWindow = Nothing
    Set mSlide = Nothing
    Set mExisting = Nothing
End Sub

Private Sub mApp_WindowSelectionChange(ByVal Sel As Selection)
    Dim shape As Shape, targetSlide As Slide
    On Error GoTo cancelled
    If Not mArmed Then Exit Sub
    If GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo cancelled
    If Not mApp.ActiveWindow Is mWindow Then GoTo cancelled
    If mWindow.View.Slide.SlideID <> mSlide.SlideID Then GoTo cancelled
    If Sel.Type <> ppSelectionShapes Then Exit Sub
    If Sel.ShapeRange.Count <> 1 Then GoTo cancelled
    Set shape = Sel.ShapeRange(1)
    If mExisting.Exists(CStr(shape.Id)) Then GoTo cancelled
    If shape.Type <> msoAutoShape Then GoTo cancelled
    If shape.AutoShapeType <> msoShapeRectangle Then GoTo cancelled
    Set targetSlide = mSlide
    Disarm
    EZHighlightsDrawn targetSlide, shape
    Exit Sub
cancelled:
    Disarm
End Sub

Private Sub mApp_WindowDeactivate(ByVal Pres As Presentation, ByVal Wn As DocumentWindow)
    Disarm
End Sub

Private Sub mApp_SlideSelectionChanged(ByVal SldRange As SlideRange)
    On Error GoTo cancelled
    If Not mArmed Then Exit Sub
    If SldRange.Count = 1 Then
        If SldRange(1).SlideID = mSlide.SlideID Then Exit Sub
    End If
cancelled:
    Disarm
End Sub

Private Sub mApp_PresentationClose(ByVal Pres As Presentation)
    Disarm
End Sub
