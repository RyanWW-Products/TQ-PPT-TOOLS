VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TimelineSettings 
   Caption         =   "Timeline Settings"
   ClientHeight    =   4710
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3690
   OleObjectBlob   =   "TimelineSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TimelineSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' At the top of the UserForm code module
Public timelineType As String
Public timelineColor As String


Private Sub UserForm_Initialize()
    ' Set default values
    optHours.Value = True
    optGreen.Value = True
End Sub

Private Sub cmdOK_Click()
    ' Assign values based on selection
    If optHours.Value Then
        timelineType = "Hours"
    ElseIf optDays.Value Then
        timelineType = "Days"
    ElseIf optMonths.Value Then
        timelineType = "Months"
    Else
        timelineType = "Years"
    End If

    If optGreen.Value Then
        timelineColor = "Green"
    Else
        timelineColor = "Grey"
    End If

    Me.Tag = "OK"
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    Me.Tag = "Cancel"
    Me.Hide
End Sub

