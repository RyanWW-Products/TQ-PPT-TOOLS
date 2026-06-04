Attribute VB_Name = "A_PROJECT_EXPORTER"
Option Explicit

'=============================================================================
' ExportAllFormsAndModules
'
' Exports every UserForm and code module from the current VBA project.
' Prompts for a destination folder, then writes:
'     <chosen folder>\Forms     -> all UserForms (.frm + .frx)
'     <chosen folder>\Modules   -> standard (.bas), class & document (.cls)
'
' REQUIRES: Trust Center > Macro Settings >
'           "Trust access to the VBA project object model"  (must be ON)
'=============================================================================

' Component type IDs (declared locally so no VBIDE reference is required)
Private Const CT_STD_MODULE   As Long = 1     ' Standard module   -> .bas
Private Const CT_CLASS_MODULE As Long = 2     ' Class module      -> .cls
Private Const CT_MSFORM       As Long = 3     ' UserForm          -> .frm
Private Const CT_DOCUMENT     As Long = 100   ' Document module   -> .cls

Public Sub ExportAllFormsAndModules()
    Dim fd As FileDialog
    Dim fso As Object
    Dim proj As Object
    Dim comp As Object
    Dim rootPath As String, formsPath As String, modulesPath As String
    Dim targetFolder As String, ext As String, filePath As String
    Dim formCount As Long, moduleCount As Long, skipped As Long, failed As Long
    Dim failedList As String, msg As String

    ' 1. Grab the project (and check programmatic access is allowed) ----------
    ' Targets the project currently active in the VB Editor.
    ' To always target this presentation instead, swap the line below for:
    '     Set proj = ActivePresentation.VBProject
    On Error Resume Next
    Set proj = Application.VBE.ActiveVBProject
    On Error GoTo 0
    If proj Is Nothing Then
        MsgBox "Could not access the VBA project." & vbCrLf & vbCrLf & _
               "Turn on: File > Options > Trust Center > Trust Center Settings >" & vbCrLf & _
               "Macro Settings > 'Trust access to the VBA project object model'," & vbCrLf & _
               "then run this macro again.", vbCritical, "Export VBA Components"
        Exit Sub
    End If

    ' 2. Ask where to put the export -----------------------------------------
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "Choose where to export this VBA project"
    fd.AllowMultiSelect = False
    If fd.Show <> -1 Then Exit Sub          ' user cancelled
    rootPath = fd.SelectedItems(1)

    ' 3. Create the Forms and Modules sub-folders ----------------------------
    Set fso = CreateObject("Scripting.FileSystemObject")
    formsPath = fso.BuildPath(rootPath, "Forms")
    modulesPath = fso.BuildPath(rootPath, "Modules")
    If Not fso.FolderExists(formsPath) Then fso.CreateFolder formsPath
    If Not fso.FolderExists(modulesPath) Then fso.CreateFolder modulesPath

    ' 4. Export every component ----------------------------------------------
    For Each comp In proj.VBComponents
        ext = ""
        Select Case comp.Type
            Case CT_MSFORM
                ext = ".frm": targetFolder = formsPath
            Case CT_STD_MODULE
                ext = ".bas": targetFolder = modulesPath
            Case CT_CLASS_MODULE, CT_DOCUMENT
                ext = ".cls": targetFolder = modulesPath
        End Select

        If ext = "" Then
            skipped = skipped + 1                 ' e.g. ActiveX designers
        Else
            filePath = fso.BuildPath(targetFolder, comp.Name & ext)
            On Error Resume Next
            Err.Clear
            comp.Export filePath                  ' .frm export also writes .frx
            If Err.Number <> 0 Then
                failed = failed + 1
                failedList = failedList & vbCrLf & "   - " & comp.Name & _
                             " (" & Err.Description & ")"
                Err.Clear
            ElseIf comp.Type = CT_MSFORM Then
                formCount = formCount + 1
            Else
                moduleCount = moduleCount + 1
            End If
            On Error GoTo 0
        End If
    Next comp

    ' 5. Report --------------------------------------------------------------
    msg = "Export complete." & vbCrLf & vbCrLf & _
          "Forms exported:   " & formCount & vbCrLf & _
          "Modules exported: " & moduleCount
    If skipped > 0 Then msg = msg & vbCrLf & "Skipped (other):  " & skipped
    If failed > 0 Then msg = msg & vbCrLf & vbCrLf & _
          "Failed (" & failed & "):" & failedList
    msg = msg & vbCrLf & vbCrLf & "Location:" & vbCrLf & rootPath

    MsgBox msg, IIf(failed > 0, vbExclamation, vbInformation), "Export VBA Components"
End Sub


