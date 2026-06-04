Attribute VB_Name = "Open_QG_Manager"

Public Sub PopulateListBox(ByRef lstBox As MSForms.ListBox, ByVal folderPath As String)
    Dim fileName As String
    lstBox.Clear
    fileName = Dir(folderPath & "*.pptx")
    Do While fileName <> ""
        lstBox.AddItem fileName
        fileName = Dir
    Loop
End Sub

