VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectGraphics 
   Caption         =   "Select QuickGraphics:"
   ClientHeight    =   11265
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6795
   OleObjectBlob   =   "SelectGraphics.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SelectGraphics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const APP_NAME As String = "TrialExAddin"
Private Const CATEGORY_SECTION As String = "SelectGraphicsCategories"
Private Const CATEGORY_KEY As String = "Categories"

' Module-level variable for the base folder.
Private BASE_FOLDER As String
' Collection to hold all ListBoxes that display categories.
Private colCategoryListBoxes As Collection


Option Explicit

Private Sub UserForm_Initialize()
    Dim folderPath1 As String, folderPath2 As String, folderPath3 As String, folderPath4 As String
    Dim fileName As String
    Dim catArray() As String
    Dim i As Long
    Dim newCategoryName As String
    Dim newFolderPath As String
    Dim newLabel As MSForms.Label
    Dim newListBox As MSForms.ListBox
    Dim newTop As Single, newLeft As Single, newWidth As Single, newHeight As Single
    Dim maxBottom As Single
    Dim ctrl As control
    Dim savedCats As String

    ' Set the title of the UserForm.
    Me.Caption = "Select Graphics"
    
    ' Set the base folder.
    BASE_FOLDER = Environ$("APPDATA") & "\Microsoft\AddIns\Trial Ex Addin\QuickGraphics\"
    
    ' Initialize the collection.
    Set colCategoryListBoxes = New Collection
    
    ' (Set up your fixed categories as before—for example:)
    folderPath1 = BASE_FOLDER                    ' "General"
    folderPath2 = BASE_FOLDER & "Circle of Care\"
    folderPath3 = BASE_FOLDER & "Calendar Days\"
    folderPath4 = BASE_FOLDER & "Highlights\"
    
    ' Assume fixed controls (ListBox1, ListBox2, etc.) are already on Frame1.
    PopulateListBox ListBox1, folderPath1
    PopulateListBox ListBox2, folderPath2
    PopulateListBox ListBox3, folderPath3
    PopulateListBox ListBox4, folderPath4
    
    ListBox1.Tag = folderPath1
    ListBox2.Tag = folderPath2
    ListBox3.Tag = folderPath3
    ListBox4.Tag = folderPath4
    
    colCategoryListBoxes.Add ListBox1
    colCategoryListBoxes.Add ListBox2
    colCategoryListBoxes.Add ListBox3
    colCategoryListBoxes.Add ListBox4
    
    ' Now, load any additional categories saved in the registry.
    savedCats = ValidateCategories()   ' Validate and update the saved categories.
    catArray = Split(savedCats, "|")
    
    ' If there are any categories beyond the fixed ones (we assume the first four are fixed)
    For i = 4 To UBound(catArray)
        newCategoryName = catArray(i)
        
        ' Determine the position for the new controls in Frame1.
        maxBottom = 0
        For Each ctrl In Frame1.Controls
            If TypeName(ctrl) = "ListBox" Or TypeName(ctrl) = "Label" Then
                If ctrl.Top + ctrl.Height > maxBottom Then
                    maxBottom = ctrl.Top + ctrl.Height
                End If
            End If
        Next ctrl
        newTop = maxBottom + 10
        
        newLeft = 6
        newWidth = 285.75
        newHeight = 80.35
        
        newFolderPath = BASE_FOLDER & newCategoryName & "\"
        
        ' Create the dynamic controls.
        Set newLabel = Frame1.Controls.Add("Forms.Label.1", "lbl" & newCategoryName, True)
        With newLabel
            .Caption = newCategoryName
            .Left = newLeft
            .Top = newTop
            .Width = 288
            .Height = 18
        End With
        
        Set newListBox = Frame1.Controls.Add("Forms.ListBox.1", "lst" & newCategoryName, True)
        With newListBox
            .Left = newLeft
            .Top = newLabel.Top + newLabel.Height + 6
            .Width = newWidth
            .Height = newHeight
            .Tag = newFolderPath
        End With
        
        PopulateListBox newListBox, newFolderPath
        colCategoryListBoxes.Add newListBox
    Next i
    
    ' (Optional: Update any scrollbar properties if needed.)
End Sub
Private Sub cmdAddSelection_Click()
    Dim sel As Selection
    Dim newPres As presentation
    Dim newSlide As slide
    Dim pastedShapes As ShapeRange
    Dim categoryInput As String
    Dim categoryName As String
    Dim categoryPath As String
    Dim graphicName As String
    Dim filePath As String
    Dim colCategories As Collection
    Dim idx As Long

    ' Get the current selection.
    On Error Resume Next
    Set sel = ActiveWindow.Selection
    On Error GoTo 0
    If sel Is Nothing Then
        MsgBox "No active selection found. Please select an object first.", vbExclamation
        Exit Sub
    End If
    
    ' Ensure the selection consists of shapes.
    If sel.Type <> ppSelectionShapes Then
        MsgBox "Please select one or more shapes.", vbExclamation
        Exit Sub
    End If

    ' Get the list of available categories.
    Set colCategories = GetAvailableCategories()
    
    ' Prompt the user for the category.
    categoryInput = InputBox("Enter the category to add the graphic to:" & vbCrLf & _
                              ListAvailableCategories(), "Select Category")
    If Trim(categoryInput) = "" Then Exit Sub
    
    If IsNumeric(categoryInput) Then
        idx = CLng(categoryInput)
        If idx < 0 Or idx >= colCategories.count Then
            MsgBox "Invalid numeric selection.", vbExclamation
            Exit Sub
        End If
        ' The collection is 1-based; numeric prefix is 0-based.
        categoryName = colCategories(idx + 1)
    Else
        categoryName = categoryInput
    End If
    
    ' Build the folder path for the selected category.
    If LCase(categoryName) = "general" Then
        categoryPath = BASE_FOLDER
    Else
        categoryPath = BASE_FOLDER & categoryName & "\"
        If Dir(categoryPath, vbDirectory) = "" Then
            If MsgBox("Category '" & categoryName & "' does not exist. Create it?", vbYesNo + vbExclamation) = vbYes Then
                On Error GoTo ErrHandler
                MkDir categoryPath
            Else
                Exit Sub
            End If
        End If
    End If
    
    ' Prompt the user for a graphic name (without extension).
    graphicName = InputBox("Enter a name for the graphic (without extension):", "Graphic Name")
    If Trim(graphicName) = "" Then Exit Sub
    
    ' Create a new blank presentation.
    Set newPres = Application.Presentations.Add(msoTrue)
    If newPres.Slides.count = 0 Then
        Set newSlide = newPres.Slides.Add(1, ppLayoutBlank)
    Else
        Set newSlide = newPres.Slides(1)
    End If
    
    ' Match the slide dimensions of the active presentation.
    If Not ActivePresentation Is Nothing Then
        With newPres.PageSetup
            .slideWidth = ActivePresentation.PageSetup.slideWidth
            .slideHeight = ActivePresentation.PageSetup.slideHeight
        End With
    End If
    
    ' Copy the selected shapes and paste them onto the new slide.
    sel.ShapeRange.Copy
    Set pastedShapes = newSlide.Shapes.Paste
    
    ' Build the full file path.
    filePath = categoryPath & graphicName & ".pptx"
    
    On Error GoTo ErrHandler
    newPres.SaveAs filePath
    newPres.Close
    
    MsgBox "Graphic added successfully as '" & graphicName & ".pptx' in category '" & categoryName & "'.", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub



' ===============================
' PopulateListBox: in a standard module you may already have this,
' but here it is for reference.
' ===============================
Public Sub PopulateListBox(ByRef lstBox As MSForms.ListBox, ByVal folderPath As String)
    Dim fileName As String
    lstBox.Clear
    fileName = Dir(folderPath & "*.pptx")
    Do While fileName <> ""
        lstBox.AddItem fileName
        fileName = Dir
    Loop
End Sub

' ===============================
' CommandButton1_Click – Import/Use the selected graphic.
' ===============================
Private Sub CommandButton1_Click()
    Dim selectedFile As String
    Dim fullPath As String
    Dim categoryFolder As String
    Dim lb As MSForms.ListBox
    Dim found As Boolean
    found = False
    
    ' Loop over all category list boxes to see which one has a selection.
    For Each lb In colCategoryListBoxes
        If lb.ListIndex <> -1 Then
            selectedFile = lb.List(lb.ListIndex)
            categoryFolder = lb.Tag  ' Folder path stored in Tag
            found = True
            Exit For
        End If
    Next lb
    
    If Not found Then
        MsgBox "Please select a file from one of the lists.", vbExclamation
        Exit Sub
    End If
    
    fullPath = categoryFolder & selectedFile
    ImportSlides fullPath
End Sub



' ===============================
' Other event handlers remain unchanged
' (e.g., your ListBox#_Click events to deselect others, cmdClose_Click, etc.)
' ===============================
Private Sub ListBox1_Click()
    ListBox2.ListIndex = -1
    ListBox3.ListIndex = -1
    ListBox4.ListIndex = -1
End Sub

Private Sub ListBox2_Click()
    ListBox1.ListIndex = -1
    ListBox3.ListIndex = -1
    ListBox4.ListIndex = -1
End Sub

Private Sub ListBox3_Click()
    ListBox1.ListIndex = -1
    ListBox2.ListIndex = -1
    ListBox4.ListIndex = -1
End Sub

Private Sub ListBox4_Click()
    ListBox1.ListIndex = -1
    ListBox2.ListIndex = -1
    ListBox3.ListIndex = -1
End Sub

Private Sub cmdClose_Click()
    Dim currentSlideIndex As Long
    
    If Not Application.ActiveWindow Is Nothing Then
        If Not Application.ActiveWindow.View.slide Is Nothing Then
            currentSlideIndex = Application.ActiveWindow.View.slide.SlideIndex
        End If
    End If

    Unload Me
    ToggleFullScreen currentSlideIndex
End Sub

Private Sub RemoveGraphic_Click()
    Dim categoryInput As String
    Dim categoryName As String
    Dim colCategories As Collection
    Dim idx As Long
    Dim fileFolder As String
    Dim fileList As New Collection
    Dim fName As String
    Dim fileDisplay As String
    Dim fileIndexInput As String
    Dim numericIndex As Long
    Dim i As Long
    Dim selectedFile As String
    Dim totalContentHeight As Single


    ' Get available categories.
    Set colCategories = GetAvailableCategories()
    
    ' Prompt the user to select a category.
    categoryInput = InputBox("Enter the category from which to remove a graphic:" & vbCrLf & _
                             "Available categories:" & vbCrLf & ListAvailableCategories(), "Select Category")
    If Trim(categoryInput) = "" Then Exit Sub

    ' Determine the category name from input.
    If IsNumeric(categoryInput) Then
        idx = CLng(categoryInput)
        If idx < 0 Or idx >= colCategories.count Then
            MsgBox "Invalid numeric selection.", vbExclamation
            Exit Sub
        End If
        categoryName = colCategories(idx + 1)
    Else
        categoryName = categoryInput
    End If
    
    ' Build the folder path for the selected category.
    If LCase(categoryName) = "general" Then
        fileFolder = BASE_FOLDER
    Else
        fileFolder = BASE_FOLDER & categoryName & "\"
    End If
    
    ' Build a collection of files (*.pptx) in the folder.
    fName = Dir(fileFolder & "*.pptx")
    Do While fName <> ""
        fileList.Add fName
        fName = Dir
    Loop
    
    If fileList.count = 0 Then
        MsgBox "No graphics found in category '" & categoryName & "'.", vbExclamation
        Exit Sub
    End If
    
    ' Build a display list with numeric prefixes.
    fileDisplay = ""
    For i = 1 To fileList.count
        fileDisplay = fileDisplay & (i - 1) & ": " & fileList(i) & vbCrLf
    Next i
    
    ' Prompt the user to choose the graphic to delete.
    fileIndexInput = InputBox("Enter the number corresponding to the graphic you want to remove:" & vbCrLf & vbCrLf & fileDisplay, "Remove Graphic")
    If Trim(fileIndexInput) = "" Then Exit Sub
    
    If Not IsNumeric(fileIndexInput) Then
        MsgBox "Please enter a valid numeric selection.", vbExclamation
        Exit Sub
    End If
    
    numericIndex = CLng(fileIndexInput)
    If numericIndex < 0 Or numericIndex >= fileList.count Then
        MsgBox "Invalid selection.", vbExclamation
        Exit Sub
    End If
    
    selectedFile = fileList(numericIndex + 1)
    
    ' Confirm deletion.
    If MsgBox("Are you sure you want to delete '" & selectedFile & "' from category '" & categoryName & "'?", vbYesNo + vbExclamation) = vbYes Then
        On Error GoTo ErrHandler
        Kill fileFolder & selectedFile
        MsgBox "Graphic '" & selectedFile & "' has been removed.", vbInformation
    End If

    Exit Sub
    
ErrHandler:
    MsgBox "Error deleting file: " & Err.Description, vbCritical
End Sub

Private Sub cmdAddCategory_Click()
    Dim newCategoryName As String
    Dim newFolderPath As String
    Dim newListBox As MSForms.ListBox
    Dim newLabel As MSForms.Label
    Dim newTop As Single, newLeft As Single, newWidth As Single, newHeight As Single
    Dim ctrl As control
    Dim maxBottom As Single
    Dim totalContentHeight As Single
    Dim savedCats As String
    
    ' Prompt the user for the new category name.
    newCategoryName = InputBox("Enter the new category name:", "Add Category")
    If Trim(newCategoryName) = "" Then Exit Sub
    
    ' Construct the new folder path.
    newFolderPath = BASE_FOLDER & newCategoryName & "\"
    
    ' Create the folder if it does not exist.
    If Dir(newFolderPath, vbDirectory) = "" Then
        On Error GoTo ErrHandler
        MkDir newFolderPath
    Else
        If MsgBox("Folder already exists. Do you still want to add this category?", vbYesNo + vbQuestion) = vbNo Then
            Exit Sub
        End If
    End If
    
    ' *** Update the saved categories in the registry ***
    savedCats = LoadCategoriesString()
    If InStr(1, savedCats, newCategoryName, vbTextCompare) = 0 Then
        savedCats = savedCats & "|" & newCategoryName
        SaveCategories savedCats
    End If
    
    ' Determine the position for the new controls.
    ' Find the bottom-most position among existing controls in Frame1.
    maxBottom = 0
    For Each ctrl In Frame1.Controls
        If TypeName(ctrl) = "ListBox" Or TypeName(ctrl) = "Label" Then
            If ctrl.Top + ctrl.Height > maxBottom Then
                maxBottom = ctrl.Top + ctrl.Height
            End If
        End If
    Next ctrl
    newTop = maxBottom + 10  ' 10 points below the lowest control in Frame1
    
    ' Define standard left and width based on your design.
    newLeft = 6
    newWidth = 285.75
    newHeight = 80.35  ' Same as your other list boxes
    
    ' Add a new label for the category to Frame1.
    Set newLabel = Frame1.Controls.Add("Forms.Label.1", "lbl" & newCategoryName, True)
    With newLabel
        .Caption = newCategoryName
        .Left = newLeft
        .Top = newTop
        .Width = 288
        .Height = 18
    End With
    
    ' Add a new ListBox below the label to Frame1.
    Set newListBox = Frame1.Controls.Add("Forms.ListBox.1", "lst" & newCategoryName, True)
    With newListBox
        .Left = newLeft
        .Top = newLabel.Top + newLabel.Height + 6 ' 6 points gap below the label
        .Width = newWidth
        .Height = newHeight
        ' Set the Tag to the folder path so that later we can retrieve it.
        .Tag = newFolderPath
    End With
    
    ' Populate the new list box with files from the new folder.
    PopulateListBox newListBox, newFolderPath
    
    ' Add the new list box to our collection.
    colCategoryListBoxes.Add newListBox
    
    ' Recalculate the total content height of Frame1.
    totalContentHeight = 0
    For Each ctrl In Frame1.Controls
        If ctrl.Top + ctrl.Height > totalContentHeight Then
            totalContentHeight = ctrl.Top + ctrl.Height
        End If
    Next ctrl
    
    ' (Optional: update a scrollbar here if needed.)
    
    MsgBox "New category '" & newCategoryName & "' added.", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Error adding category: " & Err.Description, vbCritical
End Sub


Private Function GetAvailableCategories() As Collection
    Dim col As New Collection
    Dim folderName As String
    Dim path As String
    
    col.Add "General"
    
    path = BASE_FOLDER
    folderName = Dir(path, vbDirectory)
    
    Do While folderName <> ""
        If (GetAttr(path & folderName) And vbDirectory) = vbDirectory Then
            If folderName <> "." And folderName <> ".." Then
                col.Add folderName
            End If
        End If
        folderName = Dir
    Loop
    
    Set GetAvailableCategories = col
End Function

Private Function ListAvailableCategories() As String
    Dim col As Collection
    Dim i As Long
    Dim s As String
    
    Set col = GetAvailableCategories()
    s = ""
    For i = 1 To col.count
        s = s & (i - 1) & ": " & col(i) & vbCrLf
    Next i
    ListAvailableCategories = s
End Function

Private Sub CloseButton_Click()
    Dim currentSlideIndex As Long
    Dim currentPres As presentation
    
    If Not Application.ActiveWindow Is Nothing Then
        If Not Application.ActiveWindow.View.slide Is Nothing Then
            currentSlideIndex = Application.ActiveWindow.View.slide.SlideIndex
        End If
    End If

    Unload Me
    ToggleFullScreen currentSlideIndex
End Sub

Private Sub LoadFilesInCategory(ByVal categoryName As String)
    Dim folderPath As String
    Dim fileName As String
    
    If LCase(categoryName) = "general" Then
        folderPath = BASE_FOLDER
    Else
        folderPath = BASE_FOLDER & categoryName & "\"
    End If
    
    ' Use ListBox1 instead of lstFiles (if that’s what you intended)
    ListBox1.Clear
    fileName = Dir(folderPath & "*.pptx")
    Do While fileName <> ""
        ListBox1.AddItem fileName
        fileName = Dir
    Loop
End Sub



Private Function ValidateCategories() As String
    Dim savedCats As String, catArray() As String, newCats As String
    Dim i As Long, catName As String, folderPath As String
    
    savedCats = LoadCategoriesString()  ' Retrieve the saved categories string.
    catArray = Split(savedCats, "|")
    newCats = ""
    
    For i = LBound(catArray) To UBound(catArray)
        catName = catArray(i)
        ' For "General", the folder path is just BASE_FOLDER.
        If LCase(catName) = "general" Then
            folderPath = BASE_FOLDER
        Else
            folderPath = BASE_FOLDER & catName & "\"
        End If
        
        ' If the folder exists, include it in the new list.
        If Dir(folderPath, vbDirectory) <> "" Then
            If newCats = "" Then
                newCats = catName
            Else
                newCats = newCats & "|" & catName
            End If
        End If
    Next i
    
    ' If the validated list differs from the saved list, update the registry.
    If newCats <> savedCats Then
        SaveCategories newCats
    End If
    
    ValidateCategories = newCats
End Function


Private Sub SaveCategories(ByVal catList As String)
    ' Save the delimited string of categories to the registry.
    SaveSetting APP_NAME, CATEGORY_SECTION, CATEGORY_KEY, catList
End Sub

Private Function LoadCategoriesString() As String
    ' Retrieve the saved categories; if none exist, use a default.
    LoadCategoriesString = GetSetting(APP_NAME, CATEGORY_SECTION, CATEGORY_KEY, "General|Circle of Care|Calendar Days|Highlights")
End Function

