Attribute VB_Name = "modFunctions"
Option Explicit

' #############################################################################
' #                                                                           #
' #  This module contains a load of general functions to make stuff a bit     #
' #  easier. They arnt really in any grouping or order, but there arn't that  #
' #                many of them, so its not to confusing.                     #
' #                                                                           #
' #############################################################################


Public Function Almost(x1 As Single, x2 As Single, y1 As Single, y2 As Single) As Boolean
    'This returns whether two coordinates are close to each other or not
    If x1 - 5 < x2 And x1 + 5 > x2 And y1 - 5 < y2 And y1 + 5 > y2 Then Almost = True
End Function


Public Function InvertBo(Value As Boolean) As Boolean
    'This reverses the value of a boolean. Ie. if true is given, false is returned
    If Value = True Then InvertBo = False Else InvertBo = True
End Function


Public Function IntBo(Value As Boolean) As Integer
    'This returns a integer value from a boolean, Ie. if true is given, one is returned
    If Value = True Then IntBo = 1 Else IntBo = 0
End Function


Public Sub Swap(Value1 As Variant, Value2 As Variant)
    'This just swaps two values over
    Dim Temp As Variant
    Temp = Value1: Value1 = Value2: Value2 = Temp
End Sub


Public Function Destroy(FileName) As Boolean
    'This checks if a file exists, and if so, removes it
    If Dir(FileName) <> "" Then Kill FileName: Destroy = True
End Function


Public Function SelectedNode(Tree As TreeView) As Integer
    'This returns the index of the selected node in a tree control, or a zero is no node is selected
    Dim n As Integer
    For n = 1 To Tree.Nodes.Count
        If Tree.Nodes(n).Selected = True Then SelectedNode = n: Exit Function
    Next n
End Function


Public Function GetEditOption() As Integer
    'This returns the number of the edit button selected in the toolbar
    Dim n As Integer
    For n = 0 To frmMain.optEdit.Count - 1
        If frmMain.optEdit(n) = True Then GetEditOption = n: Exit Function
    Next n
End Function


Public Function EditButton() As Integer
    'This returns the number of the edit button selected in the toolbar
    Dim n As Integer
    For n = 1 To frmMain.tbar(1).buttons.Count
        If frmMain.tbar(1).buttons(n).Value = tbrPressed Then EditButton = n: Exit Function
    Next n
End Function


Public Function MaxLength(Data As String, MaxL As Integer, SplitPos As Integer) As String
    'This takes a string and makes sure it does go over a certain length. If it does, it cuts a bit
    'of the string out at the beggining. Its used for the file history menus
    If Len(Data) > (MaxL + SplitPos) Then MaxLength = Mid(Data, 1, SplitPos) & "..." & Mid(Data, Len(Data) - (MaxL + SplitPos)) Else MaxLength = Data
End Function


Public Function NewFileKey() As String
    'This returns a unique string used as a key for file objects
    Static FileCount As Long
    FileCount = FileCount + 1
    NewFileKey = "File" & FileCount
End Function


Public Function RightClip(txtString As String, Length As Integer) As String
    'This returns a string minus the last few characters
    RightClip = Mid(txtString, 1, Len(txtString) - Length)
End Function


Public Function UpdateLayerMenu(Form As frmEdit)
    'This sets up the layers drop down menu for the given form. It unloads the existing menu,
    'and recreates it using the current data from the file class assosiated with the form.
    Dim n As Integer
    For n = Form.mnuAddtoLayer.Count - 1 To 1 Step -1
        Unload Form.mnuAddtoLayer(n)
    Next n
    For n = 1 To Am8(Form.Tablet.FileKey).Layers.CountLayers
        Load Form.mnuAddtoLayer(n)
        Form.mnuAddtoLayer(n).Caption = Am8(Form.Tablet.FileKey).Layers(n).LayerName
        Form.mnuAddtoLayer(n).Visible = True
    Next n
    Form.mnuEditPopup(12).Visible = True
    If Form.mnuAddtoLayer.Count <> 1 Then Form.mnuAddtoLayer(0).Visible = False
End Function


Public Function CheckOverwrite(FileName, Mode As Byte) As Boolean
    'This checks to see if a file already exists. If it does, it can ask you
    'whether you want to overwrite the file or not..
    If Dir(FileName) = "" Then
        CheckOverwrite = True
    Else
        If GetAttr(FileName) Mod 2 = 1 Then MsgBox amReadOnlyFile, vbInformation: Exit Function
        If MsgBox(amOverwritefile, vbQuestion + vbYesNo + vbDefaultButton2, "Replace file") = vbYes Then CheckOverwrite = True
    End If
End Function


Public Sub Make_LookUp()
    'Very important, it pre-calculates all the sine and cosine values
    'from -361 to 360 to 1 decimal place, so you can rotate down to
    'a tenth of a degree
    Const PI = 3.14159265358979
    Dim i As Single
    For i = -361 To 361 Step 0.1
        Sine(i * 10) = Sin(i / 180 * PI)
        Cosine(i * 10) = Cos(i / 180 * PI)
    Next
End Sub


Public Function RotatePoint(Vertex As clsVertex, Angle1 As Single, Angle2 As Single, Angle3 As Single, Optional CenterX As Integer = 0, Optional CenterY As Integer = 0, Optional CenterZ As Integer = 0) As clsVertex
    'This is used to rotate all the values in the array Rotated(). You must
    'fist fill the array with the values, then call this function which
    'rotates the values through the given angles. You also supply the number
    'of values in the array.
    Dim x As Single, y As Single, z As Single
    x = Vertex.x - CenterX: y = Vertex.y - CenterY: z = Vertex.z - CenterZ
    Angle1 = Angle1 Mod 3600: Angle2 = Angle2 Mod 3600: Angle3 = Angle3 Mod 3600
    Set RotatePoint = New clsVertex
    Dim XRotated As Single, YRotated As Single, ZRotated As Single
    XRotated = x
    YRotated = Cosine(Angle1) * y - Sine(Angle1) * z
    ZRotated = Sine(Angle1) * y + Cosine(Angle1) * z
    x = XRotated:      y = YRotated:      z = ZRotated
    XRotated = Cosine(Angle2) * x - Sine(Angle2) * z
    YRotated = y
    ZRotated = Sine(Angle2) * x + Cosine(Angle2) * z
    x = XRotated:      y = YRotated:      z = ZRotated
    XRotated = Cosine(Angle3) * x - Sine(Angle3) * y
    YRotated = Sine(Angle3) * x + Cosine(Angle3) * y
    ZRotated = z
    RotatePoint.x = XRotated + CenterX
    RotatePoint.y = YRotated + CenterY
    RotatePoint.z = ZRotated + CenterZ
End Function


Public Function SelectFileName(FileType, DialogTitle) As String
    'This displays the Show Open dialog box, and fills in the drop down list
    'with the required contents. It returns a file name that you select in the
    'box, or nothing if you press cancel.
    Static OpenIndex As Integer, CopyIndex As Integer, ImportIndex As Integer, CompileIndex As Integer, PictureIndex As Integer
    On Error GoTo CanceledSelect
    frmMain.GetFile.Flags = 4
    frmMain.GetFile.DialogTitle = DialogTitle
    Select Case FileType
        Case "Am8"
            frmMain.GetFile.Filter = "AnimationShop files (*.Am8) |*.Am8|Old AnimationShop files (*.Am5) |*.Am5|All files (*.*) |*.*"
            frmMain.GetFile.FilterIndex = OpenIndex
        Case "Copy"
            frmMain.GetFile.Filter = "Copy files (*.Cpy) |*.cpy|All files (*.*) |*.*"
            frmMain.GetFile.FilterIndex = CopyIndex
        Case "Import"
            frmMain.GetFile.Filter = "Compiled data files (*.dat)|*.dat|Terain data files (*.map)|*.map|Copied Files (*.cpy)|*.cpy|3D data files (*.asc)|*.asc|Pure Text (*.txt) |*.txt|All 3D files|*.dat; *.map; *.asc|All files (*.*) |*.*"
            frmMain.GetFile.FilterIndex = ImportIndex
        Case "Compile"
            frmMain.GetFile.Filter = "Compiled data files (*.dat) |*.dat|All files (*.*) |*.*"
            frmMain.GetFile.FilterIndex = CompileIndex
        Case "Picture"
            frmMain.GetFile.Filter = "Bitmaps (*.bmp) |*.Bmp|All files (*.*) |*.*"
            frmMain.GetFile.FilterIndex = PictureIndex
    End Select
    frmMain.GetFile.FileName = ""
    frmMain.GetFile.ShowOpen
    SelectFileName = frmMain.GetFile.FileName
    Select Case FileType
        Case "Am8":     OpenIndex = frmMain.GetFile.FilterIndex
        Case "Copy":    CopyIndex = frmMain.GetFile.FilterIndex
        Case "Import":  ImportIndex = frmMain.GetFile.FilterIndex
        Case "Compile": CompileIndex = frmMain.GetFile.FilterIndex
        Case "Picture": PictureIndex = frmMain.GetFile.FilterIndex
    End Select
CanceledSelect:
End Function


Public Function SetFileName(FileType, DialogTitle) As String
    'This does the save as the above function, but displays the 'Save As' dialog box instead
    On Error GoTo CanceledSet
    frmMain.GetFile.DialogTitle = DialogTitle
    Select Case FileType
        Case "Copy"
            frmMain.GetFile.Filter = "Copy files (*.Cpy) |*.cpy|All files (*.*) |*.*"
            frmMain.GetFile.FilterIndex = 1
        Case "Am8"
            frmMain.GetFile.Filter = "AnimationShop files (*.Am8) |*.Am8|All files (*.*) |*.*"
            frmMain.GetFile.FilterIndex = 1
        Case "Import"
            frmMain.GetFile.Filter = "Compiled data files (*.dat) |*.dat|Terain data files (*.map) |*.map|3D data files (*.asc) |*.asc|All 3D files |*.dat; *.map; *.asc; *.txt|All files (*.*) |*.*"
            frmMain.GetFile.FilterIndex = 1
        Case "Compile"
            frmMain.GetFile.Filter = "Compiled data files (*.dat) |*.dat|All files (*.*) |*.*"
            frmMain.GetFile.FilterIndex = 1
        Case "Picture"
            frmMain.GetFile.Filter = "Bitmaps (*.bmp) |*.Bmp|All files (*.*) |*.*"
            frmMain.GetFile.FilterIndex = 1
    End Select
    frmMain.GetFile.FileName = ""
    frmMain.GetFile.ShowSave
    SetFileName = frmMain.GetFile.FileName
CanceledSet:
End Function


Public Function FolderIsMissing(FolderName As String) As Boolean
    'This checks to see if a given folder exists. If not, it will create the
    'folder and return a True value, so a message can be displayed
    On Error GoTo FolderMissing
    Open App.Path & "\" & FolderName & "\56.567" For Output As #1
    Close
    Kill App.Path & "\" & FolderName & "\56.567"
    Exit Function
FolderMissing:
    MkDir App.Path & "\" & FolderName & "\"
    FolderIsMissing = True
End Function


Public Function GetGallaryFolders() As Boolean
    'This loads the contents of the gallaries folder into the dropdown list
    'so you can select a sub folder to look in...
    Dim StartOn As String, n As Integer, m As Integer, FolderName As String
    On Error GoTo MissingGFolder
    StartOn = frmMain.cmbGallary.Text
    frmMain.Dir1.Path = App.Path
    frmMain.Dir1.Path = App.Path & "\data\gallarys"
    frmMain.cmbGallary.Clear
    frmMain.cmbGallary.AddItem "[None]"
    For n = frmMain.Dir1.ListIndex + 1 To frmMain.Dir1.ListCount - 1
        FolderName = frmMain.Dir1.List(n)
        For m = Len(FolderName) To 1 Step -1
            If Mid(FolderName, m, 1) = "\" Then Exit For
        Next m
        frmMain.cmbGallary.AddItem Mid(FolderName, m + 1)
    Next n
    If frmMain.cmbGallary.ListCount = 1 Then
        frmMain.Gallary.Visible = False: frmMain.Label10.Visible = True
        frmMain.Label11.Visible = True
    Else
        frmMain.Gallary.Visible = True
        frmMain.cmbGallary.RemoveItem (0): frmMain.Label10.Visible = False
        frmMain.Label11.Visible = False
    End If
    GetGallaryFolders = True
MissingGFolder:
End Function


Public Function ThreeLength(Number As Integer) As String
    ThreeLength = Trim(Str(Number))
    If Len(ThreeLength) = 1 Then ThreeLength = "00" & ThreeLength
    If Len(ThreeLength) = 2 Then ThreeLength = "0" & ThreeLength
End Function



