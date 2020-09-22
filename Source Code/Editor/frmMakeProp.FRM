VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEntEdit 
   Caption         =   "Entity Creator"
   ClientHeight    =   7665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14310
   Icon            =   "frmMakeProp.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   7665
   ScaleWidth      =   14310
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAct 
      Caption         =   "Set Description"
      Height          =   375
      Index           =   4
      Left            =   6960
      TabIndex        =   10
      ToolTipText     =   "Click to edit the disciption of this entity"
      Top             =   7200
      Width           =   1335
   End
   Begin VB.CommandButton cmdAct 
      Caption         =   "Remove"
      Height          =   375
      Index           =   6
      Left            =   4080
      TabIndex        =   9
      ToolTipText     =   "Click to remove the selected entity"
      Top             =   7200
      Width           =   1335
   End
   Begin VB.CommandButton cmdAct 
      Caption         =   "New"
      Height          =   375
      Index           =   5
      Left            =   5520
      TabIndex        =   8
      ToolTipText     =   "Click to create a new entity"
      Top             =   7200
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog cmColour 
      Left            =   6480
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdAct 
      Caption         =   "Rename"
      Height          =   375
      Index           =   3
      Left            =   8400
      TabIndex        =   7
      ToolTipText     =   "Click to rename this entity"
      Top             =   7200
      Width           =   1335
   End
   Begin VB.ComboBox cmbSelect 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtText 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   1680
      TabIndex        =   5
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox EntityList 
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   120
      Width           =   12495
   End
   Begin VB.CommandButton cmdAct 
      Caption         =   "Cancel"
      Height          =   375
      Index           =   1
      Left            =   11280
      TabIndex        =   3
      ToolTipText     =   "Click to confirm the changes made in this window"
      Top             =   7200
      Width           =   1335
   End
   Begin VB.CommandButton cmdAct 
      Caption         =   "Okay"
      Height          =   375
      Index           =   2
      Left            =   9840
      TabIndex        =   2
      ToolTipText     =   "Click to cancel any changes made in this window"
      Top             =   7200
      Width           =   1335
   End
   Begin VB.CommandButton cmdAct 
      Caption         =   "Help"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Click to get help on this window"
      Top             =   7200
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid grProperties 
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   10398
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin VB.Menu menuPopup 
      Caption         =   "menuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuMain 
         Caption         =   "Add Property"
         Index           =   1
      End
      Begin VB.Menu mnuMain 
         Caption         =   "Insert Property"
         Index           =   2
      End
      Begin VB.Menu mnuMain 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuMain 
         Caption         =   "Remove Property"
         Index           =   4
      End
      Begin VB.Menu mnuMain 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuMain 
         Caption         =   "Move up"
         Index           =   6
      End
      Begin VB.Menu mnuMain 
         Caption         =   "Move down"
         Index           =   7
      End
   End
End
Attribute VB_Name = "frmEntEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RowON As Integer, ColON As Integer, Saved As Boolean, EntityName As String



Public Sub RunAtStart(Optional SetEntityName As String = "")
    'This routine can be called to display the form. If n entity name is passed to the code, then
    'it is automaticly displayed to start off with
    Dim n As Integer
    grProperties.Rows = 1
    UpdateDropdownList
    grProperties.Enabled = False
    For n = 1 To EntityList.ListCount
        If EntityList.List(n - 1) = SetEntityName Then EntityList.ListIndex = n - 1: EntityName = SetEntityName
    Next n
    Saved = True
    Show vbModal
End Sub



Private Sub Form_Load()
    'When you load the form, this sets up the column headers in the grid
    With grProperties
        .Cols = 10
        .TextMatrix(0, 1) = "Name"
        .TextMatrix(0, 2) = "Type"
        .TextMatrix(0, 3) = "ToolTip"
        .TextMatrix(0, 4) = "Choices"
        .TextMatrix(0, 5) = "Default"
        .TextMatrix(0, 6) = "Locked"
        .TextMatrix(0, 7) = "New Column"
        .TextMatrix(0, 8) = "New Page"
        .TextMatrix(0, 9) = "Width"
        .ColWidth(7) = 1300
        .ColWidth(3) = 2500
        .ColWidth(4) = 2500
        .ColWidth(0) = 400
    End With
End Sub



Private Sub UpdateDropdownList()
    'This displays the name of each entity in the drop down list at the top of the screen
    Dim Am As clsEntity
    EntityList.Clear
    grProperties.Enabled = True
    For Each Am In Am8.Properties: EntityList.AddItem Am.EntityName: Next Am
End Sub



Private Sub cmdAct_Click(Index As Integer)
    'This controls the buttons at the bottom of the screen
    Dim NewName As String, n As Integer, NewDiscription As String
    Select Case Index
    
        Case 1
            Unload Me
    
        Case 2
            SaveEntityDetails EntityName
            Am8.Properties.SaveEntities
            Unload Me
        
        Case 3
            'This renames the selected entity to a new name
            NewName = InputBox("Enter a new name for this entity", "Rename Entity", EntityName)
            If NewName <> "" Then
                Am8.Properties.Add NewName
                For n = 1 To Am8.Properties(EntityName).Count
                    With Am8.Properties(EntityName)(n)
                        Am8.Properties(NewName).Add .PropertyName, .PropertyType, .ToolTip, .Choice, .Default, .Width, .Locked, .NewColumn, .NewPage
                    End With
                Next n
                Am8.Properties(NewName).Discription = Am8.Properties(EntityName).Discription
                Am8.Properties.Remove EntityName
                EntityName = NewName
                UpdateDropdownList
                EntityList.ListIndex = EntityList.ListCount - 1
            End If

        Case 4
            'This sets the discription of the entity
            NewDiscription = InputBox("Enter the discription of this entity", , Am8.Properties(EntityList).Discription)
            If NewDiscription <> "" Then Am8.Properties(EntityList).Discription = NewDiscription

        Case 5
            'This allows you to create a new entity
            NewName = InputBox("Enter the name of the new entity")
            If NewName <> "" Then
                Am8.Properties.Add NewName
                EntityList.AddItem NewName
                grProperties.Enabled = True
                EntityList.ListIndex = EntityList.ListCount - 1
            End If
        
        Case 6
            'This removes an entity from the list
            If MsgBox(EntityList.Text & vbNewLine & vbNewLine & "Are you sure you want to remove this entity", vbYesNo + vbQuestion) = vbYes Then
                Am8.Properties.Remove EntityList
                EntityList.RemoveItem EntityList.ListIndex
                EntityName = EntityList.Text
                If EntityList.ListCount > 0 Then EntityList.ListIndex = 0 Else grProperties.Rows = 1: grProperties.Enabled = False

            End If
            
    End Select
End Sub



Private Sub EntityList_Click()
    'When you select a different entity from the list, this saves the current details, and then loads the
    'new entity
    If EntityName <> "" Then SaveEntityDetails EntityName
    EntityName = EntityList
    If EntityName <> "" Then LoadEntityDetails EntityName
End Sub



Private Sub Form_Unload(Cancel As Integer)
    'When you close the form, this reloads the entities from disk, so if changes were canceled,
    'the original ones are loaded, and if changes were saved, they are already writen onto the disk
    If Saved = False Then Am8.Properties.LoadEntities
    EntityName = ""
End Sub



Private Sub grProperties_DblClick()
    'When you double click on the grid, this displays either a text box or a combo box over the
    'cell you clicked on, and displays the data in the cell. If its a combo box, then the right
    'data is entered into the list.
    Dim n As Integer
    RowON = grProperties.Row
    ColON = grProperties.col
    Saved = False
    Select Case grProperties.col
        Case 1, 3, 4, 5, 9
            'This displays the text box control over the grid
            If grProperties.TextMatrix(grProperties.Row, 2) = "colour" And grProperties.col = 5 Then
                cmColour.ShowColor
                grProperties.TextMatrix(RowON, ColON) = cmColour.Color
            Else
                txtText.Top = grProperties.RowPos(grProperties.Row) + grProperties.Top
                txtText.Left = grProperties.ColPos(grProperties.col) + grProperties.Left + 34
                txtText.Width = grProperties.ColWidth(grProperties.col) + 17
                txtText.Text = grProperties.TextMatrix(RowON, ColON)
                txtText.Visible = True: txtText.SetFocus
            End If
                    
        Case 2, 6, 7, 8
            'This displays the dropdown box, and fills the lists the the right data
            cmbSelect.Top = grProperties.RowPos(grProperties.Row) + grProperties.Top
            cmbSelect.Left = grProperties.ColPos(grProperties.col) + grProperties.Left + 34
            cmbSelect.Width = grProperties.ColWidth(grProperties.col) + 17: cmbSelect.Clear
            
            Select Case grProperties.col
                Case 2
                    cmbSelect.AddItem "text": cmbSelect.AddItem "tick"
                    cmbSelect.AddItem "file": cmbSelect.AddItem "colour"
                    cmbSelect.AddItem "slider": cmbSelect.AddItem "combo"
                    cmbSelect.AddItem "editcombo": cmbSelect.AddItem "help"
                    For n = 1 To cmbSelect.ListCount
                        If grProperties.TextMatrix(RowON, ColON) = cmbSelect.List(n - 1) Then cmbSelect.ListIndex = n - 1
                    Next n
                    
                Case 6, 7, 8
                    cmbSelect.AddItem "True": cmbSelect.AddItem "False"
                    If grProperties.TextMatrix(RowON, ColON) = "True" Then cmbSelect.ListIndex = 0 Else cmbSelect.ListIndex = 1
            
            End Select
            cmbSelect.Visible = True
            cmbSelect.SetFocus
    End Select
End Sub



Private Sub Form_Resize()
    'This moves the controls to fit onto the form
    Dim n As Integer
    On Error Resume Next
    For n = 0 To 7: cmdAct(n).Top = ScaleHeight - 500: Next n
    cmdAct(6).Left = ScaleWidth - 8700
    cmdAct(5).Left = ScaleWidth - 7250
    cmdAct(4).Left = ScaleWidth - 5800
    cmdAct(3).Left = ScaleWidth - 4350
    cmdAct(2).Left = ScaleWidth - 2900
    cmdAct(1).Left = ScaleWidth - 1450
    grProperties.Width = ScaleWidth - (grProperties.Left * 2)
    grProperties.Height = ScaleHeight - grProperties.Top - 600
    EntityList.Width = grProperties.Width
End Sub



Private Sub grProperties_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'This controsl the popup menu for when you right click on the grid
    If Button = 2 Then PopupMenu menuPopup
End Sub



Private Sub LoadEntityDetails(EntityName As String)
    'This loads the details of an entity into the text grid
    grProperties.Rows = 1
    Dim Am As clsProperty
    Dim Counter As Integer
    For Each Am In Am8.Properties(EntityName)
        Counter = Counter + 1
        grProperties.AddItem Counter & vbTab & Am.PropertyName & vbTab & Am.PropertyType & vbTab & Am.ToolTip & vbTab & Am.Choice & vbTab & Am.Default & vbTab & Am.Locked & vbTab & Am.NewColumn & vbTab & Am.NewPage & vbTab & Am.Width
    Next Am
    grProperties.Enabled = True
End Sub



Private Sub SaveEntityDetails(EntityName As String)
    'This saves the details of an entity from the text grid to memmory
    Dim n As Integer, Discription As String
    If EntityName <> "" Then
        Discription = Am8.Properties(EntityName).Discription
        Am8.Properties.Remove EntityName
        Am8.Properties.Add EntityName
        Am8.Properties(EntityName).Discription = Discription
        With grProperties
            For n = 1 To grProperties.Rows - 1
                 Am8.Properties(EntityName).Add .TextMatrix(n, 1), .TextMatrix(n, 2), .TextMatrix(n, 3), .TextMatrix(n, 4), .TextMatrix(n, 5), Val(.TextMatrix(n, 9)), .TextMatrix(n, 6), .TextMatrix(n, 7), .TextMatrix(n, 8)
            Next n
        End With
    End If
End Sub



Private Sub mnuMain_Click(Index As Integer)
    'This is the popup menu code, which insert and deletes new properties
    Dim n As Integer
    Saved = False
    Select Case Index
        Case 1: grProperties.AddItem vbTab & vbTab & "text" & vbTab & vbTab & vbTab & vbTab & "False" & vbTab & "False" & vbTab & "False" & vbTab & "2000"
        Case 2: If grProperties.Rows > 1 Then grProperties.AddItem vbTab & vbTab & "text" & vbTab & vbTab & vbTab & vbTab & "False" & vbTab & "False" & vbTab & "False" & vbTab & "2000", grProperties.Row
        Case 4: grProperties.RemoveItem (grProperties.RowSel)
        Case 6:
        Case 7:
    End Select
    For n = 1 To grProperties.Rows - 1: grProperties.TextMatrix(n, 0) = n: Next n
End Sub



Private Sub txtText_LostFocus()
    'When the text box looses focus, it becomes invisiable
    txtText.Visible = False
    grProperties.TextMatrix(RowON, ColON) = txtText.Text
End Sub



Private Sub cmbSelect_LostFocus()
    'When the combo box looses focus, it becomes invisiable
    cmbSelect.Visible = False
    grProperties.TextMatrix(RowON, ColON) = cmbSelect.Text
End Sub
