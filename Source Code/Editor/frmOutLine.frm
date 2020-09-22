VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOutline 
   AutoRedraw      =   -1  'True
   Caption         =   "Object"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7590
   Icon            =   "frmOutLine.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   7590
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1920
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutLine.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutLine.frx":0896
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutLine.frx":0CEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutLine.frx":113E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   3480
      Width           =   6735
      Begin VB.CommandButton cmdAct 
         Caption         =   "Delete"
         Enabled         =   0   'False
         Height          =   350
         Index           =   4
         Left            =   2520
         TabIndex        =   14
         ToolTipText     =   "Remove the selected vertex or face"
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAct 
         Caption         =   "Save"
         Height          =   350
         Index           =   0
         Left            =   1320
         TabIndex        =   13
         ToolTipText     =   "Save the changes made to the object"
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAct 
         Caption         =   "Edge"
         Enabled         =   0   'False
         Height          =   350
         Index           =   1
         Left            =   5400
         TabIndex        =   12
         ToolTipText     =   "Allows you to create faces with more edges"
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAct 
         Caption         =   "Help"
         Height          =   350
         Index           =   2
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Get help on this window"
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAct 
         Caption         =   "New"
         Enabled         =   0   'False
         Height          =   350
         Index           =   3
         Left            =   3960
         TabIndex        =   10
         ToolTipText     =   "Allows you to create faces with more edges"
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Faces"
      Height          =   2775
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   5535
      Begin VB.TextBox txtFace 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2040
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   2280
         Visible         =   0   'False
         Width           =   975
      End
      Begin MSFlexGridLib.MSFlexGrid gdFace 
         Height          =   2415
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   4260
         _Version        =   393216
         Rows            =   1
         ScrollTrack     =   -1  'True
         Appearance      =   0
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Vertecies"
      Height          =   2775
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   5535
      Begin VB.TextBox txtVertex 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   1320
         Visible         =   0   'False
         Width           =   975
      End
      Begin MSFlexGridLib.MSFlexGrid gdVertex 
         Height          =   2415
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   4260
         _Version        =   393216
         Rows            =   0
         Cols            =   5
         FixedRows       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   1
         Appearance      =   0
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "General"
      Height          =   2775
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   6375
      Begin MSFlexGridLib.MSFlexGrid gdGeneral 
         Height          =   2535
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   4471
         _Version        =   393216
         Rows            =   0
         FixedRows       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   1
         Appearance      =   0
      End
   End
   Begin MSComctlLib.TabStrip TabView 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   5741
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Vertecies"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Faces"
            ImageVarType    =   2
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOutline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Model As clsFile
Dim GridX As Integer, GridY As Integer, ObjectOn As String

'#####################################################################
'#                                                                   #
'#                            FrmOutline                             #
'#                                                                   #
'# This form displays the details on an object in its most basic     #
'# state, with all the faces and vertecies displayed in a grid. You  #
'# can alter the values in the grids, and add new records to the     #
'# grid. It may not be easy to see what your doing, but gives you    #
'# the most control, if you want to use it.                          #
'#                                                                   #
'#####################################################################


Public Sub RunAtStart(AssignFile As clsFile)
    'This is how the form is started
    Set Model = AssignFile
    If Model.Geometery.CountSelected = 0 Then
        MsgBox "You must select an object before you can use this function", vbInformation
        Exit Sub
    End If
    ObjectOn = Model.Geometery(Model.Geometery.FirstSelectedObject).Key
    LoadOutLineDetails
    Show vbModal
End Sub

Private Sub LoadOutLineDetails()
    Dim GetVertex As clsVertex, n As Integer, ReadFace As String
    Dim GetFace As clsFace, GetEdge As clsEdge, m As Integer
    With Model.Geometery(ObjectOn)
        Caption = "Object Makeup : Object [" & ObjectOn & "]"
        gdGeneral.ColWidth(0) = 1500: gdGeneral.ColWidth(1) = 2300
        gdFace.ColWidth(0) = 500
        gdVertex.ColWidth(1) = 1000: gdVertex.ColWidth(2) = 1000: gdVertex.ColWidth(3) = 1000
        gdVertex.AddItem "" & vbTab & "X" & vbTab & "Y" & vbTab & "Z" & vbTab & "Target"
        For n = 2 To gdFace.Cols - 1
            gdFace.ColWidth(n) = 500
            gdFace.TextMatrix(0, n) = "Faces"
        Next n
        n = 0
        For Each GetVertex In .Vertex
            n = n + 1
            gdVertex.AddItem n & vbTab & GetVertex.x & vbTab & GetVertex.y & vbTab & GetVertex.z & vbTab & GetVertex.TargetName
        Next GetVertex
        gdVertex.FixedRows = 1: gdFace.TextMatrix(0, 1) = "Poly Count"
        gdFace.MergeCells = flexMergeFree: n = 0
        For Each GetFace In .Face
            n = n + 1
            m = 0
            ReadFace = n & vbTab & GetFace.Edge.Count
            For Each GetEdge In GetFace.Edge
                m = m + 1
                If m > gdFace.Cols - 1 Then gdFace.Cols = m + 2
                ReadFace = ReadFace & vbTab & GetEdge.Vertex
            Next GetEdge
            gdFace.AddItem ReadFace
            ReadFace = ""
        Next GetFace
        gdFace.MergeRow(0) = True
        For n = 2 To gdFace.Cols - 1
            gdFace.ColWidth(n) = 400: gdFace.TextMatrix(0, n) = "Edge"
        Next n
        gdGeneral.AddItem "Colour" & vbTab & .Colour
        gdGeneral.AddItem "Face Count" & vbTab & .Face.Count
        gdGeneral.AddItem "Vertex Count" & vbTab & .Vertex.Count
    End With
End Sub

Private Sub cmdACT_Click(index As Integer)
    'This contols all the buttons on the form
    Dim n As Integer
    Select Case index
        Case 1
            'This increases the number of edges that you can have
            If gdFace.Rows = 52 Then
                MsgBox "You can only have 50 edges per face", vbExclamation
                Exit Sub
            End If
            gdFace.Cols = gdFace.Cols + 1
            gdFace.TextMatrix(0, gdFace.Cols - 1) = "Edge"
            gdFace.ColWidth(gdFace.Cols - 1) = 500
        
        Case 2
            Am8.ShowHelp "Object Makeup Window"
        
        Case 3
            'This increases the number of rows in either of the grids, Ie another
            'vertex or another face
            If Frame(1).Visible = True Then
                gdVertex.Rows = gdVertex.Rows + 1
                gdVertex.TextMatrix(gdVertex.Rows - 1, 0) = gdVertex.Rows - 1
                gdGeneral.TextMatrix(2, 1) = gdGeneral.TextMatrix(2, 1) + 1
            End If
            If Frame(2).Visible = True Then
                gdFace.Rows = gdFace.Rows + 1
                gdFace.TextMatrix(gdFace.Rows - 1, 0) = gdFace.Rows - 1
                gdGeneral.TextMatrix(1, 1) = gdGeneral.TextMatrix(1, 1) + 1
            End If
    
        Case 4
            If Frame(1).Visible = True Then
                If MsgBox("Are you sure you want to remove vertex " & gdVertex.RowSel & "?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
                Model.Geometery(ObjectOn).Vertex.Remove (gdVertex.RowSel - 1)
                gdVertex.RemoveItem gdVertex.RowSel
                gdGeneral.TextMatrix(2, 1) = gdGeneral.TextMatrix(2, 1) - 1
                For n = 1 To gdVertex.Rows - 1: gdVertex.TextMatrix(n, 0) = n: Next n
            Else
                If MsgBox("Are you sure you want to remove face " & gdFace.RowSel & "?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
                Model.Geometery(ObjectOn).Face.Remove (gdFace.RowSel)
                gdFace.RemoveItem gdFace.RowSel
                gdGeneral.TextMatrix(1, 1) = gdGeneral.TextMatrix(1, 1) - 1
                For n = 1 To gdFace.Rows - 1: gdFace.TextMatrix(n, 0) = n: Next n
            End If
    End Select
End Sub

Private Sub Form_Resize()
    'This resizes all the objects on the form to fit the window
    Dim n As Integer, sHight As Integer
    On Error Resume Next
    sHight = ScaleHeight - Frame1.Height
    Frame1.Top = sHight: Frame1.Width = ScaleWidth
    gdFace.Width = ScaleWidth - 625: gdFace.Height = sHight - 875
    TabView.Width = ScaleWidth - 200: TabView.Height = sHight - 200
    gdVertex.Width = ScaleWidth - 625: gdVertex.Height = sHight - 875
    gdGeneral.Width = ScaleWidth - 625: gdGeneral.Height = sHight - 875
    For n = 0 To 3
        Frame(n).BorderStyle = 0
        Frame(n).Height = sHight - 700
        Frame(n).Width = ScaleWidth - 450
    Next n
    cmdAct(1).Left = Frame1.Width - cmdAct(1).Width - 80
    cmdAct(3).Left = Frame1.Width - cmdAct(3).Width - 1250
    cmdAct(4).Left = Frame1.Width - cmdAct(3).Width - 2430
End Sub

Private Sub gdFace_DblClick()
    'When you double click on the grids, this displays a text box over the cell,
    'so you can enter new data into it
    If gdFace.ColSel = 1 Then Exit Sub
    txtFace.Top = gdFace.RowPos(gdFace.RowSel) + gdFace.Top
    txtFace.Left = gdFace.ColPos(2) + gdFace.Left + (gdFace.ColWidth(gdFace.ColSel) * (gdFace.ColSel - 2))
    txtFace.Height = gdFace.RowHeight(gdFace.RowSel)
    txtFace.Width = gdFace.ColWidth(gdFace.ColSel) + 17
    txtFace = gdFace.TextMatrix(gdFace.RowSel, gdFace.ColSel)
    GridX = gdFace.ColSel: GridY = gdFace.RowSel
    txtFace.Visible = True
    txtFace.SetFocus
End Sub

Private Sub gdVertex_DblClick()
    'When you double click on the grids, this displays a text box over the cell,
    'so you can enter new data into it
    txtVertex.Top = gdVertex.RowPos(gdVertex.RowSel) + gdVertex.Top
    txtVertex.Left = gdVertex.ColPos(gdVertex.ColSel) + gdVertex.Left
    txtVertex.Height = gdVertex.RowHeight(gdVertex.RowSel)
    txtVertex.Width = gdVertex.ColWidth(gdVertex.ColSel) + 17
    txtVertex = gdVertex.TextMatrix(gdVertex.RowSel, gdVertex.ColSel)
    GridX = gdVertex.ColSel: GridY = gdVertex.RowSel
    txtVertex.Visible = True
    txtVertex.SetFocus
End Sub

Private Sub TabView_Click()
    'When you click on the tab strip, this shows the right form, enables the
    'right text buttons, and sets the caption on the buttons
    cmdAct(1).Enabled = False
    cmdAct(3).Enabled = False
    cmdAct(4).Enabled = False
    Dim n As Integer
    For n = 0 To Frame.Count - 1
        Frame(n).Visible = False
    Next n
    If TabView.SelectedItem.index = 2 Then
        cmdAct(1).Enabled = False
        cmdAct(3).Enabled = True
        cmdAct(4).Enabled = True
        cmdAct(3).Caption = "Add vertex"
        cmdAct(3).ToolTipText = "Create a new vertex in the object"
    End If
    If TabView.SelectedItem.index = 3 Then
        cmdAct(1).Enabled = True
        cmdAct(3).Enabled = True
        cmdAct(4).Enabled = True
        cmdAct(3).Caption = "Add face"
        cmdAct(3).ToolTipText = "Create a new face in the object"
    End If
    Frame(TabView.SelectedItem.index - 1).Visible = True
End Sub

Private Sub txtFace_KeyPress(KeyAscii As Integer)
    'If you press enter, then set the changes o the grid, and hide the text box
    If KeyAscii = 13 Then
        gdFace.TextMatrix(GridY, GridX) = txtFace
        txtFace.Visible = False
    End If
End Sub

Private Sub txtFace_LostFocus()
    'If you just click outside the text box, its like pressing cancel
    txtFace.Visible = False
End Sub

Private Sub txtVertex_KeyPress(KeyAscii As Integer)
    'If you press enter, then set the changes o the grid, and hide the text box
    If KeyAscii = 13 Then
        gdVertex.TextMatrix(GridY, GridX) = txtVertex
        txtVertex.Visible = False
    End If
End Sub

Private Sub txtVertex_LostFocus()
    'If you just click outside the text box, its like pressing cancel
    txtVertex.Visible = False
End Sub
