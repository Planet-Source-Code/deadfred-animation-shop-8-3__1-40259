VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmObject 
   Caption         =   "Object Editor"
   ClientHeight    =   7665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10335
   Icon            =   "frmObject.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7665
   ScaleWidth      =   10335
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList TabIcon 
      Left            =   960
      Top             =   6240
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
            Picture         =   "frmObject.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmObject.frx":0896
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmObject.frx":0CEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmObject.frx":113E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdAct 
      Caption         =   "Help"
      Height          =   350
      Index           =   3
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Click to get help on using this window"
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdAct 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   350
      Index           =   4
      Left            =   3000
      TabIndex        =   1
      ToolTipText     =   "Click to close this window without saving your changes"
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdAct 
      Caption         =   "Okay"
      Height          =   350
      Index           =   5
      Left            =   4800
      TabIndex        =   0
      ToolTipText     =   "Click to close this window and save your changes"
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   4455
      Index           =   1
      Left            =   240
      TabIndex        =   12
      Top             =   480
      Width           =   6135
      Begin Project1.Engine Engine 
         Height          =   2295
         Index           =   1
         Left            =   120
         TabIndex        =   24
         Top             =   120
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   4048
      End
      Begin VB.Frame frmTools 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   3375
         Left            =   3960
         TabIndex        =   13
         Top             =   360
         Width           =   1935
         Begin VB.OptionButton ckMode 
            Caption         =   "New Face"
            Height          =   255
            Index           =   5
            Left            =   360
            TabIndex        =   30
            Top             =   1800
            Width           =   1215
         End
         Begin VB.OptionButton ckMode 
            Caption         =   "New Vertex"
            Height          =   255
            Index           =   4
            Left            =   360
            TabIndex        =   29
            Top             =   1560
            Width           =   1215
         End
         Begin VB.OptionButton ckMode 
            Caption         =   "Reverse face"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   14
            Top             =   720
            Width           =   1455
         End
         Begin VB.OptionButton ckMode 
            Caption         =   "To Triangle"
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   28
            Top             =   1320
            Width           =   1335
         End
         Begin VB.OptionButton ckMode 
            Caption         =   "Rotate"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   16
            Top             =   120
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton ckMode 
            Caption         =   "Delete face"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   15
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Fragment"
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   1080
            Width           =   855
         End
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   4455
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   6135
      Begin VB.CommandButton cmdAct 
         Caption         =   "&Reverse face"
         Height          =   375
         Index           =   6
         Left            =   1440
         TabIndex        =   4
         ToolTipText     =   "Change the direction that the face is pointing"
         Top             =   3840
         Width           =   1455
      End
      Begin VB.ListBox lstFaceOrder 
         Height          =   2865
         IntegralHeight  =   0   'False
         Left            =   5040
         TabIndex        =   11
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdAct 
         Caption         =   "&Deselect all"
         Height          =   375
         Index           =   0
         Left            =   3000
         TabIndex        =   10
         ToolTipText     =   "Unselect all vertecies"
         Top             =   3840
         Width           =   1455
      End
      Begin VB.CommandButton cmdAct 
         Caption         =   "&Create face"
         Height          =   375
         Index           =   1
         Left            =   4560
         TabIndex        =   9
         ToolTipText     =   "Joint the vertecies together to make a new face"
         Top             =   3840
         Width           =   1455
      End
      Begin VB.CommandButton cmdUpdown 
         Height          =   495
         Index           =   0
         Left            =   5040
         Picture         =   "frmObject.frx":1596
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Move a vertex up the list"
         Top             =   3120
         Width           =   375
      End
      Begin VB.CommandButton cmdUpdown 
         Height          =   495
         Index           =   1
         Left            =   5640
         Picture         =   "frmObject.frx":18A0
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Move a vertex down the list"
         Top             =   3120
         Width           =   375
      End
      Begin VB.CheckBox ckInsert 
         Caption         =   "Show numbers"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "Shows the number of eac joint"
         Top             =   3720
         Width           =   1695
      End
      Begin VB.CheckBox ckInsert 
         Caption         =   "Create two sides face"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Creates two faces back to back, that can be seen from either side"
         Top             =   4080
         Width           =   2415
      End
      Begin Project1.Engine Engine 
         Height          =   2295
         Index           =   2
         Left            =   120
         TabIndex        =   25
         Top             =   120
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   4048
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3135
      Index           =   4
      Left            =   240
      TabIndex        =   26
      Top             =   480
      Visible         =   0   'False
      Width           =   3495
      Begin Project1.DXEngine PreView 
         Height          =   1455
         Left            =   0
         TabIndex        =   27
         Top             =   0
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   2566
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   4455
      Index           =   3
      Left            =   240
      TabIndex        =   17
      Top             =   480
      Visible         =   0   'False
      Width           =   6135
      Begin Project1.Engine Engine 
         Height          =   3495
         Index           =   3
         Left            =   120
         TabIndex        =   23
         Top             =   120
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   6165
      End
      Begin VB.CommandButton cmdAct 
         Caption         =   "&Slice"
         Height          =   375
         Index           =   2
         Left            =   2160
         TabIndex        =   21
         ToolTipText     =   "Click to remove all parts of the object below the clip line"
         Top             =   3960
         Width           =   1455
      End
      Begin VB.CheckBox chslOp 
         Caption         =   "Don't Compress"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   20
         ToolTipText     =   "Stops the compress function taking place, which can be slow"
         Top             =   3960
         Width           =   1815
      End
      Begin VB.CheckBox chslOp 
         Caption         =   "Don't Re-aline"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   19
         ToolTipText     =   "Stops the object being rotated as well as sliced "
         Top             =   4200
         Width           =   1695
      End
      Begin VB.CheckBox chslOp 
         Alignment       =   1  'Right Justify
         Caption         =   "Smooth"
         Height          =   255
         Index           =   0
         Left            =   4920
         TabIndex        =   18
         ToolTipText     =   "Stops the compress function taking place, which can be slow"
         Top             =   3960
         Width           =   1095
      End
   End
   Begin MSComctlLib.TabStrip ViewTab 
      Height          =   4935
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   8705
      HotTracking     =   -1  'True
      ImageList       =   "TabIcon"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit Faces"
            Object.Tag             =   "1"
            Object.ToolTipText     =   "Remove or reverse existing faces in the selected objects"
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Insert Faces"
            Object.Tag             =   "2"
            Object.ToolTipText     =   "Add new faces when you have a single object selected"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tumble Editor"
            Object.Tag             =   "3"
            Object.ToolTipText     =   "Rotate the objects and remove parts below a certain line"
            ImageVarType    =   2
            ImageIndex      =   3
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Preview"
            Object.ToolTipText     =   "See a preview of the changes using DirectX"
            ImageVarType    =   2
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Angle1 As Single, Angle2 As Single, Angle3 As Single
Dim OldX As Integer, OldY As Integer, Rasie As Single
Dim Model As clsFile

Public Sub RunAtStart(AssignedFile As clsFile)
    Dim n As Integer
    If AssignedFile.Geometery.CountSelected = 0 Then MsgBox amMustSelectObject, vbInformation: Exit Sub
    Set Model = AssignedFile
    PreView.AssignDXEngineTo Model
    PreView.pSelectedOnly = True
    For n = 1 To 3
        Engine(n).AssignEngineTo Model
        Engine(n).pSelectedOnly = True
        Engine(n).pAutoRotate = True
        Engine(n).pDrawObjects = True
        Engine(n).pPerspecitve = True
        Engine(n).pAllFace = True
    Next n
    Engine(3).pClipFaces = True
    Engine(2).pHightlightVertex = True
    Engine(1).pHighlightFace = True
    Engine(2).pDrawEdgePreview = True
    Dim GSize As Integer
    Angle1 = 0: Angle2 = 0: Angle3 = 0
    Form_Resize
    Model.Geometery.DeselectAllVertecies
    Engine(2).pAutoZoom = True
    Show vbModal
End Sub

Private Sub Engine_MouseDown(Index As Integer, x As Single, y As Single, Button As Integer, Shift As Integer)
    Dim n As Integer, VertOver As Integer
    Select Case Index
    
        Case 1
            Engine(1).BeginRotate Int(x), Int(y)
            Engine(1).RefreshView
            Engine(1).BeginRotate 0, 0
            If ckMode(1) = True Then
                If Engine(1).FaceOver <> 0 Then Model.Geometery(Model.Geometery.FirstSelectedObject).Face.Remove Engine(1).FaceOver
            End If
            If ckMode(3) = True Then If Engine(1).FaceOver <> 0 Then Model.Geometery(Model.Geometery.FirstSelectedObject).FragmentFace Engine(1).FaceOver, 0
            If ckMode(4) = True Then If Engine(1).FaceOver <> 0 Then Model.Geometery(Model.Geometery.FirstSelectedObject).FragmentFace Engine(1).FaceOver, 1
            If ckMode(5) = True Then If Engine(1).FaceOver <> 0 Then Model.Geometery(Model.Geometery.FirstSelectedObject).FragmentFace Engine(1).FaceOver, 2, 0.5
            Engine(1).FaceOver = 0
            Engine(1).RefreshView
    
        Case 2
            Engine(2).BeginRotate Int(x), Int(y)
            Engine(2).RefreshView
            VertOver = Engine(2).VertexOver
            Engine(2).VertexOver = 0
            If VertOver <> 0 Then
                Select Case Model.Geometery(Model.Geometery.FirstSelectedObject).Vertex(VertOver).Selected
                    Case True
                        Model.Geometery(Model.Geometery.FirstSelectedObject).Vertex(VertOver).Selected = False
                        For n = 1 To lstFaceOrder.ListCount
                            If lstFaceOrder.List(n - 1) = VertOver Then lstFaceOrder.RemoveItem n - 1: Exit For
                        Next n
                    Case False
                        If lstFaceOrder.ListCount < 25 Then
                            Model.Geometery(Model.Geometery.FirstSelectedObject).Vertex(VertOver).Selected = True
                            lstFaceOrder.AddItem VertOver
                        Else
                            MsgBox "You have reached the maximum number of 25 edges", vbInformation
                        End If
                End Select
                Engine(2).BeginRotate 0, 0
                Engine(2).RefreshView
            End If
    End Select
End Sub

Private Sub Form_Resize()
    Dim n As Integer
    On Error Resume Next
    ViewTab.Width = ScaleWidth - 240
    ViewTab.Height = ScaleHeight - 640
    For n = 1 To 4
        Frame(n).Move ViewTab.ClientLeft, ViewTab.ClientTop, ViewTab.ClientWidth, ViewTab.ClientHeight
    Next n
    PreView.Move 120, 120, Frame(4).Width - 240, Frame(4).Height - 240
    Engine(1).Height = Frame(1).Height - 500
    Engine(1).Width = Frame(1).Width - frmTools.Width
    Engine(2).Height = Frame(2).Height - 900
    Engine(2).Width = Frame(2).Width - frmTools.Width + 500
    Engine(3).Width = Frame(3).Width - 200
    lstFaceOrder.Height = Engine(2).Height - 750
    ckInsert(0).Top = Engine(2).Height + 250
    ckInsert(1).Top = Engine(2).Height + 600
    Engine(3).Height = Frame(3).Height - 1000
    chslOp(0).Top = Engine(3).Height + 300
    chslOp(1).Top = Engine(3).Height + 300
    chslOp(2).Top = Engine(3).Height + 600
    chslOp(0).Left = Engine(3).Width - chslOp(0).Width + 100
    lstFaceOrder.Left = Frame(1).Width - lstFaceOrder.Width - 200
    frmTools.Left = Frame(1).Width - frmTools.Width - 50
    cmdUpdown(0).Left = lstFaceOrder.Left
    frmTools.Left = ViewTab.Width - frmTools.Width
    cmdUpdown(1).Left = lstFaceOrder.Left + lstFaceOrder.Width - cmdUpdown(0).Width
    cmdUpdown(0).Top = lstFaceOrder.Height + 350
    cmdUpdown(1).Top = lstFaceOrder.Height + 350
    For n = 0 To 2: cmdAct(n).Top = ScaleHeight - 1600: Next
    cmdAct(6).Top = ScaleHeight - 1600
    cmdAct(2).Left = (Frame(1).Width / 2) - (cmdAct(2).Width / 2)
    cmdAct(1).Left = Frame(1).Width - cmdAct(1).Width - 50
    cmdAct(0).Left = Frame(1).Width - (cmdAct(0).Width + 100) * 2
    cmdAct(6).Left = Frame(1).Width - (cmdAct(0).Width + 120) * 3
    For n = 3 To 5: cmdAct(n).Top = ScaleHeight - 450: Next
    cmdAct(5).Left = ScaleWidth - (cmdAct(5).Width + 100) * 2
    cmdAct(4).Left = ScaleWidth - 100 - cmdAct(4).Width
    DrawMe
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.ActiveForm.Tablet.Refresh
End Sub

Private Sub ViewTab_Click()
    Dim n As Integer
    If Model.Geometery.CountSelected <> 1 And ViewTab.SelectedItem.Index = 2 Then
        MsgBox "This tool can only be used when a single object is selected", vbExclamation
        ViewTab.Tabs(1).Selected = True
    End If
    For n = 1 To Frame.Count:   Frame(n).Visible = False: Next
    Frame(ViewTab.SelectedItem.Index).Visible = True
    DrawMe
End Sub

Private Sub cmdAct_Click(Index As Integer)
    Dim n As Integer, OriginalFace As Integer
    Select Case Index
    
        Case 0
            Model.Geometery.DeselectAllVertecies
            lstFaceOrder.Clear
            Engine(2).RefreshView
        
        Case 1
            Model.Geometery(Model.Geometery.FirstSelectedObject).Face.Add lstFaceOrder.ListCount
            For n = 0 To lstFaceOrder.ListCount - 1
                Model.Geometery(Model.Geometery.FirstSelectedObject).Face(Model.Geometery(Model.Geometery.FirstSelectedObject).Face.Count).Edge.Add lstFaceOrder.List(n)
            Next n
            lstFaceOrder.Clear
            Model.Geometery.DeselectAllVertecies
            Engine(2).RefreshView
    
        Case 3: Am8.ShowHelp "Object Properties Window"
        Case 4: Unload Me
        Case 5: Unload Me
        
        Case 6
            OriginalFace = lstFaceOrder.ListCount
            For n = lstFaceOrder.ListCount - 1 To 0 Step -1
                lstFaceOrder.AddItem (lstFaceOrder.List(n))
            Next n
            For n = 0 To OriginalFace - 1: lstFaceOrder.RemoveItem (0): Next n
            Engine(2).RefreshView
        
    End Select
End Sub

Private Sub cmdUpdown_Click(Index As Integer)
    Dim Onn As Integer, n As Integer, Temp As Integer
    Onn = -1
    For n = 0 To lstFaceOrder.ListCount - 1: If lstFaceOrder.Selected(n) = True Then Onn = n
    Next n
    If Onn = -1 Then Exit Sub
    If Index = 0 And Onn = 0 Then Exit Sub
    If Index = 1 And Onn + 1 = lstFaceOrder.ListCount Then Exit Sub
    If Index = 0 Then
        Temp = lstFaceOrder.List(Onn - 1)
        lstFaceOrder.RemoveItem (Onn - 1)
        lstFaceOrder.AddItem Temp, Onn
        lstFaceOrder.Selected(Onn - 1) = True
    End If
    If Index = 1 Then
        Temp = lstFaceOrder.List(Onn)
        lstFaceOrder.RemoveItem (Onn)
        lstFaceOrder.AddItem Temp, Onn + 1
        lstFaceOrder.Selected(Onn + 1) = True
    End If
    Engine(2).RefreshView
End Sub

Private Sub Engine_MouseDrag(Index As Integer, x As Integer, y As Integer, StartX As Integer, StartY As Integer, Button As Integer, Shift As Integer)
    If Button = 1 And Shift = 0 Then
        Angle2 = Angle2 - (x - StartX) * 10
        Angle1 = Angle1 - (y - StartY) * 10
        Angle1 = Angle1 Mod 3600
        Angle2 = Angle2 Mod 3600
        Angle3 = Angle3 Mod 3600
        DrawMe
    End If
End Sub

Private Function DrawMe()
    If ViewTab.SelectedItem.Index = 4 Then PreView.RefreshModel Else Engine(ViewTab.SelectedItem.Index).RefreshView
End Function


