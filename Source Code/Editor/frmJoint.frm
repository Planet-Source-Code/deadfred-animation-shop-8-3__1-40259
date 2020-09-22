VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmJoint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Joint Properties"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7410
   Icon            =   "frmJoint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList greyicons 
      Left            =   600
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJoint.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJoint.frx":0896
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList EditIcons 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJoint.frx":0BB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJoint.frx":1006
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView Joints 
      Height          =   4575
      Left            =   240
      TabIndex        =   25
      Top             =   600
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   8070
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   441
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "EditIcons"
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog getColour 
      Left            =   1680
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Flags           =   1
   End
   Begin VB.CommandButton cmdAct 
      Caption         =   "Okay"
      Height          =   350
      Index           =   1
      Left            =   5040
      TabIndex        =   24
      ToolTipText     =   "Click to confim the changes made"
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton cmdAct 
      Caption         =   "Cancel"
      Height          =   350
      Index           =   2
      Left            =   6240
      TabIndex        =   23
      ToolTipText     =   "Click to close this window without saving your changes"
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton cmdAct 
      Caption         =   "Help"
      Height          =   350
      Index           =   0
      Left            =   120
      TabIndex        =   22
      ToolTipText     =   "Click to get help on this window"
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Frame fmONModel 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4575
      Left            =   3480
      TabIndex        =   26
      Top             =   600
      Width           =   3735
      Begin VB.Label lblCountSelect 
         Alignment       =   2  'Center
         Caption         =   "Select a joint from the list to edit it"
         Height          =   255
         Left            =   0
         TabIndex        =   29
         Top             =   1440
         Width           =   3735
      End
      Begin VB.Label lblShowCount 
         Alignment       =   2  'Center
         Caption         =   "Select a joint from the list to edit it"
         Height          =   255
         Left            =   0
         TabIndex        =   28
         Top             =   1200
         Width           =   3735
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Select a joint from the list to edit it"
         Height          =   255
         Left            =   0
         TabIndex        =   27
         Top             =   2280
         Width           =   3735
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   4575
      Index           =   0
      Left            =   3600
      TabIndex        =   1
      Top             =   600
      Width           =   3615
      Begin VB.PictureBox pkColour 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   960
         ScaleHeight     =   225
         ScaleWidth      =   2025
         TabIndex        =   11
         ToolTipText     =   "The sets the colour of the selected joint"
         Top             =   2640
         Width           =   2055
      End
      Begin VB.CheckBox chkGray 
         Alignment       =   1  'Right Justify
         Caption         =   "Greyed"
         Height          =   255
         Left            =   180
         TabIndex        =   9
         ToolTipText     =   "This grays out the selected joint"
         Top             =   2160
         Width           =   960
      End
      Begin VB.CheckBox chkHide 
         Alignment       =   1  'Right Justify
         Caption         =   "Hidden"
         Height          =   255
         Left            =   180
         TabIndex        =   8
         ToolTipText     =   "This hides the selected joint"
         Top             =   1800
         Width           =   960
      End
      Begin VB.CheckBox chkLock 
         Alignment       =   1  'Right Justify
         Caption         =   "Locked"
         Height          =   255
         Left            =   180
         TabIndex        =   7
         ToolTipText     =   "This locks the selected joint"
         Top             =   1440
         Width           =   960
      End
      Begin VB.ComboBox cmbTarget 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "This sets the target of this joint"
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   960
         TabIndex        =   3
         ToolTipText     =   "This sets the name of this joint"
         Top             =   480
         Width           =   2055
      End
      Begin VB.CommandButton cmdEntity 
         Caption         =   "Entity"
         Height          =   350
         Left            =   840
         TabIndex        =   2
         ToolTipText     =   "Click to set the entity properties of this joint"
         Top             =   3840
         Width           =   1695
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         X1              =   120
         X2              =   3360
         Y1              =   3495
         Y2              =   3495
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000015&
         X1              =   120
         X2              =   3360
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Colour"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Target"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Name"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   495
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   4695
      Index           =   1
      Left            =   3360
      TabIndex        =   12
      Top             =   480
      Visible         =   0   'False
      Width           =   3855
      Begin VB.ComboBox cmdSetup 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   15
         ToolTipText     =   "Contains a list of skeliton profiles to choose from"
         Top             =   2040
         Width           =   3615
      End
      Begin VB.PictureBox PreView 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   1815
         Left            =   120
         ScaleHeight     =   117
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   237
         TabIndex        =   13
         ToolTipText     =   "A Preview of this skeletons profile"
         Top             =   120
         Width           =   3615
      End
      Begin VB.Frame HideCustom 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1935
         Left            =   240
         TabIndex        =   14
         Top             =   2520
         Width           =   3255
         Begin MSComctlLib.Slider JWidth 
            Height          =   255
            Index           =   2
            Left            =   600
            TabIndex        =   16
            ToolTipText     =   "This sets the level of faces used to cover the skeliton"
            Top             =   1320
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   1
            Min             =   1
            Max             =   20
            SelStart        =   5
            Value           =   5
         End
         Begin MSComctlLib.Slider JWidth 
            Height          =   1215
            Index           =   3
            Left            =   1320
            TabIndex        =   17
            ToolTipText     =   "This sets the amount of curve in the skeliton"
            Top             =   0
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   2143
            _Version        =   393216
            Orientation     =   1
            LargeChange     =   1
            Min             =   -20
            Max             =   20
            TickStyle       =   2
            TickFrequency   =   4
         End
         Begin MSComctlLib.Slider JWidth 
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   18
            ToolTipText     =   "The sets the width of the skeleton at the far end of the link"
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   1
            Max             =   20
            TickFrequency   =   2
         End
         Begin MSComctlLib.Slider JWidth 
            Height          =   255
            Index           =   1
            Left            =   2040
            TabIndex        =   19
            ToolTipText     =   "This sets the width of the skeleton at the end next to the joint"
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   1
            Max             =   20
            TickFrequency   =   2
         End
         Begin MSComctlLib.Slider JWidth 
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   20
            ToolTipText     =   "The sets the position of the link on the far end of the skeliton"
            Top             =   720
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   1
            Min             =   1
            SelStart        =   1
            TickFrequency   =   2
            Value           =   1
         End
         Begin MSComctlLib.Slider JWidth 
            Height          =   255
            Index           =   5
            Left            =   2040
            TabIndex        =   21
            ToolTipText     =   "The sets the position of the link on the end of the skeliton near the joint"
            Top             =   720
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   1
            Min             =   1
            SelStart        =   1
            TickFrequency   =   2
            Value           =   1
         End
      End
   End
   Begin MSComctlLib.TabStrip TabMain 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   9128
      HotTracking     =   -1  'True
      ImageList       =   "greyicons"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Joints"
            Object.ToolTipText     =   "Allows you to set the names, targets and edit properties of the selected joint"
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Skeliton"
            Object.ToolTipText     =   "Allows you to set the skelitons 3D profile"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmJoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Model As clsFile

'#####################################################################
'#                                                                   #
'#   This form originally had all the compile code in it, but its    #
'#    been moved to the Timei form so that life is easier. This      #
'#     now does nothing other that allow the user to select the      #
'#                      tick boxes they want.                        #
'#                                                                   #
'#####################################################################


Public Sub RunAtStart(AssignedFile As clsFile)
    'This starts the form running
    Dim Jm As clsJoint
    If AssignedFile.Joint.CountSelected = 0 Then MsgBox amMustSelectJoint, vbInformation: Exit Sub
    Set Model = AssignedFile
    cmbTarget.AddItem ""
    For Each Jm In Model.Joint
        cmbTarget.AddItem Jm.Name
    Next Jm
    Model.Joint.DisplayTreeInWindow Joints, , True
    Joints.Nodes(1).Selected = True
    cmdSetup.ListIndex = 0
    Show vbModal
End Sub

Private Sub chkGray_Click()
    Model.Joint(Joints.SelectedItem.Key).Grayed = chkGray
End Sub

Private Sub chkHide_Click()
    Model.Joint(Joints.SelectedItem.Key).Hidden = chkHide
End Sub

Private Sub chkLock_Click()
    Model.Joint(Joints.SelectedItem.Key).Locked = chkLock
End Sub

Private Sub cmdEntity_Click()
    frmEntity.RunAtStart Model
End Sub

Private Sub Form_Load()
    'This draws the profile and adds the preset items to the dowpdown list
    cmdSetup.Clear
    cmdSetup.AddItem "Blank"
    cmdSetup.AddItem "Straight"
    cmdSetup.AddItem "Wide"
    cmdSetup.AddItem "Balloon"
    cmdSetup.AddItem "Point"
    cmdSetup.AddItem "Reverse Point"
    cmdSetup.AddItem "Strech"
    cmdSetup.AddItem "Custom..."
    lblShowCount = "Total Joints : " & Model.Joint.CountChildren
    lblCountSelect = "Selected Joints : " & Model.Joint.CountSelected
End Sub

Private Sub cmdACT_Click(Index As Integer)
    'This controls the three buttons along the bottom of the form
    Select Case Index
        Case 0: Am8.ShowHelp "Joint Properties window"
        Case 1:
            Model.Joint.DisplayTreeInWindow frmMain.Joints
            Unload Me
        Case 2: Unload Me
    End Select
End Sub

Private Sub cmdSetup_Click()
    'When you click on the profile dropdown list, this resets the slide bars, or revals the
    'settings frame, if that is what the user requested.
    If cmdSetup.Text = "Custom..." Then HideCustom.Visible = True Else HideCustom.Visible = False
    Select Case LCase(cmdSetup.Text)
        Case "blank":           SetPreSet 0, 0, 0, 0, 0, 0
        Case "straight":        SetPreSet 5, 5, 1, 10, 0, 1
        Case "wide":            SetPreSet 12, 12, 2, 9, 0, 1
        Case "balloon":         SetPreSet 5, 5, 1, 10, -10, 20
        Case "point":           SetPreSet 2, 20, 2, 9, 0, 1
        Case "reverse point":   SetPreSet 20, 2, 2, 9, 0, 1
        Case "strech":          SetPreSet 10, 10, 2, 9, 4, 20
    End Select
    If Joints.SelectedItem.Key <> "BaseJoint" Then
        Model.Joint(Joints.SelectedItem.Key).JointProfileIndex = cmdSetup.ListIndex
    End If
End Sub

Private Sub SetPreSet(WidthA As Integer, WidthB As Integer, PosA As Integer, PosB As Integer, Curve As Integer, Detail As Integer)
    'This is a quick way of setting all the slide bars to new values. You pass the values you
    'want into the routine, and they are all set correctly.
    JWidth(0) = WidthA
    JWidth(1) = WidthB
    JWidth(3) = Curve
    JWidth(2) = Detail
    JWidth(4) = PosA
    JWidth(5) = PosB
    RedrawPreview
End Sub

Private Sub Joints_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim n As Integer
    cmbTarget.ListIndex = -1
    If Node.Key = "BaseJoint" Then
        fmONModel.Visible = True
    Else
        fmONModel.Visible = False
        pkColour.BackColor = Model.Joint(Node.Key).Colour
        txtName = Model.Joint(Node.Key).Name
        chkGray = Abs(Int(Model.Joint(Node.Key).Grayed))
        chkLock = Abs(Int(Model.Joint(Node.Key).Locked))
        chkHide = Abs(Int(Model.Joint(Node.Key).Hidden))
        For n = 0 To 5
            JWidth(n) = Model.Joint(Node.Key).JProf(n)
        Next n
        cmdSetup.ListIndex = Model.Joint(Node.Key).JointProfileIndex
        For n = 1 To Model.Joint.CountChildren
            If Model.Joint(Node.Key).Target = Model.Joint(n).Key Then
                cmbTarget.ListIndex = n
            End If
        Next n
        RedrawPreview
    End If
End Sub

Private Sub JWidth_Click(Index As Integer)
    'This redraws the preview window as you drag the slide bars around
    Model.Joint(Joints.SelectedItem.Key).JProf(Index) = JWidth(Index)
    RedrawPreview
End Sub

Private Sub JWidth_Scroll(Index As Integer)
    'This redraws the preview window as you drag the slide bars around
    RedrawPreview
End Sub

Private Sub pkColour_DblClick()
    'This simply allows you to change the colour of the selected joint
    On Error GoTo PressedCancel
    GetColour.Flags = cdlCCFullOpen
    GetColour.DialogTitle = "Select Joint Colour"
    GetColour.ShowColor
    pkColour.BackColor = GetColour.Color
    Model.Joint(Joints.SelectedItem.Key).Colour = GetColour.Color
PressedCancel:
End Sub

Private Sub TabMain_Click()
    'When you click on the tabs, this displays the desired frame
    Dim n As Integer
    For n = 0 To 1: Frame(n).Visible = False: Next n
    Frame(TabMain.SelectedItem.Index - 1).Visible = True
End Sub

Private Sub RedrawPreview()
    'This draws the profile of the skeliton, using the parameters held in the slide bars
    Dim Xon As Single, YOn As Single, x1 As Single, y1 As Single
    Dim XX As Single, YY As Single, HeightA As Single, HeightB As Single
    Dim Curve As Single, Vxx1 As Integer, n As Integer, Vxx2 As Integer
    PreView.Cls
    XX = PreView.ScaleWidth / 2:    YY = PreView.ScaleHeight / 2
    HeightA = Me.JWidth(0) * 2:     HeightB = Me.JWidth(1) * 2
    Vxx1 = (XX * 0.9) - (XX / 11 * (JWidth(4)))
    Vxx2 = (XX * 0.9) - (XX / 11 * (11 - JWidth(5)))
    y1 = (HeightA - HeightB) / JWidth(2)
    x1 = (Vxx1 + Vxx2) / JWidth(2)
    Xon = XX - Vxx1: YOn = YY - HeightA
    PreView.ForeColor = vbRed
    PreView.Line (XX - 100, YY)-(XX - Vxx1, YY - HeightA)
    PreView.Line (XX + 100, YY)-(XX + Vxx2, YY - HeightB)
    PreView.Line (XX - 100, YY)-(XX - Vxx1, YY + HeightA)
    PreView.Line (XX + 100, YY)-(XX + Vxx2, YY + HeightB)
    For n = 1 To JWidth(2)
        Curve = (-(1 + JWidth(2)) * 0.5) + n
        Curve = -Curve * JWidth(3) / (JWidth(2) ^ 2) * 20
        PreView.Line (Xon, YOn)-(Xon + x1, YOn + y1 + Curve)
        PreView.Line (Xon, -YOn + YY + YY)-(Xon + x1, -(YOn + y1 + Curve) + YY + YY)
        Xon = Xon + x1
        YOn = YOn + y1 + Curve
    Next n
    PreView.ForeColor = RGB(100, 155, 100)
    PreView.Line (XX - Vxx1, YY - HeightA)-(XX - Vxx1, YY + HeightA)
    PreView.Line (XX + Vxx2, YY - HeightB)-(XX + Vxx2, YY + HeightB)
    PreView.ForeColor = vbBlue
    PreView.Line (XX - 100, YY)-(XX + 100, YY)
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then txtName = Model.Joint(Joints.SelectedItem.Key).Name
    If KeyAscii = 13 Then
        Model.Joint(Joints.SelectedItem.Key).Name = txtName
        Model.Joint.DisplayTreeInWindow Joints, , True
    End If
End Sub

Private Sub txtName_LostFocus()
    Model.Joint(Joints.SelectedItem.Key).Name = txtName
    Model.Joint.DisplayTreeInWindow Joints, , True
End Sub
