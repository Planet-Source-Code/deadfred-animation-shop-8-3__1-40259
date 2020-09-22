VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSettings 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Animation Shop 8"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7080
   Icon            =   "frmSettings.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   7080
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1800
      Top             =   4920
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
            Picture         =   "frmSettings.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":0896
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":0CEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":113E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdAct 
      Caption         =   "Reset"
      Height          =   350
      Index           =   3
      Left            =   4680
      TabIndex        =   18
      ToolTipText     =   "Resets all settings to the factory defaults"
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdAct 
      Caption         =   "Okay"
      Height          =   350
      Index           =   2
      Left            =   3480
      TabIndex        =   3
      ToolTipText     =   "Click to close this window and save your changes"
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdAct 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   350
      Index           =   1
      Left            =   5880
      TabIndex        =   2
      ToolTipText     =   "Click to close this window without saving your changes"
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdAct 
      Caption         =   "Help"
      Height          =   350
      Index           =   0
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Click to get help on using this window"
      Top             =   4800
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog GetColour 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   4095
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   6615
      Begin ComCtl2.UpDown UpDown4 
         Height          =   285
         Left            =   5895
         TabIndex        =   15
         Top             =   1440
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   327681
         BuddyControl    =   "txtFileHist"
         BuddyDispid     =   196611
         OrigLeft        =   2160
         OrigTop         =   2040
         OrigRight       =   2400
         OrigBottom      =   2295
         Max             =   9
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtFileHist 
         Height          =   285
         Left            =   5280
         TabIndex        =   14
         Text            =   "5"
         ToolTipText     =   "Sets the maximum number of file stored in the file history"
         Top             =   1440
         Width           =   615
      End
      Begin ComCtl2.UpDown UpDown2 
         Height          =   285
         Left            =   5880
         TabIndex        =   12
         Top             =   960
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   327681
         Value           =   50
         BuddyControl    =   "txtLarge"
         BuddyDispid     =   196612
         OrigLeft        =   2160
         OrigTop         =   2160
         OrigRight       =   2400
         OrigBottom      =   2415
         Max             =   200
         Min             =   10
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtLarge 
         Height          =   285
         Left            =   5280
         TabIndex        =   11
         Text            =   "50"
         ToolTipText     =   "The size of the large grid lines"
         Top             =   960
         Width           =   615
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   5880
         TabIndex        =   8
         Top             =   480
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   327681
         BuddyControl    =   "txtGrid"
         BuddyDispid     =   196613
         OrigLeft        =   2160
         OrigTop         =   1680
         OrigRight       =   2400
         OrigBottom      =   1965
         Max             =   200
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtGrid 
         Height          =   285
         Left            =   5280
         TabIndex        =   7
         Text            =   "10"
         ToolTipText     =   "The size of the grid to snap objects to"
         Top             =   480
         Width           =   615
      End
      Begin VB.CheckBox ckONLeft 
         Caption         =   "Sidebar on left"
         Height          =   255
         Left            =   360
         TabIndex        =   26
         ToolTipText     =   "Moves the Sidebar onto the left hand side of the screen"
         Top             =   2880
         Width           =   1335
      End
      Begin VB.CheckBox chkNew 
         Caption         =   "Show New File diolog box at start"
         Height          =   255
         Left            =   360
         TabIndex        =   25
         ToolTipText     =   "Displays the New File diolog window when you start the program"
         Top             =   2400
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.CheckBox chkTip 
         Caption         =   "Show tips on start up"
         Height          =   255
         Left            =   360
         TabIndex        =   24
         ToolTipText     =   "Displays one of a list of helpful tips when the program starts"
         Top             =   1920
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   255
         Left            =   2520
         TabIndex        =   17
         ToolTipText     =   "Empty the file history"
         Top             =   1440
         Width           =   855
      End
      Begin VB.CheckBox chkFull 
         Caption         =   "Full file path in file history"
         Height          =   255
         Left            =   360
         TabIndex        =   16
         ToolTipText     =   "Dislpays the file path as well as the name in the file history"
         Top             =   1440
         Width           =   2055
      End
      Begin VB.CheckBox chkCenter 
         Caption         =   "Always center when changing views"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         ToolTipText     =   "Positions the selected objects in the centre of the screen when you change views"
         Top             =   480
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.CheckBox chkHighlightBox 
         Caption         =   "Highlight corners of selection box"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         ToolTipText     =   "Displays the corners and edges of the selected objects"
         Top             =   960
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "File history"
         Height          =   255
         Left            =   4440
         TabIndex        =   13
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Grid size"
         Height          =   255
         Left            =   4560
         TabIndex        =   10
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Span to grid size"
         Height          =   255
         Left            =   3960
         TabIndex        =   9
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   4095
      Index           =   2
      Left            =   240
      TabIndex        =   19
      Top             =   480
      Visible         =   0   'False
      Width           =   6615
      Begin VB.TextBox txtEnter 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   960
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CommandButton cmdLight 
         Caption         =   "Remove"
         Height          =   350
         Index           =   1
         Left            =   5520
         TabIndex        =   22
         ToolTipText     =   "Click to remove the current light style"
         Top             =   3720
         Width           =   1095
      End
      Begin VB.CommandButton cmdLight 
         Caption         =   "Add"
         Height          =   350
         Index           =   0
         Left            =   4320
         TabIndex        =   21
         ToolTipText     =   "Click to create a new light style"
         Top             =   3720
         Width           =   1095
      End
      Begin MSFlexGridLib.MSFlexGrid gdLights 
         Height          =   3495
         Left            =   0
         TabIndex        =   20
         Top             =   120
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   6165
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         SelectionMode   =   1
         Appearance      =   0
      End
   End
   Begin MSComctlLib.TabStrip ViewTab 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   8070
      HotTracking     =   -1  'True
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Editor"
            Object.ToolTipText     =   "Set different values to control the look of the editor"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Light Styles"
            Object.ToolTipText     =   "Change or add different light patterns"
            ImageVarType    =   2
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' #############################################################################
' #                                                                           #
' #   This settings form allows you to change the avaliable settigns for the  #
' # program. You can edit the light styles using the the Direct X view and    #
' #  set basic editor options. Settings are saved when you close the program  #
' #                                                                           #
' #############################################################################

Dim gdX As Integer, gdY As Integer, ColOver As Integer

Public Sub RunAtStart()
    'This code is used to display the settings window. It loads the details from the Am8 class, and sets up the flixi-grid
    'witht the correct headdings and light data
    Dim X As String, n As Integer, Entity As String, Lm As clsLightStyle
    txtGrid = Am8.SnapSize
    gdLights.TextMatrix(0, 1) = "Name"
    gdLights.TextMatrix(0, 2) = "Pattern"
    gdLights.ColWidth(0) = 0
    gdLights.ColWidth(2) = 5000
    txtFileHist = Am8.FileHistory.Lenght
    ckONLeft = IntBo(Am8.LeftSidebar)
    chkCenter = IntBo(Am8.AlwaysCenter)
    chkHighlightBox = IntBo(Am8.HighLightSection)
    chkFull = IntBo(Am8.FullPath)
    chkTip = IntBo(Am8.ShowTips)
    chkNew = IntBo(Am8.ShowNewWindow)
    For n = 1 To Am8.LightStyle.CountStyles
        gdLights.AddItem vbTab & Am8.LightStyle(n).Name & vbTab & Am8.LightStyle(n).Pattern
    Next n
    If gdLights.Rows > 2 Then gdLights.RemoveItem 1 Else gdLights.Enabled = False
    Show vbModal
End Sub

Private Sub cmdACT_Click(Index As Integer)
    'This controls the buttons along the bottom of the window. The Hekp and Cancel buttons are simple enought
    'The Ok button just writes the data back into the AM8 class, and the reset button deletes settings file,
    'and creates default light styles
    Dim n As Integer, m As Integer, txtTest As String, cOpen As Integer, sTag As String, cClose As Integer
    Select Case Index
        Case 0: Am8.ShowHelp "Settings Window"
        Case 1: Unload Me
        Case 2
            Am8.SnapSize = txtGrid
            Am8.FileHistory.Lenght = txtFileHist
            Am8.LightStyle.ClearStyles
            For n = 1 To gdLights.Rows - 1
                Am8.LightStyle.AddStyle gdLights.TextMatrix(n, 1), gdLights.TextMatrix(n, 2)
            Next n
            If ckONLeft = 1 Then frmMain.SideFrame.Align = 3 Else frmMain.SideFrame.Align = 4
            Am8.LeftSidebar = ckONLeft
            Am8.AlwaysCenter = chkCenter
            Am8.HighLightSection = chkHighlightBox
            Am8.FullPath = chkFull
            Am8.ShowTips = chkTip
            Am8.ShowNewWindow = chkNew
            Am8.LightStyle.UpdateLightList frmMain.lstLightStyle
            frmMain.UpdateHistoryMenu
            Unload Me
        Case 3
            If MsgBox(amRestoreSettigns, vbYesNo + vbQuestion) = vbYes Then
                Destroy App.Path & "data\settings.dat"
                Am8.LightStyle.ClearStyles
                Am8.LightStyle.AddStyle "Strobe", "AZ"
                Am8.LightStyle.AddStyle "Pulse", "ABCDEFGHIIIIIIIHGFEDCBAAAAAAA"
                Am8.LightStyle.AddStyle "Lightning", "AJJAJJDJSJJJFDFGDJJJGFDJJJJJ"
                Am8.LightStyle.AddStyle "Buzz", "AAASAASDAAAAAAAFGGGGAGGA"
                MsgBox amPleaseRestart, vbInformation
                Unload Me
            End If
    End Select
Exit Sub
NoCorrect:
    MsgBox amIncorrectSettings, vbCritical, "Warning!"
End Sub

Private Sub cmdClear_Click()
    If MsgBox(amClearHistory, vbQuestion + vbYesNo) = vbYes Then Am8.FileHistory.ClearHistory
End Sub

Private Sub cmdLight_Click(Index As Integer)
    Dim FileName As String
    Select Case Index
        Case 0
            FileName = Trim(InputBox(amNewLightName))
            If FileName = "" Then Exit Sub
            gdLights.AddItem vbTab & FileName
            If gdLights.TextMatrix(1, 1) = "" Then gdLights.RemoveItem 1
            gdLights.Enabled = True
        Case 1
            If MsgBox(gdLights.TextMatrix(gdLights.Row, 1) & vbNewLine & vbNewLine & amRemoveLight, vbQuestion + vbYesNo) = vbNo Then Exit Sub
            If gdLights.Rows = 2 Then
                gdLights.TextMatrix(1, 1) = ""
                gdLights.TextMatrix(1, 2) = ""
            Else
                gdLights.RemoveItem gdLights.Row
            End If
            If gdLights.Rows = 2 Then gdLights.Enabled = False
    End Select
End Sub

Private Sub gdLights_DblClick()
    txtEnter.Top = gdLights.RowPos(gdLights.Row) + gdLights.Top
    txtEnter.Left = gdLights.ColPos(ColOver) + gdLights.Left
    txtEnter.Height = gdLights.RowHeight(gdLights.Row)
    txtEnter.Width = gdLights.ColWidth(ColOver) + 17
    txtEnter.Visible = True
    gdX = ColOver: gdY = gdLights.Row
    txtEnter.Text = gdLights.TextMatrix(gdY, gdX)
    txtEnter.SetFocus
End Sub

Private Sub gdLights_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X > gdLights.ColPos(2) Then ColOver = 2 Else ColOver = 1
End Sub

Private Sub txtEnter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        gdLights.TextMatrix(gdX, gdY) = txtEnter
        txtEnter.Visible = False
    End If
End Sub

Private Sub txtEnter_KeyPress(KeyAscii As Integer)
    Dim Chara As String
    If gdX = 2 Then
        Chara = Chr(KeyAscii)
        Chara = UCase(Chara)
        If Chara < "A" Or Chara > "Z" Then KeyAscii = 0: Exit Sub
        KeyAscii = Asc(Chara)
    End If
End Sub

Private Sub txtEnter_LostFocus()
    gdLights.TextMatrix(gdY, gdX) = txtEnter
    txtEnter.Visible = False
End Sub

Private Sub ViewTab_Click()
    Dim n As Integer
    For n = 1 To 2: Frame(n).Visible = False: Next n
    Frame(ViewTab.SelectedItem.Index).Visible = True
End Sub

