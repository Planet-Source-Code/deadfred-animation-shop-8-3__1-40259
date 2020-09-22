VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEdit 
   Caption         =   "Form1"
   ClientHeight    =   4650
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5190
   Icon            =   "frmEdit.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4650
   ScaleWidth      =   5190
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin Project1.Engine Engine 
      Height          =   975
      Left            =   2160
      TabIndex        =   6
      Top             =   2880
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1720
   End
   Begin Project1.PaintWindow TexMap 
      Height          =   855
      Left            =   2160
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1508
   End
   Begin Project1.DXEngine DXEngine 
      Height          =   615
      Left            =   1200
      TabIndex        =   3
      Top             =   1320
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1085
   End
   Begin Project1.Tablet Tablet 
      Height          =   2175
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1815
      _ExtentX        =   5953
      _ExtentY        =   3413
   End
   Begin MSComctlLib.TabStrip MainTab 
      Height          =   3015
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5318
      MultiRow        =   -1  'True
      HotTracking     =   -1  'True
      ImageList       =   "GrayIcons"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   6
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Top View"
            Object.Tag             =   "1"
            Object.ToolTipText     =   "View your model from the top"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Front View"
            Object.Tag             =   "1"
            Object.ToolTipText     =   "View your model from the front"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Side View"
            Object.Tag             =   "1"
            Object.ToolTipText     =   "View your model from the side"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Texture Map"
            Object.Tag             =   "10"
            Object.ToolTipText     =   "View your models texture map"
            ImageVarType    =   2
            ImageIndex      =   4
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "3D View"
            Object.Tag             =   "7"
            Object.ToolTipText     =   "View your model in 3D"
            ImageVarType    =   2
            ImageIndex      =   3
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Preview"
            Object.Tag             =   "9"
            Object.ToolTipText     =   "Preview your model using DirectX"
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.LayerDisplay LayerTab 
      Height          =   300
      Left            =   240
      TabIndex        =   2
      Top             =   3120
      Visible         =   0   'False
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   529
   End
   Begin MSComctlLib.ImageList GrayIcons 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483644
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEdit.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEdit.frx":0894
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEdit.frx":0CE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEdit.frx":1138
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox TakeFocus 
      Height          =   285
      Left            =   360
      MaxLength       =   1
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   120
      Width           =   255
   End
   Begin VB.Menu menuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "&New..."
         Index           =   1
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Open..."
         Index           =   2
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Close"
         Index           =   3
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Save"
         Index           =   5
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Save &as..."
         Index           =   6
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Properties..."
         Index           =   7
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Inport..."
         Index           =   9
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Export..."
         Index           =   10
      End
      Begin VB.Menu mnuOldFile 
         Caption         =   "-"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExit 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Index           =   2
      End
   End
   Begin VB.Menu menuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEdit 
         Caption         =   "&Undo"
         Index           =   1
         Shortcut        =   %{BKSP}
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Cut"
         Index           =   3
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "C&opy"
         Index           =   4
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Paste"
         Index           =   5
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "D&uplicate"
         Index           =   6
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Delete"
         Index           =   7
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Group"
         Index           =   9
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Ungroup"
         Index           =   10
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Selec&t all"
         Index           =   11
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "D&eselect all"
         Index           =   12
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Hide selected"
         Index           =   14
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Lock selected"
         Index           =   15
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Grey Selected"
         Index           =   16
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Reveal all"
         Index           =   17
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Unl&ock all"
         Index           =   18
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Colour all"
         Index           =   19
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   20
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Settings..."
         Index           =   21
         Shortcut        =   +{F1}
      End
   End
   Begin VB.Menu PaintEdit 
      Caption         =   "&Edit"
      Visible         =   0   'False
      Begin VB.Menu mnuPEdit 
         Caption         =   "&Tile Image"
         Index           =   1
      End
      Begin VB.Menu mnuPEdit 
         Caption         =   "&Select Image"
         Index           =   2
      End
      Begin VB.Menu mnuPEdit 
         Caption         =   "&Clear Texture Map"
         Index           =   3
      End
      Begin VB.Menu mnuPEdit 
         Caption         =   "Texture &Properties"
         Index           =   4
      End
      Begin VB.Menu mnuPEdit 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuPEdit 
         Caption         =   "Settings"
         Index           =   6
      End
   End
   Begin VB.Menu menuView 
      Caption         =   "&View"
      Begin VB.Menu mnuView 
         Caption         =   "&Snap to grid"
         Checked         =   -1  'True
         Index           =   1
      End
      Begin VB.Menu mnuView 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuView 
         Caption         =   "Colour &Objects"
         Index           =   3
         Begin VB.Menu mnuColourObjects 
            Caption         =   "&Normal"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu mnuColourObjects 
            Caption         =   "&Grey"
            Index           =   2
         End
         Begin VB.Menu mnuColourObjects 
            Caption         =   "By &Joint"
            Index           =   3
         End
         Begin VB.Menu mnuColourObjects 
            Caption         =   "By &Layer"
            Index           =   4
         End
         Begin VB.Menu mnuColourObjects 
            Caption         =   "&Black"
            Index           =   5
         End
      End
      Begin VB.Menu mnuView 
         Caption         =   "Colour &Skeliton"
         Index           =   4
         Begin VB.Menu mnuColourSkeliton 
            Caption         =   "&Normal"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu mnuColourSkeliton 
            Caption         =   "&Grey"
            Index           =   2
         End
         Begin VB.Menu mnuColourSkeliton 
            Caption         =   "By &Layer"
            Index           =   3
         End
         Begin VB.Menu mnuColourSkeliton 
            Caption         =   "&Black"
            Index           =   4
         End
      End
      Begin VB.Menu mnuView 
         Caption         =   "&Tool bars"
         Index           =   5
         Begin VB.Menu mnuToolbars 
            Caption         =   "&File toolbar"
            Index           =   1
         End
         Begin VB.Menu mnuToolbars 
            Caption         =   "&Edit toolbar"
            Index           =   2
         End
         Begin VB.Menu mnuToolbars 
            Caption         =   "&Play toolbar"
            Index           =   3
         End
         Begin VB.Menu mnuToolbars 
            Caption         =   "&Select Toolbar"
            Index           =   4
         End
         Begin VB.Menu mnuToolbars 
            Caption         =   "-"
            Index           =   5
         End
         Begin VB.Menu mnuToolbars 
            Caption         =   "&3D toolbars"
            Checked         =   -1  'True
            Index           =   6
         End
         Begin VB.Menu mnuToolbars 
            Caption         =   "&Flat toolbars"
            Index           =   7
         End
      End
      Begin VB.Menu mnuView 
         Caption         =   "&Zoom"
         Index           =   6
         Begin VB.Menu mnuZoom 
            Caption         =   "25%"
            Index           =   1
         End
         Begin VB.Menu mnuZoom 
            Caption         =   "50%"
            Index           =   2
         End
         Begin VB.Menu mnuZoom 
            Caption         =   "100%"
            Checked         =   -1  'True
            Index           =   3
            Shortcut        =   {F9}
         End
         Begin VB.Menu mnuZoom 
            Caption         =   "150%"
            Index           =   4
         End
         Begin VB.Menu mnuZoom 
            Caption         =   "200%"
            Index           =   5
         End
         Begin VB.Menu mnuZoom 
            Caption         =   "400%"
            Index           =   6
         End
         Begin VB.Menu mnuZoom 
            Caption         =   "800%"
            Index           =   7
         End
         Begin VB.Menu mnuZoom 
            Caption         =   "-"
            Index           =   8
         End
         Begin VB.Menu mnuZoom 
            Caption         =   "-25%"
            Index           =   9
            Shortcut        =   {F7}
         End
         Begin VB.Menu mnuZoom 
            Caption         =   "+25%"
            Index           =   10
            Shortcut        =   {F8}
         End
         Begin VB.Menu mnuZoom 
            Caption         =   "-"
            Index           =   11
         End
         Begin VB.Menu mnuZoom 
            Caption         =   "&Previous zoom"
            Index           =   12
            Shortcut        =   {F6}
         End
         Begin VB.Menu mnuZoom 
            Caption         =   "&Selection"
            Index           =   13
         End
      End
      Begin VB.Menu mnuView 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuView 
         Caption         =   "S&idebar"
         Index           =   8
      End
      Begin VB.Menu mnuView 
         Caption         =   "&Layers"
         Index           =   9
      End
      Begin VB.Menu mnuView 
         Caption         =   "S&tatus Bar"
         Index           =   10
      End
      Begin VB.Menu mnuView 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu mnuView 
         Caption         =   "&Center View"
         Index           =   12
      End
   End
   Begin VB.Menu menuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuSelection 
         Caption         =   "&Align"
         Index           =   1
         Begin VB.Menu mnuAlign 
            Caption         =   "Align &Top"
            Index           =   1
         End
         Begin VB.Menu mnuAlign 
            Caption         =   "Align &Center"
            Index           =   2
         End
         Begin VB.Menu mnuAlign 
            Caption         =   "Align &Bottom"
            Index           =   3
         End
         Begin VB.Menu mnuAlign 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu mnuAlign 
            Caption         =   "Align &Left"
            Index           =   5
         End
         Begin VB.Menu mnuAlign 
            Caption         =   "Align &Middle"
            Index           =   6
         End
         Begin VB.Menu mnuAlign 
            Caption         =   "Align &Right"
            Index           =   7
         End
      End
      Begin VB.Menu mnuTools 
         Caption         =   "&Gallaries"
         Index           =   2
         Begin VB.Menu mnuGallary 
            Caption         =   "&Create new Gallary"
            Index           =   1
         End
         Begin VB.Menu mnuGallary 
            Caption         =   "Re&name Gallary"
            Index           =   2
         End
         Begin VB.Menu mnuGallary 
            Caption         =   "&Remove Gallary"
            Index           =   3
         End
         Begin VB.Menu mnuGallary 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu mnuGallary 
            Caption         =   "&Add selected to Gallary"
            Index           =   5
         End
      End
      Begin VB.Menu mnuTools 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuTools 
         Caption         =   "&Object Editor..."
         Index           =   4
      End
      Begin VB.Menu mnuTools 
         Caption         =   "&Surface Editor..."
         Index           =   5
      End
      Begin VB.Menu mnuTools 
         Caption         =   "&Joint Editor..."
         Index           =   6
      End
      Begin VB.Menu mnuTools 
         Caption         =   "&Entity Editor..."
         Index           =   7
      End
      Begin VB.Menu mnuTools 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuTools 
         Caption         =   "Object from &Bitmap..."
         Index           =   9
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTools 
         Caption         =   "Object &Makeup..."
         Index           =   10
      End
      Begin VB.Menu mnuTools 
         Caption         =   "Record Animation..."
         Index           =   11
      End
      Begin VB.Menu mnuTools 
         Caption         =   "Animation Viewer..."
         Index           =   12
      End
      Begin VB.Menu mnuTools 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu mnuTools 
         Caption         =   "Compile &Folder..."
         Index           =   14
      End
   End
   Begin VB.Menu menuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindow 
         Caption         =   "Arrange &Vertically"
         Index           =   1
      End
      Begin VB.Menu mnuWindow 
         Caption         =   "Arrange &Horizontally"
         Index           =   2
      End
      Begin VB.Menu mnuWindow 
         Caption         =   "&Cascade"
         Index           =   3
      End
      Begin VB.Menu mnuWindow 
         Caption         =   "Arrange &Icons"
         Index           =   4
      End
      Begin VB.Menu mnuWindow 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuWindow 
         Caption         =   "&New Window"
         Index           =   6
      End
   End
   Begin VB.Menu menuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "&About"
         Index           =   1
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&Help"
         Index           =   2
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&Tip of the day"
         Index           =   4
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&Whats this sidebar?"
         Index           =   5
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&On the web"
         Index           =   6
         Begin VB.Menu mnuOnTheWeb 
            Caption         =   "&Animation Shop 8.3"
            Index           =   1
         End
         Begin VB.Menu mnuOnTheWeb 
            Caption         =   "Get &Registration Code"
            Index           =   2
            Visible         =   0   'False
         End
         Begin VB.Menu mnuOnTheWeb 
            Caption         =   "&Download Samples"
            Index           =   3
         End
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "How do &I..."
         Index           =   7
         Begin VB.Menu meuQuickHelp 
            Caption         =   "meuQuickHelp"
            Index           =   0
         End
      End
   End
   Begin VB.Menu menuEditPopup 
      Caption         =   "menuEditPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuEditPopup 
         Caption         =   "Undo"
         Index           =   0
      End
      Begin VB.Menu mnuEditPopup 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuEditPopup 
         Caption         =   "Cut"
         Index           =   2
      End
      Begin VB.Menu mnuEditPopup 
         Caption         =   "Copy"
         Index           =   3
      End
      Begin VB.Menu mnuEditPopup 
         Caption         =   "Paste"
         Index           =   4
      End
      Begin VB.Menu mnuEditPopup 
         Caption         =   "Duplicate"
         Index           =   5
      End
      Begin VB.Menu mnuEditPopup 
         Caption         =   "Delete"
         Index           =   6
      End
      Begin VB.Menu mnuEditPopup 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuEditPopup 
         Caption         =   "Group"
         Index           =   8
      End
      Begin VB.Menu mnuEditPopup 
         Caption         =   "Ungroup"
         Index           =   9
      End
      Begin VB.Menu mnuEditPopup 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuEditPopup 
         Caption         =   "Joint object to"
         Index           =   11
      End
      Begin VB.Menu mnuEditPopup 
         Caption         =   "Add to layer"
         Index           =   12
         Begin VB.Menu mnuAddtoLayer 
            Caption         =   ":-)"
            Index           =   0
         End
      End
   End
   Begin VB.Menu menuLayer 
      Caption         =   "menuLayer"
      Visible         =   0   'False
      Begin VB.Menu mnuLayer 
         Caption         =   "New layer"
         Index           =   0
      End
      Begin VB.Menu mnuLayer 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuLayer 
         Caption         =   "Rename layer"
         Index           =   2
      End
      Begin VB.Menu mnuLayer 
         Caption         =   "Remove layer"
         Index           =   3
      End
      Begin VB.Menu mnuLayer 
         Caption         =   "Select layer"
         Index           =   4
      End
      Begin VB.Menu mnuLayer 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuLayer 
         Caption         =   "Locked"
         Index           =   6
      End
   End
   Begin VB.Menu menuRightDrag 
      Caption         =   "menuRightDrag"
      Visible         =   0   'False
      Begin VB.Menu mnuRightDrag 
         Caption         =   "Copy Here"
         Index           =   1
      End
      Begin VB.Menu mnuRightDrag 
         Caption         =   "Move Here"
         Index           =   2
      End
      Begin VB.Menu mnuRightDrag 
         Caption         =   "Attach to..."
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ############################################################################
' #                                                                          #
' #  The edit window is the main part of the program, as it holds all the    #
' #  controls used to display a model. It allows you to set the view mode,   #
' #   and allows you to edit the model using the mouse. It contains most     #
' #   of the menus found in the program, and the code that opperates them    #
' #                                                                          #
' ############################################################################

Public WindowKey As String
Public File As clsFile


Private Sub Engine_MouseDown(X As Single, Y As Single, Button As Integer, Shift As Integer)
    Engine.pRenderSolid = False
End Sub



Private Sub Form_Load()
    'This is run whenever a new edit form is loaded. This code just loads the quick help feature into the menus
    Am8.LoadQuickHelp Me
End Sub

Private Sub Form_Activate()
    'This is run every time an edit form is activeated, which means when you swap from one form to another.
    'It updates information shown in the sidebars about each file, and updates the menu history
    Dim n As Integer
    For n = 1 To 5
        If Val(RightClip(frmMain.cmdZoomLevels.List(n), 1)) = Tablet.ZoomLevel * 100 Then frmMain.cmdZoomLevels.ListIndex = n
    Next n
    ActiveFile = File.Key
    File.Joint.DisplayTreeInWindow frmMain.Joints
    Am8(ActiveFile).Scene.AddSceneToWindow frmMain.trFrames
    MainTab_Mouseup 0, 0, 0, 0
    UpdateEditHistory
    
    Engine.PicketAnimateOver = 1
    Engine.PicketAnimate = 3
    
End Sub


Private Sub mnuZoom_Click(Index As Integer)
    Dim n As Integer
    For n = 1 To 6: mnuZoom(n).Checked = False: Next n
    Select Case Index
        Case 1 To 6
            Tablet.ZoomLevel = Val(Mid(mnuZoom(Index).Caption, 1, Len(mnuZoom(Index).Caption) - 1)) / 100
            mnuZoom(Index).Checked = True
        Case 9:  Tablet.ZoomLevel = Tablet.ZoomLevel - 0.25: If Tablet.ZoomLevel < 0.25 Then Tablet.ZoomLevel = 0.25
        Case 10: Tablet.ZoomLevel = Tablet.ZoomLevel + 0.25: If Tablet.ZoomLevel > 8 Then Tablet.ZoomLevel = 8
        Case 12: Tablet.ZoomLevel = Tablet.LastZoomLevel
        Case 13: Tablet.ZoomLevel = Tablet.ZoomToSelected
    End Select
    frmMain.cmdZoomLevels.List(5) = (Tablet.ZoomLevel * 100) & "%"
    frmMain.cmdZoomLevels.ListIndex = 5
    Tablet.Refresh
End Sub

Public Sub AssignWindowToFile(AssignFile As clsFile)
    'Each edit window must point at a File class in order to work. This command sets that file class
    Set File = AssignFile
    Caption = File.ModelName
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then If frmMain.RemoveWindow(Caption, , WindowKey) = False Then Cancel = 1
End Sub

Public Sub CauseFormResize()
    'This allows other forms and modules to resize the form
    Form_Resize
End Sub

Private Sub Form_Resize()
    'This code does the actual resizing of the objects to fit the shape of the form
    On Error Resume Next
    If LayerTab.Visible = True Then MainTab.Height = ScaleHeight - (MainTab.Top * 2) - 200 Else MainTab.Height = ScaleHeight - (MainTab.Top * 2)
    MainTab.Width = ScaleWidth - (MainTab.Left * 2)
    Tablet.Move MainTab.ClientLeft, MainTab.ClientTop, MainTab.ClientWidth, MainTab.ClientHeight
    TexMap.Move MainTab.ClientLeft, MainTab.ClientTop, MainTab.ClientWidth, MainTab.ClientHeight
    DXEngine.Move MainTab.ClientLeft, MainTab.ClientTop, MainTab.ClientWidth, MainTab.ClientHeight
    Engine.Move MainTab.ClientLeft, MainTab.ClientTop, MainTab.ClientWidth, MainTab.ClientHeight
    LayerTab.Move 60, ScaleHeight - LayerTab.Height, ScaleWidth - 120
    Tablet.Refresh
End Sub

Private Sub mnuAddtoLayer_Click(Index As Integer)
    'This moves all the selected objects into the selected layer
    Dim Am As clsObject
    frmMain.sBar.Panels(2) = "Objects moved to " & mnuAddtoLayer(Index).Caption
    For Each Am In File.Geometery
        If Am.Selected = True Then Am.Layer = File.Layers(Index).LayerKey
    Next Am
End Sub

Private Sub UpdateEditHistory()
    Dim n As Integer, Length As Integer
    For n = mnuOldFile.Count - 1 To 1 Step -1: Unload mnuOldFile(n): Next n
    Length = Am8.FileHistory.Lenght
    If Length > Am8.FileHistory.CountHistory Then Length = Am8.FileHistory.CountHistory
    For n = 1 To Length
        Load mnuOldFile(n)
        mnuOldFile(n).Visible = True
        If Am8.FullPath = True Then
            mnuOldFile(n).Caption = "&" & n & ". " & MaxLength(Am8.FileHistory(n).FilePath, 20, 3)
        Else
            mnuOldFile(n).Caption = "&" & n & ". " & Am8.FileHistory(n).FileName
        End If
    Next n
    If Length = 0 Then mnuOldFile(0).Visible = False Else mnuOldFile(0).Visible = True
    frmMain.UpdateHistoryMenu
End Sub

Private Sub MainTab_Mouseup(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim n As Integer
'    For n = 1 To MainTab.Tabs.Count
'        If X >= MainTab.Tabs(n).Left + MainTab.Left And X <= MainTab.Tabs(n).Left + MainTab.Left + MainTab.Tabs(n).Width Then MainTab.Tabs(n).Selected = True
'        MainTab.hi
'    Next n
    File.Geometery.RemoveObject "GroundPlain"
    Select Case MainTab.SelectedItem.Index
        Case 1 To 3
            DXEngine.Visible = False: Engine.Visible = False
            TexMap.Visible = False: Tablet.Visible = True
            For n = 1 To 6: frmMain.tbar(1).buttons(n).Visible = True: Next n
            For n = 7 To 10: frmMain.tbar(1).buttons(n).Visible = False: Next n
            Tablet.ViewMode = MainTab.SelectedItem.Index
            frmMain.sBar.Panels(2) = amSet2DViewMode
            If Am8.AlwaysCenter = True Then Tablet.CenterView
            Tablet.Refresh: MenuEdit.Visible = True: PaintEdit.Visible = False
        
        Case 4
            DXEngine.Visible = False: Engine.Visible = False
            TexMap.Visible = True: Tablet.Visible = False
            For n = 1 To 9: frmMain.tbar(1).buttons(n).Visible = False: Next n
            frmMain.tbar(1).buttons(10).Visible = True
            frmMain.sBar.Panels(2) = amSetTexMap
            MenuEdit.Visible = False: PaintEdit.Visible = True
            TexMap.Refresh
        
        Case 5
            DXEngine.Visible = False: Engine.Visible = True
            TexMap.Visible = False: Tablet.Visible = False
            If frmMain.ck3Dopt(2) = 1 Then
                File.Geometery.CreateObject "GroundPlain"
                File.Geometery("GroundPlain").CreateObject "Grid", 1, -400, 400, 50, 50, -400, 400, 3, 3
                File.Geometery("GroundPlain").Layer = "Main"
            End If
            For n = 1 To 10: frmMain.tbar(1).buttons(n).Visible = False: Next n
            frmMain.tbar(1).buttons(7).Visible = True
            frmMain.tbar(1).buttons(8).Visible = True
            Engine.RefreshView
            frmMain.sBar.Panels(2) = "View your model in 3D"
            MenuEdit.Visible = True: PaintEdit.Visible = False
        
        Case 6
            MenuEdit.Visible = True: PaintEdit.Visible = False
            DXEngine.Visible = True: Engine.Visible = False
            TexMap.Visible = False: Tablet.Visible = False
            For n = 1 To 10: frmMain.tbar(1).buttons(n).Visible = False: Next n
            frmMain.tbar(1).buttons(9).Visible = True
            DXEngine.AssignDXEngineTo File
            DXEngine.RefreshModel: frmMain.sBar.Panels(2) = amSetDirectXViewMode

    End Select
    frmMain.tbar(1).buttons(Val(MainTab.SelectedItem.Tag)).Value = tbrPressed
    frmMain.ShowSidebar frmMain.tbar(1).buttons(EditButton).Key
    Tablet.Refresh
End Sub

Private Sub mnuAlign_Click(Index As Integer)
    'The align menu calls the align function that moves the selected object into straight lines
    Am8(Tablet.FileKey).Geometery.Align Index, Tablet.ViewMode
    Am8(Tablet.FileKey).FindModelOutline: Tablet.Refresh
End Sub

Private Sub mnuGallary_Click(Index As Integer)
    Dim NewName As String, Resp As String, FileName As String, n As Integer
    On Error GoTo FailedGalleryMove
    Select Case Index
        Case 1
            NewName = InputBox(amNewGallaryName)
            If NewName = "" Then Exit Sub
            MkDir App.Path & "\data\Gallarys\" & Trim(NewName)
            GetGallaryFolders
            For n = 0 To frmMain.cmbGallary.ListCount - 1
                If frmMain.cmbGallary.List(n) = NewName Then frmMain.cmbGallary.ListIndex = n
            Next n
        
        Case 2
            If frmMain.cmbGallary.Text = "[None]" Then MsgBox amNoGallary, vbInformation: Exit Sub
            NewName = InputBox(frmMain.cmbGallary & vbNewLine & vbNewLine & amRenameGallary, , frmMain.cmbGallary)
            If NewName = "" Then Exit Sub
            ChDir App.Path & "\data\Gallarys\" & Trim(NewName)
            GetGallaryFolders
            For n = 0 To frmMain.cmbGallary.ListCount - 1
                If frmMain.cmbGallary.List(n) = NewName Then frmMain.cmbGallary.ListIndex = n
            Next n
        
        Case 3
            If frmMain.cmbGallary.Text = "[None]" Then MsgBox amNoGallary, vbInformation: Exit Sub
            If MsgBox(frmMain.cmbGallary & vbNewLine & vbNewLine & amConfirmRemoveGallary, vbQuestion + vbYesNo + vbDefaultButton2) = 7 Then Exit Sub
            RmDir App.Path & "\data\Gallarys\" & frmMain.cmbGallary
            GetGallaryFolders
            frmMain.cmbGallary.ListIndex = 0
        
        Case 5
            If frmMain.cmbGallary.Text = "[None]" Then MsgBox amSelectGallaryFirst, vbInformation: Exit Sub
            Resp = InputBox(frmMain.cmbGallary.Text & vbNewLine & vbNewLine & amAddToGallary)
            If Trim(Resp) = "" Then Exit Sub
            If modFunctions.CheckOverwrite(App.Path & "\data\Gallarys\" & frmMain.cmbGallary & "\" & Resp & ".cpy", 0) = False Then Exit Sub
            With Am8(ActiveFile)
                .SaveToFile App.Path & "\data\Gallarys\" & frmMain.cmbGallary & "\" & Resp & ".cpy", 1, (.MinX + .MaxX) / 2, (.MinY + .MaxY) / 2, (.MinZ + .MaxZ) / 2
                frmMain.Gallary.FolderLocation = App.Path & "\data\gallarys\" & frmMain.cmbGallary.List(frmMain.cmbGallary.ListIndex)
                frmMain.sBar.Panels(2) = "New gallary item created '" & Resp & "'"
            End With
    End Select
Exit Sub

FailedGalleryMove:
    Select Case Index
        Case 1
            MsgBox amFailedToRemoveGallary
        Case 3
            If MsgBox(frmMain.cmbGallary & vbNewLine & vbNewLine & amGallaryNotEmpty, vbInformation + vbYesNo + vbDefaultButton2) = 7 Then Exit Sub
            Kill App.Path & "\data\Gallarys\" & frmMain.cmbGallary & "\*.*"
            RmDir App.Path & "\data\Gallarys\" & frmMain.cmbGallary
    End Select
    GetGallaryFolders
End Sub

Private Sub mnuEdit_Click(Index As Integer)
    'This menu is the edit menu at the top of the screen, not the popup menu, but they both do just
    'about the same thing. This menu also has the Settings option
    With Am8(Tablet.FileKey)
        Select Case Index
            Case 3: .SaveToFile App.Path & "\data\clipboard", 1: .DeleteSelected 'CUT to clipboard
            Case 4: .SaveToFile App.Path & "\data\clipboard", 1:  'COPY to clipboard
            Case 5: .LoadFromFile App.Path & "\data\clipboard", 1 'Paste from clipboard
            Case 6: .DuplicateSelection Tablet.ViewMode 'Duplicate selection
            Case 7: .DeleteSelected 'Delete Selection
            Case 9: .Geometery.GroupSelected 'Group the selected objects together
            Case 10: .Geometery.UngroupSelected 'Ungroup the selected object group
            Case 11: .SelectAll: .FindModelOutline 'Select all
            Case 12: .DeselectAll: .FindModelOutline 'Deselect all
            Case 14: .HideSelected: .DeselectAll 'Hide selected
            Case 15: .LockSelected: .DeselectAll 'Lock selected
            Case 16: .GreySelected 'Grey selected
            Case 17: .UnHideAll
            Case 18: .UnLockAll
            Case 19: .UnGreyAll
            Case 21: Am8.ShowSettings: UpdateEditHistory 'Show the settings window, and update file hisotry menus
        End Select
    End With
    Tablet.Refresh
End Sub

Private Sub mnuEditPopup_Click(Index As Integer)
    Dim JOver As Integer, Am As clsObject
    With Am8(Tablet.FileKey)
        Select Case Index
            Case 2:  .SaveToFile App.Path & "\data\clipboard", 1: .DeleteSelected 'CUT to clipboard
            Case 3:  .SaveToFile App.Path & "\data\clipboard", 1 ' COPY to clipboard
            Case 4:  .LoadFromFile App.Path & "\data\clipboard", 1 'Paste from clipboard
            Case 5:  .DuplicateSelection Tablet.ViewMode 'Duplicate selection
            Case 6:  .DeleteSelected 'Delete selection
            Case 8:  .Geometery.GroupSelected 'Group selected objects
            Case 9:  .Geometery.UngroupSelected 'Ungroup selected objects
            Case 11: For Each Am In .Geometery: If Am.Selected = True Then Am.AttachObjectTo Tablet.JointOver
                     Next Am
        End Select
    End With
    Tablet.Refresh
End Sub

Private Sub mnuFile_Click(Index As Integer)
    Dim FileName As String
    Select Case Index
        Case 1: Am8.ShowNew
        Case 2
            FileName = SelectFileName("Am8", amOpenFileName)
            If FileName <> "" Then frmMain.LoadExistingFileWithWindow FileName
        Case 3: frmMain.RemoveWindow Caption, File.Key
        Case 5:
            If File.CurrentFilePath = "" Then
                FileName = SetFileName("Am8", amSaveModelTo)
                If FileName <> "" Then
                    File.SaveToFile FileName
                    File.CurrentFilePath = FileName
                    File.ModelName = RightClip(Mid(FileName, InStrRev(FileName, "\") + 1), 4)
                End If
            Else
                File.SaveToFile File.CurrentFilePath
            End If
            Caption = File.ModelName
                
        Case 6
            FileName = SetFileName("Am8", amSaveModelTo)
            If FileName <> "" Then
                File.SaveToFile FileName
                File.CurrentFilePath = FileName
                File.ModelName = RightClip(Mid(FileName, InStrRev(FileName, "\") + 1), 4)
            End If
            Caption = File.ModelName
        Case 7
            frmProperties.RunAtStart File
            
            
        Case 9
            If MainTab.SelectedItem.Index = 4 Then
                FileName = SelectFileName("Picture", "Import texture...")
                If FileName = "" Then Exit Sub
                TexMap.LoadImage FileName
            Else
                FileName = SelectFileName("Import", amInportFileName)
                If FileName = "" Then Exit Sub
                modImport.ImportModel FileName, Am8(Tablet.FileKey)
                Tablet.Refresh
            End If
        Case 10
            If MainTab.SelectedItem.Index = 4 Then
                FileName = SetFileName("Picture", "Save texture...")
                If FileName = "" Then Exit Sub
                TexMap.SaveImage FileName
            Else
                frmCompile.RunAtStart File
            End If
            
    End Select
End Sub

Private Sub mnuHelp_Click(Index As Integer)
    'This is the Help menu object. It does all the standard help stuff, such as About box, Help file, whats
    'this toolbar, tip-of-the-day and so on.
    Dim n As Integer, FrameON As Integer
    Select Case Index
        Case 1: Am8.ShowAbout
        Case 2: Am8.ShowHelp
        Case 4: Am8.ShowTipofDay
        Case 5
            For n = 1 To frmMain.Frame.Count
                If frmMain.Frame(n).Visible = True Then FrameON = n
            Next n
            Select Case FrameON
                Case 1: Am8.ShowHelp "Creating New Objects"
                Case 2: Am8.ShowHelp "Select Sidebar"
                Case 3: Am8.ShowHelp "Gallary Sidebar"
                Case 4: Am8.ShowHelp "Skeleton Sidebar"
                Case 5: Am8.ShowHelp "Rotate Sidebar"
                Case 6: Am8.ShowHelp "Scale Sidebar"
                Case 7: Am8.ShowHelp "AI Links Sidebar"
                Case 8: Am8.ShowHelp "Wire-frame Sidebar"
                Case 9: Am8.ShowHelp "Shading Sidebar"
                Case 10: Am8.ShowHelp "Preview Sidebar"
                Case 11: Am8.ShowHelp "Edit 1 Sidebar"
                Case 12: Am8.ShowHelp "Edit 2 Sidebar"
                Case 13: Am8.ShowHelp "Edit 3 Sidebar"
                Case 14: Am8.ShowHelp "Scenes Sidebar"
                Case 15: Am8.ShowHelp "Position Sidebar"
            End Select
    End Select
End Sub

Private Sub meuQuickHelp_Click(Index As Integer)
    'This is the Quick Help feature, showing questions and answers
    MsgBox meuQuickHelp(Index).Tag, vbInformation, "How do I " & meuQuickHelp(Index).Caption
End Sub

Private Sub mnuOldFile_Click(Index As Integer)
    'This allows you to load files from the file history shown in the file menu
    frmMain.CreateNewFileWithWindow Am8.FileHistory(Index).FileName
    Am8(ActiveFile).LoadFromFile Am8.FileHistory(Index).FilePath
    frmMain.ActiveForm.Tablet.Refresh
End Sub

Private Sub mnuOnTheWeb_Click(Index As Integer)
    'This menu item enables you to connect to websites on the internet so that you can
    'get updates, help and sample files
    Dim Ie As InternetExplorer
    Set Ie = New InternetExplorer
    Select Case Index
        Case 1: Ie.Navigate "www.geocities.com/animationshop8/index.htm"
        Case 2: Ie.Navigate "www.geocities.com/animationshop8/register.htm"
        Case 3: Ie.Height = 0: Ie.Width = 0: Ie.Navigate "www.geocities.com/animationshop8/samples.zip"
    End Select
    Ie.Visible = True
End Sub

Private Sub mnuTools_Click(Index As Integer)
    Dim SceneName As String
    Select Case Index
        Case 4: frmObject.RunAtStart File
        Case 5: frmSurface.RunAtStart File
        Case 6: frmJoint.RunAtStart File
        Case 10: frmOutline.RunAtStart File
        Case 11
            If mnuTools(11).Checked = True Then
                mnuTools(11).Checked = False
                Else
                SceneName = SetFileName("Picture", "Select Frame name")
                If SceneName <> "" Then
                    mnuTools(11).Checked = True
                    Engine.AnimationFrame = 0
                    Engine.AnimationSceneName = RightClip(SceneName, 4)
                End If
            End If
        Case 12: Am8.OpenAnimator
        Case 14: frmCompile.RunAtStart
    End Select
End Sub

Private Sub mnuView_Click(Index As Integer)
    'This controls the view options, such as snap to grid, show sidebar, status bar, toolbars. When you
    'tick a menu item in one edit menu, the code goes through each of the other edit windows, and ticks
    'their menus aswell, so they are all the same
    Dim n As Integer, NewValue As Boolean
    NewValue = InvertBo(mnuView(Index).Checked)
    
    Select Case Index
        Case 1 'Snap to grid
            For n = 1 To Am8.Forms.Count: Am8.Forms(n).mnuView(Index).Checked = NewValue: Next n
            
        Case 8 'Show or hide the sidebar
            For n = 1 To Am8.Forms.Count
                Am8.Forms(n).mnuView(Index).Checked = NewValue
                frmMain.SideFrame.Visible = NewValue
                If NewValue = True Then frmMain.tbar(0).buttons(5).Value = tbrPressed Else frmMain.tbar(0).buttons(5).Value = tbrUnpressed
            Next n: Am8.ShowSidebar = NewValue

        Case 9 'Show or hide the layers tab in each edit window
            For n = 1 To Am8.Forms.Count
                Am8.Forms(n).mnuView(Index).Checked = NewValue
                Am8.Forms(n).LayerTab.Visible = NewValue
                Am8.Forms(n).CauseFormResize
            Next n: Am8.ShowLayers = NewValue

        Case 10 'Show or hide the status bar
            For n = 1 To Am8.Forms.Count
                Am8.Forms(n).mnuView(Index).Checked = NewValue
                frmMain.sBar.Visible = NewValue
            Next n: Am8.ShowStatusBar = NewValue
            
        Case 12: Tablet.CenterView 'Centre the view mode on the selected objects
    End Select
End Sub

Private Sub mnuWindow_Click(Index As Integer)
    'This arranges the edit windows in the MDI form, and also allows you to create more than one window per file
    Select Case Index
        Case 1: frmMain.Arrange 2
        Case 2: frmMain.Arrange 1
        Case 3: frmMain.Arrange 0
        Case 4: frmMain.Arrange 0
        Case 6: frmMain.AddWindowToCurrentFile File.Key
    End Select
End Sub

Private Sub mnuToolbars_Click(Index As Integer)
    Dim n As Integer, NewValue As Boolean
    NewValue = InvertBo(mnuToolbars(Index).Checked)
    Select Case Index
        Case 1, 2, 3, 4 'This controls each of the four toolbars
            For n = 1 To Am8.Forms.Count: Am8.Forms(n).mnuToolbars(Index).Checked = NewValue: Next n
            frmMain.cBar.Bands(Index).Visible = NewValue
            If mnuToolbars(1).Checked = False And mnuToolbars(2).Checked = False And mnuToolbars(3).Checked = False And mnuToolbars(4).Checked = False Then frmMain.cBar.Visible = False Else frmMain.cBar.Visible = True
            frmMain.CauseFormResize
        
        Case 6, 7 'Controls the Flat / 3D appearence tick options
            For n = 1 To Am8.Forms.Count
                Am8.Forms(n).mnuToolbars(6).Checked = InvertBo(Am8.Forms(n).mnuToolbars(6).Checked)
                Am8.Forms(n).mnuToolbars(7).Checked = InvertBo(Am8.Forms(n).mnuToolbars(7).Checked)
            Next n
            For n = 0 To frmMain.tbar.Count - 1
                If Index = 6 Then frmMain.tbar(n).Style = ccFlat Else frmMain.tbar(n).Style = cc3D
                frmMain.tbar(n).Refresh
            Next n
            frmMain.cBar.Refresh
    End Select
End Sub

Private Sub mnuPEdit_Click(Index As Integer)
    Dim FileName As String
    Select Case Index
        Case 1: TexMap.TileImage
        Case 2: FileName = SelectFileName("Picture", "Choose Bitmap")
                If FileName <> "" Then TexMap.PictureFileName = FileName
        Case 3: If MsgBox(amClearTexture, vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then TexMap.ClearImage
        Case 4: ' frmTextWidth.RunAtStart TexMap
        Case 6: Am8.ShowSettings
    End Select
End Sub

Private Sub mnuExit_Click(Index As Integer)
    Unload frmMain
End Sub

Private Sub LayerTab_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single, TabOver As Integer)
    Dim Locked As Byte, Layer As clsLayer, Am As clsObject
    If TabOver <> 0 Then
        If Button = 1 Then
            If Am8(Tablet.FileKey).Layers(TabOver).LayerLocked = False Then
                For Each Am In Am8(ActiveFile).Geometery
                    If Am.Layer = Am8(Tablet.FileKey).Layers(TabOver).LayerKey Then Am.Selected = False
                Next Am
                Am8(Tablet.FileKey).FindModelOutline
            End If
        End If
        If Button = 2 Then
            mnuLayer(2).Caption = "Rename '" & Am8(Tablet.FileKey).Layers(TabOver).LayerName & "'"
            If Am8(Tablet.FileKey).Layers(TabOver).LayerLocked = False Then
                mnuLayer(6).Caption = "Lock '" & Am8(Tablet.FileKey).Layers(TabOver).LayerName & "'"
            Else
                mnuLayer(6).Caption = "UnLock '" & Am8(Tablet.FileKey).Layers(TabOver).LayerName & "'"
            End If
            PopupMenu menuLayer
        End If
    End If
    Tablet.Refresh
    For Each Layer In Am8(Tablet.FileKey).Layers
        If Layer.Selected = True Then frmMain.sBar.Panels(2) = "": Exit Sub
    Next Layer
    frmMain.sBar.Panels(2) = amAllLayersHidden
End Sub

Private Sub mnuLayer_Click(Index As Integer)
    Dim n As Integer, LayerName As String, X As Integer, RemovingKey As String
    Dim NewTag As String, hiden As Byte, Am As clsObject
    mnuEditPopup(12).Visible = False: mnuAddtoLayer(0).Visible = True

    Select Case Index
        Case 0
            LayerName = Trim(InputBox(amEnterNewLayerName, "Create layer"))
            If LayerName = "" Then Exit Sub
            File.Layers.AddLayer LayerName, "Key" & Timer
            For n = 1 To Am8.Forms.Count: Am8.Forms(n).LayerTab.Update: Next n
            
        Case 2
            LayerName = InputBox(amChangeLayerName, "Rename layer", Am8(Tablet.FileKey).Layers(LayerTab.LayerOver).LayerName)
            If LayerName = "" Then Exit Sub
            File.Layers(LayerTab.LayerOver).LayerName = LayerName
            LayerTab.Update

        Case 3
            If File.Layers.CountLayers = 1 Then MsgBox amCantDeleteAllLayers, vbInformation: Exit Sub
            Select Case MsgBox("The layer '" & Am8(Tablet.FileKey).Layers(LayerTab.LayerOver).LayerName & "' will be removed" & vbNewLine & "Do you also want to delete all the objects that are in this layer?", vbYesNoCancel + vbInformation + vbDefaultButton3)
                Case vbYes
                    For Each Am In Am8(ActiveFile).Geometery
                        If Am.Layer = File.Layers(LayerTab.LayerOver).LayerKey Then File.Geometery.RemoveObject Am.Key
                    Next Am
                    File.Layers.RemoveLayers LayerTab.LayerOver
                
                Case vbNo
                    For Each Am In Am8(ActiveFile).Geometery
                        If Am.Layer = File.Layers(LayerTab.LayerOver).LayerKey Then Am.Layer = File.Layers(2).LayerKey
                    Next Am
                    File.Layers.RemoveLayers LayerTab.LayerOver
                    
                Case vbCancel
            End Select
            LayerTab.Update
        
        Case 4
            For n = 1 To File.Geometery.CountObjects
                If File.Geometery(n).Layer = File.Layers(LayerTab.LayerOver).LayerKey Then File.Geometery(n).Selected = True
            Next n
            Am8(Tablet.FileKey).FindModelOutline
            Tablet.Refresh
    
        Case 6
            If Mid(mnuLayer(6).Caption, 1, 1) = "U" Then File.Layers(LayerTab.LayerOver).LayerLocked = False Else File.Layers(LayerTab.LayerOver).LayerLocked = True
            LayerTab.Update
    
    End Select
    UpdateLayerMenu Me
End Sub




Private Sub Tablet_UpdateOtherWindows()
    Dim fm As frmEdit
    For Each fm In Am8.Forms
        If fm.WindowKey <> WindowKey Then
            If Not fm.File Is Nothing Then
                If fm.MainTab.SelectedItem.Index = 5 Then
                    fm.Engine.RefreshView
                Else
                    If fm.File.Key = File.Key Then fm.Tablet.Refresh False
                End If
            End If
        End If
    Next fm
End Sub












