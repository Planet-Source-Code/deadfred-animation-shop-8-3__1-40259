VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Animation Shop 8"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6435
   Icon            =   "frmNew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2895
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   6015
      Begin MSComctlLib.ListView lstfiles 
         Height          =   2895
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   5106
         View            =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ColHdrIcons     =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "File Name"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.CommandButton cmbAct 
      Caption         =   "Okay"
      Default         =   -1  'True
      Height          =   350
      Index           =   2
      Left            =   4080
      TabIndex        =   3
      ToolTipText     =   "Click to close this window and save your changes"
      Top             =   4680
      Width           =   1095
   End
   Begin VB.FileListBox GetTemplates 
      Height          =   285
      Left            =   4200
      Pattern         =   "*.am8"
      TabIndex        =   8
      Top             =   4800
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2895
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   6015
      Begin MSComctlLib.ListView lsRecent 
         Height          =   2895
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   5106
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList2"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Filename"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.CommandButton cmbAct 
      Caption         =   "Help"
      Height          =   350
      Index           =   0
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "Click to get help on using this window"
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmbAct 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   350
      Index           =   1
      Left            =   5280
      TabIndex        =   4
      ToolTipText     =   "Click to close this window without saving your changes"
      Top             =   4680
      Width           =   1095
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   2160
      Top             =   4440
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
            Picture         =   "frmNew.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNew.frx":0896
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1560
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNew.frx":0CEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNew.frx":100E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNew.frx":1462
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNew.frx":18B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNew.frx":1D0A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip NewTab 
      Height          =   3375
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   5953
      HotTracking     =   -1  'True
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Welcome"
            Object.ToolTipText     =   "Select the option you want to preform"
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Recent"
            Object.ToolTipText     =   "Open a model that has recently been worked on"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VB.Image AppLogo 
      BorderStyle     =   1  'Fixed Single
      Height          =   1005
      Left            =   120
      Picture         =   "frmNew.frx":215E
      Stretch         =   -1  'True
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'#####################################################################
'#                                                                   #
'#   This form holds the code that creates new files. It can either  #
'#  display the form and allow you to choose your own type of file,  #
'#     or it can just create a simple blank file without a user      #
'#   prompt, depending on the value of 'MODE' in the routine below   #
'#                                                                   #
'#####################################################################

Public Sub RunAtStart()
    'This function is used to load the form from other forms/modules
    Show vbModal
End Sub

Private Sub Form_Load()
    'This adds all the differnt icons into the list box. Each icon repersents a different type of
    'file that you can create or load, such as templates, inport files etc. It also fills in the
    'file history box on the second page, with every file in the history, not just the first 5
    Dim n As Integer: lstfiles.ListItems.Clear
    lstfiles.ListItems.Add , , "Blank model", 1, 1
    lstfiles.ListItems.Add , , "Existing file", 2, 2
    lstfiles.ListItems.Add , , "Inport file", 3, 3
    lstfiles.ListItems.Add , , "Compile Folder", 5, 5
    For n = 1 To Am8.FileHistory.CountHistory
        lsRecent.ListItems.Add , "FH" & n, MaxLength(Am8.FileHistory(n).FilePath, 40, 6), 1, 1
        lsRecent.ListItems("FH" & n).ToolTipText = Am8.FileHistory(n).FilePath
    Next n
    GetTemplates.Path = App.Path & "\data\templates"
    GetTemplates.Refresh
    For n = 0 To GetTemplates.ListCount - 1
        lstfiles.ListItems.Add , , RightClip(GetTemplates.List(n), 4), 4, 4
    Next n
    lsRecent.ColumnHeaders(1).Width = lsRecent.Width - 70
    lstfiles.ListItems(1).Selected = True
    frmMain.sBar.Panels(2) = amNewFileWindow
End Sub

Private Sub cmbAct_Click(Index As Integer)
    'When you click on one of the three buttons, this selects what you want
    'to do. You can either view the help, choose okay, or choose cancel
    Dim FileName As String, Key As String
    Select Case Index
        Case 0: Am8.ShowHelp "New Window" 'Show help
        Case 1: Visible = False: Refresh: DoEvents: Unload Me 'Choose the cancel button
        Case 2
            Select Case lstfiles.SelectedItem.Index
                Case 1 'Create a new file and window
                    SetUpSBar
                    frmMain.CreateNewFileWithWindow
                    Unload Me
                    
                Case 2 'Open an existing file off
                    FileName = SelectFileName("Am8", amOpenFileName)
                    If FileName <> "" Then
                        SetUpSBar
                        frmMain.LoadExistingFileWithWindow FileName
                        Unload Me
                    End If
                    
                Case 3 'Inport a file from one of the avaliable file formats
                    FileName = SelectFileName("Import", amInportFileName)
                    If FileName <> "" Then
                        SetUpSBar
                        frmMain.CreateNewFileWithWindow
                        modImport.ImportModel FileName, Am8(ActiveFile)
                        frmMain.ActiveForm.Tablet.Refresh
                        Unload Me
                    End If
                    
                Case 4
                    frmCompile.RunAtStart
                    
                Case Is > 4 'Load one of the model templates as though it was creating a new model
                    FileName = App.Path & "\data\templates\" & lstfiles.SelectedItem.Text & ".am8"
                    Unload Me
                    frmMain.CreateNewFileWithWindow
                    Am8(ActiveFile).LoadFromFile FileName
                    Visible = False: Refresh: DoEvents
                    Am8(ActiveFile).ModelName = "Untitled"
                    Am8(ActiveFile).CurrentFilePath = ""
                    frmMain.ActiveForm.Caption = Am8(ActiveFile).ModelName
                    
                    
            
            End Select
    End Select
End Sub

Private Sub SetUpSBar()
    'This routine is called from several places, and just displays a message that somthing is being loaded
    'on the status bar, and updates the screen to shown this changes
    frmMain.sBar.Panels(2) = amPleaseWait
    Visible = False: DoEvents
    frmMain.sBar.Panels(2) = amReady
End Sub

Private Sub NewTab_Click()
    'This changes the frame that you can see when you click on the tabs along the top of the window
    Dim n As Byte
    For n = 0 To Frame.Count - 1: Frame(n).Visible = False: Next n
    Frame(NewTab.SelectedItem.Index - 1).Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'This unloads the form, and sets the timer on the main form to wait .01 of a second before resizing the form.
    'Otherwise, the form does not resize when it should, and it looks messed up
    frmMain.setResize.Interval = 10
End Sub

Private Sub lsRecent_DblClick()
    'When you double click on the recent history list, this loads the file that you clicked on
    frmMain.sBar.Panels(2) = amPleaseWait
    Visible = False: DoEvents
    frmMain.LoadExistingFileWithWindow Am8.FileHistory(lsRecent.SelectedItem.Index).FilePath
    Am8.FileHistory.AddHistory Am8.FileHistory(lsRecent.SelectedItem.Index).FilePath
    Unload Me
End Sub

Private Sub lstFiles_DblClick()
    'When you doube click on an icon in the list, it creates that type of file for you.
    cmbAct_Click 2
End Sub

