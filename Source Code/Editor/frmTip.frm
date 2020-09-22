VERSION 5.00
Begin VB.Form frmTip 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tip of the Day"
   ClientHeight    =   3285
   ClientLeft      =   2355
   ClientTop       =   2385
   ClientWidth     =   5415
   Icon            =   "frmTip.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkLoadTipsAtStartup 
      Caption         =   "&Show Tips at Startup"
      Height          =   315
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Tick to cause this box to appear when you start Animation Shop"
      Top             =   2880
      Width           =   2055
   End
   Begin VB.CommandButton cmdNextTip 
      Caption         =   "&Next Tip"
      Height          =   350
      Left            =   4080
      TabIndex        =   2
      ToolTipText     =   "Click to show the next tip"
      Top             =   600
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   2715
      Left            =   120
      Picture         =   "frmTip.frx":0442
      ScaleHeight     =   2655
      ScaleWidth      =   3675
      TabIndex        =   1
      Top             =   120
      Width           =   3735
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Did you know..."
         Height          =   255
         Left            =   540
         TabIndex        =   5
         Top             =   180
         Width           =   2655
      End
      Begin VB.Label lblTipText 
         BackColor       =   &H00FFFFFF&
         Height          =   1635
         Left            =   180
         TabIndex        =   4
         Top             =   840
         Width           =   3255
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   350
      Left            =   4080
      TabIndex        =   0
      ToolTipText     =   "Click to close this window"
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ############################################################################
' #                                                                          #
' #     This is the Tips form, which simply displays a list of helpful       #
' #   hints and tips for new users of Animation Shop. Its basicly the Tip-   #
' #            of-the-day template supplied with visual basic                #
' #                                                                          #
' ############################################################################

Dim Tips As New Collection

Public Sub RunAtStart()
    'When this sub is called from another form/modlule, it loads the Tips file, and displays the form
    Randomize: If Am8.ShowTips = True Then chkLoadTipsAtStartup = 1
    LoadTips App.Path & "\data\TipOfDay.dat": Show vbModal
End Sub

Private Sub DoNextTip()
    'This function reads a random tip from the collection, and displays it in the window
    If Tips.Count > 0 Then lblTipText.Caption = Tips.Item(Int((Tips.Count * Rnd) + 1))
End Sub

Private Function LoadTips(sFile As String) As Boolean
    'This function loads the Tips file into the collection. If the file is empty or not there, suitable
    'error handling takes place to avoid crashing the program
    Dim NextTip As String, InFile As Integer
    InFile = FreeFile
    If sFile = "" Then Exit Function
    If Dir(sFile) = "" Then Exit Function
    Open sFile For Input As InFile
    While Not EOF(InFile)
        Input #InFile, NextTip
        Tips.Add NextTip
    Wend
    Close InFile
    DoNextTip
    LoadTips = True
End Function

Private Sub chkLoadTipsAtStartup_Click()
    'This sets the Show Tip value to be saved in the program settings
    Am8.ShowTips = chkLoadTipsAtStartup
End Sub

Private Sub cmdNextTip_Click()
    'This button shows another tip
    DoNextTip
End Sub

Private Sub cmdOK_Click()
    'This unloads the form
    Unload Me
End Sub
