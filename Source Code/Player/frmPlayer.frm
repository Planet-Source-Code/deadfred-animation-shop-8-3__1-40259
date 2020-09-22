VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Animation Viewer"
   ClientHeight    =   7095
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9015
   Icon            =   "frmPlayer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7095
   ScaleWidth      =   9015
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog OpenFile 
      Left            =   3360
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3360
      Top             =   2880
   End
   Begin VB.Menu menuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFile 
         Caption         =   "Open Animation"
         Index           =   1
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Exit"
         Index           =   3
      End
   End
   Begin VB.Menu menuFrameRate 
      Caption         =   "Framerate"
      Begin VB.Menu mnuFrame 
         Caption         =   "Fastest"
         Index           =   1
      End
      Begin VB.Menu mnuFrame 
         Caption         =   "20 / Sec"
         Index           =   2
      End
      Begin VB.Menu mnuFrame 
         Caption         =   "10 / Sec"
         Checked         =   -1  'True
         Index           =   3
      End
      Begin VB.Menu mnuFrame 
         Caption         =   "5 /Sec"
         Index           =   4
      End
      Begin VB.Menu mnuFrame 
         Caption         =   "2 / Sec"
         Index           =   5
      End
      Begin VB.Menu mnuFrame 
         Caption         =   "1 per second"
         Index           =   6
      End
      Begin VB.Menu mnuFrame 
         Caption         =   "2 seconds"
         Index           =   7
      End
      Begin VB.Menu mnuFrame 
         Caption         =   "5 seconds"
         Index           =   8
      End
   End
   Begin VB.Menu menuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "About"
         Index           =   1
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SceneName As String
Dim FrameIndex As Integer
Dim MaxFrame As Integer

Private Sub Form_Load()
    If Command = "" Then
        LoadAnimation
    Else
        If LoadAnimation = False Then End
    End If
End Sub


Private Function LoadAnimation() As Boolean
    OpenFile.DialogTitle = "Select a frame from the desired Animation"
    OpenFile.Filter = "Animation Frames (*.bmp)|*.bmp"
    OpenFile.FilterIndex = 1
    OpenFile.ShowOpen
    If OpenFile.FileName <> "" Then
        SceneName = Mid(OpenFile.FileName, 1, Len(OpenFile.FileName) - 7)
        Caption = "Animation Viewer [" & SceneName & "]"
        FrameIndex = 0
        MaxFrame = -1
        LoadAnimation = True
    End If
End Function


Private Sub mnuFile_Click(Index As Integer)
    Select Case Index
        Case 1: LoadAnimation
        Case 3: End
    End Select
End Sub

Private Sub mnuFrame_Click(Index As Integer)
    Dim n As Integer
    For n = 1 To 8: mnuFrame(n).Checked = False: Next n
    mnuFrame(Index).Checked = True
    Select Case Index
        Case 1: Timer1.Interval = 1
        Case 2: Timer1.Interval = 50
        Case 3: Timer1.Interval = 100
        Case 4: Timer1.Interval = 200
        Case 5: Timer1.Interval = 500
        Case 6: Timer1.Interval = 1000
        Case 7: Timer1.Interval = 2000
        Case 8: Timer1.Interval = 5000
    End Select
End Sub

Private Sub mnuHelp_Click(Index As Integer)
    MsgBox "Animation Viewer" & vbNewLine & vbNewLine & "Display animation sequences created with Animation Shop. See the Help file for details", vbInformation
End Sub

Private Sub Timer1_Timer()
    Dim ErrorCount As Integer
    If SceneName <> "" Then
        On Error GoTo Failed
        FrameIndex = FrameIndex + 1
        If FrameIndex = MaxFrame Then FrameIndex = 0
        Me.Picture = LoadPicture(SceneName & ThreeLength(FrameIndex) & ".bmp")
        Exit Sub
Failed:
        MaxFrame = FrameIndex
        FrameIndex = 0
        ErrorCount = ErrorCount + 1
        If ErrorCount >= 4 Then
            MsgBox Err.Description & vbNewLine & vbNewLine & "The animation could not be played, and has been de-activated", vbCritical
        End If
        Resume
    End If
End Sub


Public Function ThreeLength(Number As Integer) As String
    ThreeLength = Trim(Str(Number))
    If Len(ThreeLength) = 1 Then ThreeLength = "00" & ThreeLength
    If Len(ThreeLength) = 2 Then ThreeLength = "0" & ThreeLength
End Function

