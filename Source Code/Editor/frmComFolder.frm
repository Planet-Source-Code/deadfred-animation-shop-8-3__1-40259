VERSION 5.00
Begin VB.Form frmComFolder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compile Folder"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5550
   Icon            =   "frmComFolder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   5550
   StartUpPosition =   1  'CenterOwner
   Begin VB.FileListBox flFiles 
      Height          =   2820
      Left            =   3240
      Pattern         =   "*.am8"
      TabIndex        =   6
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton cmdACT 
      Caption         =   "Help"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Click to get help on this window"
      Top             =   3360
      Width           =   1335
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   480
      Pattern         =   "*.am5"
      TabIndex        =   5
      Top             =   3480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdACT 
      Caption         =   "Okay"
      Default         =   -1  'True
      Height          =   375
      Index           =   2
      Left            =   4080
      TabIndex        =   4
      ToolTipText     =   "Click to compile all of the files in the selected folder"
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmdACT 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Index           =   0
      Left            =   2640
      TabIndex        =   3
      ToolTipText     =   "Click to close this window"
      Top             =   3360
      Width           =   1335
   End
   Begin VB.DriveListBox GetDrive 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   2895
   End
   Begin VB.DirListBox GetFolder 
      Height          =   2340
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5400
      Y1              =   3135
      Y2              =   3135
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   5400
      Y1              =   3150
      Y2              =   3150
   End
End
Attribute VB_Name = "frmComFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdACT_Click(Index As Integer)
    Dim FileName As String, NewFileName As String, Mes As String, TotalFiles As Integer
    Dim Resp As Integer, X As Integer, STime As Long, n As Integer, GoodCount As Integer
    Select Case Index
        Case 0
            Unload Me
            
        Case 2
            File1.Path = GetFolder.List(GetFolder.ListIndex)
            Mes = "Are you sure you want to compile every file in the following folder?" & vbNewLine & vbNewLine & GetFolder.List(GetFolder.ListIndex)
            If MsgBox(Mes, vbQuestion + vbYesNo) = vbNo Then Exit Sub


            TotalFiles = flFiles.ListCount - 1: ReDim filestore(flFiles.ListCount - 1)
            For n = 0 To flFiles.ListCount - 1: filestore(n) = flFiles.List(n): Next n

            Visible = False
            STime = Timer
            For n = 0 To flFiles.ListCount - 1
                FileName = flFiles.Path & "\" & filestore(n)
                Am8.File.Add "ComFolder" & n
                ActiveFile = "ComFolder" & n
                If Am8("ComFolder" & n).LoadFromFile(FileName) = True Then
                    GoodCount = GoodCount + 1
                    NewFileName = RightClip(FileName, 4) & ".dat"
                    frmTimer.RunAtStart Am8("ComFolder" & n), NewFileName, n, TotalFiles
                End If
                Am8.File.Remove "ComFolder" & n
            Next n


            MsgBox "Compile Folder complete" & vbNewLine & vbNewLine & _
                   GoodCount & " files succesfully compiled" & vbNewLine & _
                   "Opperation took " & Timer - STime & " seconds", vbInformation
            Unload frmTimer
            Unload Me
            
        Case 1
            Am8.ShowHelp "Main"
    
    End Select
End Sub


Public Sub RunAtStart()
    GetFolder.Path = App.Path
    GetDrive = App.Path
    Show vbModal
End Sub


Private Sub GetDrive_Change()
    GetFolder = GetDrive
End Sub


Private Sub GetFolder_Change()
    flFiles.Path = GetFolder.Path
End Sub


Private Sub GetFolder_Click()
    flFiles.Path = GetFolder.List(GetFolder.ListIndex)
End Sub
