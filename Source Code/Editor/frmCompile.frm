VERSION 5.00
Begin VB.Form frmCompile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compile Model"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5115
   Icon            =   "frmCompile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   5115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmbAct 
      Caption         =   "Cancel"
      Height          =   375
      Index           =   2
      Left            =   2640
      TabIndex        =   9
      ToolTipText     =   "Click close this window without compiling"
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmbAct 
      Caption         =   "Okay"
      Height          =   375
      Index           =   0
      Left            =   3840
      TabIndex        =   1
      ToolTipText     =   "Click to choose a filename and compile"
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmbAct 
      Caption         =   "Help"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Click to get help on using this window"
      Top             =   2880
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   3360
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   2655
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4815
      Begin VB.CheckBox chkCompile 
         Caption         =   "Include comments, starting with"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   7
         ToolTipText     =   "Designed to make understanding the format easier"
         Top             =   360
         Width           =   2655
      End
      Begin VB.TextBox txtCommentMark 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2880
         TabIndex        =   6
         Text            =   "//"
         ToolTipText     =   "This patern is used to identify the start of a commented line"
         Top             =   360
         Width           =   495
      End
      Begin VB.CheckBox chkCompile 
         Caption         =   "Vertex list starts at 1"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   5
         ToolTipText     =   "Starts the list at zero, instead of one"
         Top             =   840
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkCompile 
         Caption         =   "Include entities"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   4
         ToolTipText     =   "Includes entity objects, such as lights and doors"
         Top             =   1320
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkCompile 
         Caption         =   "Include skelital structure and links"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   3
         ToolTipText     =   "Allows the level to be more dynamic, with detailed movement"
         Top             =   1800
         Value           =   1  'Checked
         Width           =   3015
      End
   End
End
Attribute VB_Name = "frmCompile"
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



Public Sub RunAtStart(Optional AssignedFile As clsFile = Nothing)
    'This starts the form running
    If AssignedFile Is Nothing Then
        Caption = "Compile Folder"
        cmbAct(0).Caption = "Folder..."
    Else
        Set Model = AssignedFile
    End If
    Show vbModal
End Sub




Private Sub chkCompile_Click(Index As Integer)
    'This makes the text box for the comment marker enabled or disabled
    If chkCompile(0) = 1 Then
        txtCommentMark.Enabled = True
        txtCommentMark.BackColor = vbWhite
    Else
        txtCommentMark.Enabled = False
        txtCommentMark.BackColor = BackColor
    End If
End Sub




Private Sub cmbAct_Click(Index As Integer)
    'This checks to see which button you pressed, then does the right thing
    Dim FileName As String, Resp As Integer, StartTime As Long, Mess As String, X As String, c As String
    Select Case Index
    
        Case 0
            'Compile the model that is currently loaded.
            If Model Is Nothing Then
                frmComFolder.RunAtStart
            Else
                FileName = SetFileName("Compile", "Compile model to...")
                If FileName <> "" Then
                    If CheckOverwrite(FileName, 0) = False Then Exit Sub
                    Visible = False
                    StartTime = Timer
                    frmTimer.RunAtStart Model, FileName, 0, 0
                    Mess = "File created successfully :- " & vbNewLine & vbNewLine & FileName & vbNewLine & vbNewLine
                    Mess = Mess & "Time elapsed : "
                    Mess = Mess & Int((Timer - StartTime) * 100) / 100 & " seconds" & vbNewLine & vbNewLine
                    Mess = Mess & "As the created file is just text, you can open it in a simle editor, such as notepad." & vbNewLine & "Would you like to open it now?"
                    Resp = MsgBox(Mess, vbYesNo + vbInformation, "Compile Successfull")
                    If Resp = 6 Then
                        X = "notepad.exe": c = FileName
                        Shell "notepad.exe " + c, vbNormalFocus
                    End If
                End If
            End If
            
            
        Case 1
            'Show the help program
            Am8.ShowHelp "Main"
    
        Case 2
            'Cancel the window
            Unload Me
            
    End Select
End Sub









