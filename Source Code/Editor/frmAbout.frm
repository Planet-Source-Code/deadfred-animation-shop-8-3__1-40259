VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Animation Shop 8"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6465
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOkay 
      Cancel          =   -1  'True
      Caption         =   "Okay"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   $"frmAbout.frx":030A
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   6135
   End
   Begin VB.Label Label1 
      Caption         =   $"frmAbout.frx":0393
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   6135
   End
   Begin VB.Image AppLogo 
      BorderStyle     =   1  'Fixed Single
      Height          =   1005
      Left            =   120
      Picture         =   "frmAbout.frx":04A8
      Stretch         =   -1  'True
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ############################################################################
' #                                                                          #
' #   This 'About' box does nothing other that display itself, so you can    #
' #                       read the text on the form                          #
' #                                                                          #
' ############################################################################

Public Sub RunAtStart()
    'So that you can display the form from elsewhere
    Show vbModal
End Sub

Private Sub cmdOkay_Click()
    'This closes the About box
    Unload frmAbout
End Sub
