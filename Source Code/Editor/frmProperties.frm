VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Model Properties"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6555
   Icon            =   "frmProperties.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   6555
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2040
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProperties.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProperties.frx":0896
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProperties.frx":0BB2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdAct 
      Caption         =   "Help"
      Height          =   350
      Index           =   1
      Left            =   60
      TabIndex        =   2
      ToolTipText     =   "Click to get help on using this window"
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdAct 
      Cancel          =   -1  'True
      Caption         =   "Okay"
      Default         =   -1  'True
      Height          =   350
      Index           =   2
      Left            =   5400
      TabIndex        =   1
      ToolTipText     =   "Click to close this window and save your changes"
      Top             =   4440
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid gdStats 
      Height          =   3495
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   6165
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FixedCols       =   0
      GridLines       =   0
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin MSComctlLib.TabStrip ViewTab 
      Height          =   4275
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   7541
      HotTracking     =   -1  'True
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Statistics"
            Object.ToolTipText     =   "Details about the model"
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Model As clsFile

'#####################################################################
'#                                                                   #
'# This is the propertys window. It just shows a load of stats and   #
'#   numbers about the current model. It also allows you to enter    #
'#  some custom values for the model, that get saved along with the  #
'# model when you export it. You can also set the model notes, that  #
'#         can be displayed whenever the model is opened             #
'#                                                                   #
'#####################################################################

Public Sub RunAtStart(AssignedFile As clsFile)
    'This is the start up function for the form, and it just
    'puts aload of values into the different form objects
    Set Model = AssignedFile
    gdStats.ColWidth(0) = 0
    gdStats.ColWidth(1) = 1500
    gdStats.ColWidth(2) = 4300
    gdStats.TextMatrix(0, 1) = "Property"
    gdStats.TextMatrix(0, 2) = "Value"
    gdStats.AddItem vbTab & "File Path" & vbTab & " " & Model.CurrentFilePath
    gdStats.AddItem vbTab & "Object Count" & vbTab & " " & Model.Geometery.CountObjects
    gdStats.AddItem vbTab & "Face Count" & vbTab & " " & Model.Geometery.CountFaces
    gdStats.AddItem vbTab & "Vertex Count" & vbTab & " " & Model.Geometery.CountVertecies
    gdStats.AddItem vbTab & "Joint Count" & vbTab & " " & Model.Joint.CountChildren
    gdStats.AddItem vbTab & "Scene Count" & vbTab & " " & Model.Scene.SceneCount
    Show vbModal
End Sub

Private Sub cmdAct_Click(Index As Integer)
    'This controls the 3 buttons at the botom of the form, either Help, Cancel or Okay
    Select Case Index
        Case 1: Am8.ShowHelp "File Properties Window"
        Case 2: Unload Me
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'I dont know if this is nessessary, but it makes sure that the model object is emptied
    Set Model = Nothing
End Sub
