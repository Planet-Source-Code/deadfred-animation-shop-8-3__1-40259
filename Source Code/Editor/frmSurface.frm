VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSurface 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Surface Properties"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   Icon            =   "frmSurface.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList TabIcon 
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
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSurface.frx":0442
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdAct 
      Caption         =   "&Help"
      Height          =   350
      Index           =   0
      Left            =   120
      TabIndex        =   16
      ToolTipText     =   "Click to get help on using this window"
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdAct 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   350
      Index           =   1
      Left            =   5160
      TabIndex        =   15
      ToolTipText     =   "Click to close this window without saving your changes"
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdAct 
      Caption         =   "Okay"
      Default         =   -1  'True
      Height          =   350
      Index           =   2
      Left            =   3960
      TabIndex        =   14
      ToolTipText     =   "Click to close this window and save your changes"
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   5895
      Begin Project1.Engine ShowExample 
         Height          =   1695
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   2990
      End
      Begin VB.CheckBox chkBack 
         Alignment       =   1  'Right Justify
         Caption         =   "&Always draw hidden faces"
         Height          =   255
         Left            =   3360
         TabIndex        =   13
         Top             =   1560
         Width           =   2295
      End
      Begin MSComctlLib.Slider sldGrain 
         Height          =   255
         Left            =   3720
         TabIndex        =   11
         ToolTipText     =   "Gives the object a rough, grainy appearence"
         Top             =   3000
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         _Version        =   393216
         Max             =   100
         TickFrequency   =   10
      End
      Begin VB.PictureBox picColour 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3720
         ScaleHeight     =   225
         ScaleWidth      =   2025
         TabIndex        =   7
         ToolTipText     =   "The normal colour of the object"
         Top             =   480
         Width           =   2055
      End
      Begin VB.ComboBox cmbTexture 
         Height          =   315
         Left            =   3720
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         ToolTipText     =   "Lists the avalible textures"
         Top             =   1080
         Width           =   2055
      End
      Begin MSComctlLib.Slider sldTransparant 
         Height          =   255
         Left            =   3720
         TabIndex        =   3
         ToolTipText     =   "Sets the level of transparancy for the object"
         Top             =   2040
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         _Version        =   393216
         Max             =   100
         TickFrequency   =   10
      End
      Begin MSComctlLib.Slider sldDiffuse 
         Height          =   255
         Left            =   3720
         TabIndex        =   5
         ToolTipText     =   "Breaks up the edges of an object to give the effect of dust"
         Top             =   2520
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         _Version        =   393216
         Max             =   100
         TickFrequency   =   10
      End
      Begin VB.Label Label7 
         Caption         =   "Preview"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "&Grain"
         Height          =   255
         Left            =   3120
         TabIndex        =   8
         Top             =   3000
         Width           =   495
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "&Colour"
         Height          =   255
         Left            =   3120
         TabIndex        =   6
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "&Diffusion"
         Height          =   255
         Left            =   2880
         TabIndex        =   4
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "&Texture"
         Height          =   255
         Left            =   3000
         TabIndex        =   9
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Trans&parancy"
         Height          =   255
         Left            =   2640
         TabIndex        =   2
         Top             =   2040
         Width           =   975
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   7646
      HotTracking     =   -1  'True
      ImageList       =   "TabIcon"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Surface Properties"
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSurface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Model As clsFile
Dim Example As clsFile
Dim Angle1 As Single, Angle2 As Single, Angle3 As Single

Sub RunAtStart(AssignedFile As clsFile)
    'This is the fucnction used to display the window. The settings are set to show the
    'properties of the first selected object
    If AssignedFile.Geometery.CountSelected = 0 Then MsgBox amMustSelectObject, vbInformation: Exit Sub
    Set Model = AssignedFile
    Set Example = New clsFile
    With Model.Geometery(Model.Geometery.FirstSelectedObject)
        picColour.BackColor = .Colour
        sldGrain.Value = .grain
        sldDiffuse.Value = .Diffusion
        sldTransparant.Value = .Transparancy
        chkBack = IntBo(.ForceShowFace)
        Example.Geometery.CreateObject "Example"
        Example.Geometery("Example").CreateObject "Cube", 1, 32, -32, 32, -32, 32, -32
        Example.Geometery("Example").Colour = .Colour
        Example.Geometery("Example").grain = .grain
        Example.Geometery("Example").Diffusion = .Diffusion
        Example.Geometery("Example").Transparancy = .Transparancy
        Example.Geometery("Example").Layer = "Main"
    End With
    Example.Layers.AddLayer "Main", "Main"
    ShowExample.AssignEngineTo Example
    ShowExample.pAutoRotate = True
    ShowExample.pPerspecitve = True
    ShowExample.pRenderSolid = True
    ShowExample.pClipLine = -800
    ShowExample.ShapeFX = 5
    ShowExample.pDrawObjects = True
    ShowExample.RefreshView
    Show vbModal
End Sub

Private Sub cmbTexture_Click()
    'This code runs when you click on the textures dropdown list. If you select the -Add- item,
    'an Open File window is shown, alowing you to select and add a bitmap to the dropdown list
    Dim FileName As String, n As Integer, Found As Boolean
    If cmbTexture.ListIndex = 0 Then
        FileName = SelectFileName("Picture", "Select a texture file")
        If FileName <> "" Then
            For n = 0 To cmbTexture.ListCount
                If cmbTexture.List(n) = FileName Then Found = True
            Next n
            If Found = False Then cmbTexture.AddItem FileName
            For n = 0 To cmbTexture.ListCount
                If cmbTexture.List(n) = FileName Then cmbTexture.ListIndex = n
            Next n
        Else
            cmbTexture.ListIndex = 1
        End If
    End If
End Sub

Private Sub cmdACT_Click(Index As Integer)
    'This controls the three buttons along the bottom, Help, Cancel and Ok
    Dim Am As clsObject
    Select Case Index
        Case 0: Am8.ShowHelp "Surface Properties window"
        Case 1: Unload Me
        Case 2
            For Each Am In Model.Geometery
                If Am.Selected = True Then
                    Am.Colour = picColour.BackColor
                    Am.grain = sldGrain
                    Am.Diffusion = sldDiffuse
                    Am.ForceShowFace = IntBo(chkBack)
                    Am.Transparancy = sldTransparant
                End If
            Next Am
            Unload Me
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'This gets rid of the example object, or else it would suddenly appear in your file
    Set Example = Nothing
End Sub

Private Sub picColour_Click()
    'When you alter the colour box, this sets the example object to the same
    frmMain.GetFile.ShowColor
    picColour.BackColor = frmMain.GetFile.Color
    Example.Geometery("Example").Colour = frmMain.GetFile.Color
End Sub

Private Sub sldDiffuse_Validate(Cancel As Boolean)
    'When you alter the diffuse slider, this sets the example object to the same
    Example.Geometery("Example").Diffusion = sldDiffuse
End Sub

Private Sub sldGrain_Validate(Cancel As Boolean)
    'When you alter the grain slider, this sets the example object to the same
    Example.Geometery("Example").grain = sldGrain
End Sub

Private Sub sldTransparant_Validate(Cancel As Boolean)
    'When you alter the transparancy slider, this sets the example object to the same
    Example.Geometery("Example").Transparancy = sldTransparant
End Sub
