VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.UserControl LayerDisplay 
   AutoRedraw      =   -1  'True
   ClientHeight    =   705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   705
   ScaleWidth      =   4800
   Begin ComCtl2.UpDown MoveTabs 
      Height          =   240
      Left            =   3720
      TabIndex        =   0
      Top             =   50
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   423
      _Version        =   327681
      Min             =   -10
      Orientation     =   1
      Enabled         =   -1  'True
   End
End
Attribute VB_Name = "LayerDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ############################################################################
' #                                                                          #
' #     This control is used to display the layers in a file, and to set     #
' #   whether they are shown or hidden. It works basicly like a tab stripe,  #
' #       only that you can select more than one tab at a time, etc..        #
' #                                                                          #
' ############################################################################

Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single, TabOver As Integer)
Dim Model As clsFile, StartPosition As Integer, TotalWidth As Integer, AvoidLoop As Boolean
Public LayerOver As Integer

Private Sub MoveTabs_Change()
    'Changes the position where the layers are drawn. It allows you to see more tabs that
    'could fit on the screen
    Dim n As Integer
    If AvoidLoop = True Then Exit Sub
    Select Case MoveTabs.Value
        Case -1: For n = 1 To 10: If TotalWidth < Width - 500 Then Exit Sub
                 StartPosition = StartPosition - 60: Refresh: DoEvents: Next n
        Case 1:  For n = 1 To 10: If StartPosition > 90 Then StartPosition = 90: Exit Sub
                 StartPosition = StartPosition + 60: Refresh: DoEvents: Next n
    End Select
    AvoidLoop = True: MoveTabs.Value = 0: AvoidLoop = False
End Sub

Private Function Refresh(Optional MousePosition As Single = -1) As Integer
    'Draws the tabs. Also checks to see which tab the mouse is over, when you pass a horizontal
    'value to the function
    Dim Am As clsLayer, Count As Integer, OldX As Single, NewX As Single
    Cls
    CurrentX = StartPosition
    If Model Is Nothing Then Exit Function
    For Each Am In Model.Layers
        Count = Count + 1:          CurrentY = 50
        OldX = CurrentX:            ForeColor = BackColor
        Print Am.LayerName;:        ForeColor = Am.LayerColour
        Font.Italic = False:        NewX = CurrentX
        Font.Strikethrough = False
        If Am.Default = True Then Font.Italic = True Else Font.Italic = False
        If MousePosition > OldX - 30 And MousePosition < NewX + 130 Then LayerOver = Count: Refresh = Count
        If Am.LayerGrayed = True Then ForeColor = vbGrayText
        If Am.LayerLocked = True Then Font.Strikethrough = True
        If Am.Selected = True Then
            CurrentX = OldX + 17
            CurrentY = 57
            Print Am.LayerName;
            Line (OldX - 60, 0)-(OldX - 60, 267), &H80000010
            Line (OldX - 77, 0)-(OldX - 77, 267), 0
            Line (NewX + 80, 250)-(NewX + 80, 0), &H80000014
            Line (OldX - 60, 250)-(NewX + 97, 250), &H80000014
        Else
            CurrentX = OldX
            CurrentY = 57
            Print Am.LayerName;
            CurrentX = CurrentX + 17
            Line (OldX - 60, 250)-(NewX + 97, 250), &H80000010
            Line (OldX - 47, 267)-(NewX + 97, 267), 0
            Line (OldX - 60, 0)-(OldX - 60, 267), &H80000014
            Line (NewX + 80, 250)-(NewX + 80, 0), &H80000010
            Line (NewX + 97, 250)-(NewX + 97, 0), 0
        End If
        CurrentX = CurrentX + 130
    Next Am
    TotalWidth = CurrentX
    If CurrentX - StartPosition > ScaleWidth - 400 Then
        MoveTabs.Visible = True: MoveTabs.Left = ScaleWidth - MoveTabs.Width
    Else
        MoveTabs.Visible = False
    End If
End Function

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'When you click on a Tab, this checks which on it was, and then sets it to either seleccted or not selected.
    Dim TabOver As Integer, n As Integer
    TabOver = Refresh(x)
    If Button = 1 Then
        If TabOver <> 0 Then
            If Model.Layers(TabOver).Selected = True Then Model.Layers(TabOver).Selected = False Else Model.Layers(TabOver).Selected = True
            For n = 1 To Model.Layers.CountLayers: Model.Layers(n).Default = False: Next n
            Model.Layers(TabOver).Default = True: Refresh
        End If
    End If
    RaiseEvent MouseDown(Button, Shift, x, y, TabOver)
End Sub

Public Sub AssignLayerDisplayTo(AssignedModel As clsFile)
    'This allows the control to be pointed at a file class
    Set Model = AssignedModel
End Sub

Public Sub Update()
    'Allows the users to redraw what is on the screen
    Refresh
End Sub

Private Sub UserControl_Initialize()
    'Seta the collection
    StartPosition = 90: Refresh
End Sub

