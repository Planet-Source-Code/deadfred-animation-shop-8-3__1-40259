VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl EnterGrid 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin MSComCtl2.UpDown Adjust 
      Height          =   735
      Left            =   2280
      TabIndex        =   4
      Top             =   1440
      Visible         =   0   'False
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   1296
      _Version        =   393216
      Max             =   360
      Min             =   -360
      Wrap            =   -1  'True
      Enabled         =   -1  'True
   End
   Begin VB.CheckBox ShowOption 
      Caption         =   "Scale"
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   3
      Top             =   3240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox ShowOption 
      Caption         =   "Origin"
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   2
      Top             =   3240
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CheckBox ShowOption 
      Caption         =   "Angle"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   1
      Top             =   3240
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   2990
      _Version        =   393216
      Rows            =   20
      ScrollTrack     =   -1  'True
      Appearance      =   0
   End
End
Attribute VB_Name = "EnterGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'##############################################################################################
'#                                                                                            #
'#                                    Display Frame Control                                   #
'#                                                                                            #
'#   This control is basicly a FlexiGrid control, but with a few extras around the outside.   #
'# Its main purpose is so that it can display the heading ('Origin', 'Angle' & 'Scale') in    #
'#  a merges cell, and still allow large selections to take place. It also does alot of the   #
'#  other messy work like aligning the updown arrow, and the more code that can be taken      #
'# out of the main form the better. This is OOP programming after all. Also, it allow you     #
'#         to have more than one of these controls, should you ever need it                   #
'#                                                                                            #
'##############################################################################################

Private IncreaseStepSize As Integer
Private LocalScene As String, LocalFrame As String

Private Function AlignArrow()
    'This positions the updown arrow control so that it is next to the selected cells on the flexi grid
    Dim x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer
    y1 = Grid.Row: x1 = Grid.col
    y2 = Grid.RowSel: x2 = Grid.ColSel
    If x1 > x2 Then Swap x1, x2
    If y1 > y2 Then Swap y1, y2
    Adjust.Top = Grid.RowPos(y1) + Grid.Top
    Adjust.Left = Grid.ColPos(x2) + Grid.ColWidth(x2) + Grid.Left
    Adjust.Height = Grid.RowPos(y2) - Grid.RowPos(y1) + Grid.RowHeight(y2)
    If Adjust.Left > Grid.ColWidth(0) And Adjust.Top + Adjust.Height > Grid.RowHeight(0) Then Adjust.Visible = True Else Adjust.Visible = False
End Function

Private Sub Adjust_Change()
    'This alters the value of the selected cells in the flexi grid. If you click
    'up, it increases each value by one, or 10 if you have held the mouse down
    'for a certain length of time.
    Dim x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer, n As Integer, m As Integer
    Static AvoidLooping As Boolean
    If AvoidLooping = True Then Exit Sub
    IncreaseStepSize = IncreaseStepSize + 1
    If IncreaseStepSize = 10 Then Adjust.Increment = 10
    y1 = Grid.Row: x1 = Grid.col
    y2 = Grid.RowSel: x2 = Grid.ColSel
    If x1 > x2 Then Swap x1, x2
    If y1 > y2 Then Swap y1, y2
    For n = x1 To x2
        For m = y1 To y2
            Grid.TextMatrix(m, n) = (Val(Grid.TextMatrix(m, n)) + Adjust.Value) Mod 360
            With Am8(ActiveFile).Scene(LocalScene)(LocalFrame)(m)
                Select Case n
                    Case 1: .AngleX = (Val(Grid.TextMatrix(m, n)) * 10)
                    Case 2: .AngleY = (Val(Grid.TextMatrix(m, n)) * 10)
                    Case 3: .AngleZ = (Val(Grid.TextMatrix(m, n)) * 10)
                    Case 4: .OriginX = Val(Grid.TextMatrix(m, n))
                    Case 5: .OriginY = Val(Grid.TextMatrix(m, n))
                    Case 6: .OriginZ = Val(Grid.TextMatrix(m, n))
                    Case 7: .ScaleX = Val(Grid.TextMatrix(m, n))
                    Case 8: .ScaleY = Val(Grid.TextMatrix(m, n))
                    Case 9: .ScaleZ = Val(Grid.TextMatrix(m, n))
                End Select
            End With
        Next m
    Next n
    Am8(ActiveFile).Scene.CopyToAnimate LocalScene, LocalFrame
    Am8(ActiveFile).Saved = False
    frmMain.ActiveForm.Engine.RefreshView
    AvoidLooping = True
    Adjust = 0
    AvoidLooping = False
End Sub

Private Sub Adjust_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    'When you release the mouse, it sets the step size of the control back to zero
    IncreaseStepSize = 0
    Adjust.Increment = 1
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
    'This checks to see if you have pressed the delete or zero key. If you press delete,
    'this removes all values from the selected cells. This is NOT the same as setting the
    'values to zero, which occurs when you press either of the zero buttons
    Dim x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer
    Dim n As Integer, m As Integer
    Select Case KeyCode
        Case 46, 48, 45, 96
            y1 = Grid.Row
            x1 = Grid.col
            y2 = Grid.RowSel
            x2 = Grid.ColSel
            If x1 > x2 Then Swap x1, x2
            If y1 > y2 Then Swap y1, y2
            For n = x1 To x2
                For m = y1 To y2
                    If KeyCode = 46 Then Grid.TextMatrix(m, n) = "" Else Grid.TextMatrix(m, n) = 0
                    With Am8(ActiveFile).Scene(LocalScene)(LocalFrame)(m)
                        Select Case n
                            Case 1: .AngleX = (Val(Grid.TextMatrix(m, n)) * 10)
                            Case 2: .AngleY = (Val(Grid.TextMatrix(m, n)) * 10)
                            Case 3: .AngleZ = (Val(Grid.TextMatrix(m, n)) * 10)
                            Case 4: .OriginX = Val(Grid.TextMatrix(m, n))
                            Case 5: .OriginY = Val(Grid.TextMatrix(m, n))
                            Case 6: .OriginZ = Val(Grid.TextMatrix(m, n))
                            Case 7: .ScaleX = Val(Grid.TextMatrix(m, n))
                            Case 8: .ScaleY = Val(Grid.TextMatrix(m, n))
                            Case 9: .ScaleZ = Val(Grid.TextMatrix(m, n))
                        End Select
                    End With
                Next m
            Next n
            frmMain.ActiveForm.Engine.RefreshView
    End Select
    Am8(ActiveFile).Saved = False
End Sub

Private Sub Grid_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    'This hides the updown scroll button
    Adjust.Visible = False
End Sub

Private Sub Grid_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    'This displays the arrow
    AlignArrow
End Sub

Private Sub Grid_Scroll()
    'This redisplay the arrow and headding when you scroll the grid around
    UserControl_Paint
    AlignArrow
End Sub

Private Sub ShowOption_Click(Index As Integer)
    'When you click on one of the three tick boxes, this checks to see if you have deselected
    'them all. If not, it redisplays the grid and headers
    If ShowOption(0) + ShowOption(1) + ShowOption(2) = 0 Then
        MsgBox amEnterGridShowSection, vbExclamation
        ShowOption(Index) = 1
    End If
    UpdateDisplay
    AlignArrow
End Sub

Public Sub UpdateDisplay(Optional SceneName As String = "", Optional FrameName As String = "")
    'This displays the contants of a frame in a given Flexi Grid control. You
    'can choose which of the Origin, Angle and Scale value sets are displayed
    'This sets up the grid and sub headdings (x,y,z), and sets the column width
    Dim n As Integer, Cols As Byte, stLine As String, Am As clsJointRow
    If SceneName = "" Then SceneName = LocalScene
    If FrameName = "" Then FrameName = LocalFrame
    LocalScene = SceneName: LocalFrame = FrameName
    Grid.Cols = 1 + (ShowOption(0) + ShowOption(1) + ShowOption(2)) * 3
    Grid.TextMatrix(0, 0) = ""
    For n = 1 To Grid.Cols - 1 Step 3
        Grid.TextMatrix(0, n) = "x"
        Grid.TextMatrix(0, n + 1) = "y"
        Grid.TextMatrix(0, n + 2) = "z"
    Next n
    For n = 1 To Grid.Cols - 1: Grid.ColWidth(n) = 400: Next n
    Grid.Visible = False: Grid.Rows = 2: Grid.FixedRows = 1
    If ActiveFile = "" Then Exit Sub
    For Each Am In Am8(ActiveFile).Scene(SceneName)(FrameName)
        stLine = ""
        If ShowOption(1) = 1 Then stLine = stLine & vbTab & Am.AngleX * 0.1 & vbTab & Am.AngleY * 0.1 & vbTab & Am.AngleZ * 0.1
        If ShowOption(2) = 1 Then stLine = stLine & vbTab & Am.OriginX & vbTab & Am.OriginY & vbTab & Am.OriginZ
        If ShowOption(0) = 1 Then stLine = stLine & vbTab & Am.ScaleX & vbTab & Am.ScaleY & vbTab & Am.ScaleZ
        Grid.AddItem Am8(ActiveFile).Joint(Am.Key).Name & stLine
    Next Am
    If Grid.Rows <> 2 Then Grid.RemoveItem 1
    Grid.Visible = True
    UserControl_Paint
End Sub

Private Function DrawHeading(C1 As Integer, c2 As Integer, Message As String) As Boolean
    'This draws a 3D border and message directly onto the control above the grid cells
    'given in the paramter call
    Const Slide As Integer = 17
    Line (Grid.ColPos(c2) + Grid.ColWidth(c2), Grid.Top)-(Grid.ColPos(C1) + Slide, 17), , B
    CurrentX = CurrentX + 90
    CurrentY = CurrentY + 15
    Print Message
    Line (Grid.ColPos(c2) + Grid.ColWidth(c2), 17)-(Grid.ColPos(C1) + Slide, 17), vbWhite
    Line (Grid.ColPos(C1) + Slide, 0)-(Grid.ColPos(C1) + Slide, Grid.Top), vbWhite
End Function

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    'This allows you to select all cells in the Origin, Angle or Scale group by
    'clicking on the heading above the group you want
    Dim ColON  As Integer, n As Integer
    ColON = 1
    If y < Grid.Top Then
        For n = 2 To 0 Step -1
            If ShowOption(n) = 1 Then
                If X > Grid.ColPos(ColON) And X < Grid.ColPos(ColON + 2) + Grid.ColWidth(ColON + 2) Then
                      Grid.Row = 1: Grid.col = ColON
                      Grid.RowSel = Grid.Rows - 1: Grid.ColSel = ColON + 2
                      AlignArrow
                End If
                ColON = ColON + 3
            End If
        Next n
    End If
End Sub

Private Sub UserControl_Paint()
    'This goes through each of the three tick boxes, and draws the headers for each of
    'the ticked boxes
    Dim ColON  As Integer
    Cls
    ColON = 1
    If ShowOption(2) = 1 Then DrawHeading ColON, ColON + 2, "Angle": ColON = ColON + 3
    If ShowOption(1) = 1 Then DrawHeading ColON, ColON + 2, "Origin": ColON = ColON + 3
    If ShowOption(0) = 1 Then DrawHeading ColON, ColON + 2, "Scale": ColON = ColON + 3
    Line (Grid.ColPos(ColON - 1) + Grid.ColWidth(ColON - 1), Grid.Top)-(Grid.ColPos(1), 0), , B
    Line (Grid.ColPos(0) + Grid.ColWidth(0) - 17, Grid.Top)-(Grid.ColPos(0), 0), BackColor, BF
End Sub

Private Sub UserControl_Resize()
    'This resizes the grid to fit the control
    On Error Resume Next
    Dim n As Integer
    Grid.Width = ScaleWidth: Grid.Height = ScaleHeight - Grid.Top - ShowOption(1).Height - 200
    For n = 0 To 2: ShowOption(n).Top = ScaleHeight - ShowOption(0).Height: Next n
End Sub

Public Property Get TextMatrix(X As Integer, y As Integer) As Variant
    'This returns the contents of the grid at the location specified
    TextMatrix = Grid.TextMatrix(X, y)
End Property
Public Property Let TextMatrix(X As Integer, y As Integer, ByVal vNewValue As Variant)
    Grid.TextMatrix(X, y) = vNewValue
End Property

Public Property Get Rows() As Integer
    'This returns the number of rows in the grid
    Rows = Grid.Rows
End Property
Public Property Let Rows(ByVal vNewValue As Integer)
    Grid.Rows = vNewValue + 1
End Property

Public Property Get ShowAngle() As Boolean
    'This allows you to determin whether the angles are displayed
    If ShowOption(2) = 1 Then ShowAngle = True Else ShowAngle = False
End Property
Public Property Let ShowAngle(ByVal vNewValue As Boolean)
    If vNewValue = True Then ShowOption(2) = 1 Else ShowOption(2) = 0
End Property

Public Property Get ShowOrigin() As Boolean
    'This allows you to determin whether the origins are displayed
    If ShowOption(1) = 1 Then ShowOrigin = True Else ShowOrigin = False
End Property
Public Property Let ShowOrigin(ByVal vNewValue As Boolean)
    If vNewValue = True Then ShowOption(1) = 1 Else ShowOption(1) = 0
End Property

Public Property Get ShowScale() As Boolean
    'This allows you to determin whether the scales are displayed
    If ShowOption(0) = 1 Then ShowScale = True Else ShowScale = False
End Property
Public Property Let ShowScale(ByVal vNewValue As Boolean)
    If vNewValue = True Then ShowOption(0) = 1 Else ShowOption(0) = 0
End Property

