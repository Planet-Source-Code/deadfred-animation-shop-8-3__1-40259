VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl Tablet 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.CommandButton cmdBlock 
      Enabled         =   0   'False
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   1080
      Width           =   255
   End
   Begin MSComCtl2.FlatScrollBar sbBar 
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Arrows          =   65536
      LargeChange     =   50
      Min             =   -1000
      Max             =   1000
      Orientation     =   8323073
      SmallChange     =   10
   End
   Begin MSComCtl2.FlatScrollBar sbBar 
      Height          =   1095
      Index           =   1
      Left            =   1080
      TabIndex        =   1
      Top             =   0
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   1931
      _Version        =   393216
      Appearance      =   0
      LargeChange     =   50
      Min             =   -1000
      Max             =   1000
      Orientation     =   8323072
      SmallChange     =   10
   End
   Begin VB.Line gguide 
      Visible         =   0   'False
      X1              =   96
      X2              =   176
      Y1              =   136
      Y2              =   192
   End
End
Attribute VB_Name = "Tablet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'#################################################################
'#                                                               #
'#  This control allows you to view and edit file objects. It    #
'#  contains loads of code for editing the files etc..           #
'#                                                               #
'#################################################################

Private Model As clsFile

Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long

Public Event MouseDown(X As Single, Y As Single, Button As Integer, Shift As Integer)
Public Event MouseMove(X As Single, Y As Single, Button As Integer, Shift As Integer)
Public Event UpdateOtherWindows()

Private API As POINTAPI
Private StartX As Single, StartY As Single
Private Xf As Integer, Yf As Integer
Private StartBoxSelect As Boolean
Private ObjectScaleMode As Byte
Private ObjectRotateMode As Boolean

Public ShapeX1 As Integer, ShapeX2 As Integer, ShapeY1 As Integer, ShapeY2 As Integer
Public pShowVertecies As Boolean
Public pShowFaces As Boolean
Public pShowGrid As Boolean
Public pEnableEdit As Boolean
Public ViewMode As Integer
Public lZoomLevel As Single
Public JointOver As String

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Public LastZoomLevel As Single


Public Property Get ZoomLevel() As Single
    ZoomLevel = lZoomLevel
End Property

Public Property Let ZoomLevel(ByVal vNewValue As Single)
    LastZoomLevel = lZoomLevel
    lZoomLevel = vNewValue
End Property



Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'This handles Drag and Drop from the Gallarys list box
    Dim XX As Single, YY As Single, Zz As Single
    If frmMain.Gallary.DraggedItemName <> "" Then
        Select Case ViewMode
            Case 1: XX = AbsoluteX(Int(X)): Zz = AbsoluteY(Int(Y))
            Case 2: XX = AbsoluteX(Int(X)): YY = AbsoluteY(Int(Y))
            Case 3: Zz = AbsoluteX(Int(X)): YY = AbsoluteY(Int(Y))
        End Select
        Model.LoadFromFile App.Path & "\data\gallarys\" & frmMain.cmbGallary.List(frmMain.cmbGallary.ListIndex) & "\" & frmMain.Gallary.DraggedItemName, 1, XX, YY, Zz
        Refresh
        frmMain.Gallary.DraggedItemName = ""
    End If
    Model.Saved = False
    'RaiseEvent SendMouseUp(X, Y, Button, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim DontExitSub As Boolean, Am As clsObject, ObjectOver As Integer
    Dim VertexOver As Integer, JOver As Integer, FoundJoint As Boolean
    Dim XX As Integer, YY As Integer, Zz As Integer, JKeyOver As String
    RaiseEvent MouseDown(X, Y, Button, Shift)
    If pEnableEdit = True Then
        gGuide.x1 = X: gGuide.y1 = Y
        gGuide.x2 = X: gGuide.y2 = Y
        StartX = X: StartY = Y
        ObjectScaleMode = SelectionOutlineHittest(X, Y)
        If Button = 1 Then
            
            If EditButton = 2 And frmMain.cmdSidebar.SelectedItem.Index = 2 Then
                If frmMain.ShapeList.SelectedItem.Text = "Wrap" Then
                    Model.Wraper.Add AbsoluteX(X), AbsoluteY(Y)
                    DrawShapeGuide frmMain.ShapeList.SelectedItem.Text, ShapeX1, ShapeY1, ShapeX2, ShapeY2
                    Exit Sub
                End If
            End If
            
            
            If EditButton = 3 Then
                Select Case GetEditOption
                    
                    Case 6
                        ObjectOver = FaceObjectHitTest(X, Y)
                        If ObjectOver <> 0 Then
                            VertexOver = FaceHitTest(Model.Geometery(ObjectOver), X, Y)
                            If VertexOver <> 0 Then
                                Model.Geometery(ObjectOver).Face.Remove VertexOver
                            End If
                        End If
                        Model.FindModelOutline
                        Refresh
                        Exit Sub
                    
                    Case 7
                        ObjectOver = ObjectHitTest(X, Y, True)
                        If ObjectOver <> 0 Then
                            VertexOver = VertexHitTest(Model.Geometery(ObjectOver), X, Y)
                            If VertexOver <> 0 Then
                                Model.Geometery(ObjectOver).DeleteVertecies VertexOver
                            End If
                        End If
                        Model.FindModelOutline
                        Refresh
                        Exit Sub
                    
                    Case 8
                        Select Case ViewMode
                            Case 1: XX = AbsoluteX(X): Zz = AbsoluteY(Y)
                            Case 2: XX = AbsoluteX(X): YY = AbsoluteY(Y)
                            Case 3: YY = AbsoluteX(X): Zz = AbsoluteY(Y)
                        End Select
                        Model.Geometery(Model.Geometery.FirstSelectedObject).Vertex.Add XX, YY, Zz
                        Model.Geometery(Model.Geometery.FirstSelectedObject).FindObjectOutline
                        Model.FindModelOutline
                        Refresh
                        Exit Sub
            
                    Case 13, 14
                        gGuide.x1 = X
                        gGuide.y1 = Y
              
                End Select
            End If
            
            
            If EditButton = 4 Then
                Select Case ObjectScaleMode
                    Case 4, 2: MousePointer = 6
                    Case 3, 1: MousePointer = 8
                    Case 5, 6: MousePointer = 9
                    Case 7, 8: MousePointer = 7
                End Select
                If ObjectScaleMode = 0 Then DontExitSub = True
                Refresh
                AutoRedraw = False
                If DontExitSub = False Then StartBoxSelect = False: Exit Sub
            End If
            If EditButton = 5 Then
                ObjectRotateMode = True
                AutoRedraw = False
            End If
            If frmMain.chkSelect(6) = 0 Then
                If InsideBoundingBox(X, Y) = False And Shift = 0 Then Model.DeselectAll
                If InsideBoundingBox(X, Y) = False Then StartBoxSelect = True
                ObjectOver = ObjectHitTest(X, Y)
                JKeyOver = JointHitTest(X, Y)
                If InsideBoundingBox(X, Y) = True And ObjectScaleMode = 0 Then MousePointer = 5
                If ObjectOver <> 0 And frmMain.chkSelect(1) = 1 Then
                    Model.Geometery(ObjectOver).Selected = True
                    If Model.Geometery(ObjectOver).Group.Count > 0 Then
                        For Each Am In Model.Geometery
                            If Am.Group.Count > 0 Then
                                If Am.Group(1).GroupID = Model.Geometery(ObjectOver).Group(1).GroupID Then
                                    Am.Selected = True
                                End If
                            End If
                        Next Am
                    End If
                    StartBoxSelect = False
                ElseIf JKeyOver <> "" Then
                    Model.Joint(JKeyOver).Selected = True
                    StartBoxSelect = False
                ElseIf frmMain.opChangeJ = True Then
                    gGuide.Visible = True
                End If
            Else

                If InsideBoundingBox(X, Y) = False Then Model.Geometery.DeselectAllVertecies
                FoundJoint = False
                For Each Am In Model.Geometery
                    If Am.Selected = True Then
                        VertexOver = VertexHitTest(Am, X, Y)
                        If VertexOver > 0 Then
                            If Am.Vertex(VertexOver).Selected = True Then Am.Vertex(VertexOver).Selected = False Else Am.Vertex(VertexOver).Selected = True
                            FoundJoint = True
                        End If
                    End If
                Next Am
                If Shift = 0 And FoundJoint = False Then StartBoxSelect = True
            End If
        End If
        Model.FindModelOutline
        AutoRedraw = False
        DrawOutline 0, 0
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim CenterX As Integer, CenterY As Integer, RotateMode As Byte
    If pEnableEdit = True Then
        If Button = 1 Then
            If EditButton = 2 And frmMain.cmdSidebar.SelectedItem.Index = 2 Then
                ShapeX1 = X: ShapeY1 = Y
                ShapeX2 = StartX: ShapeY2 = StartY
                DrawShapeGuide frmMain.ShapeList.SelectedItem.Text, ShapeX1, ShapeY1, ShapeX2, ShapeY2
                
            ElseIf EditButton = 6 And frmMain.cmdSidebar.SelectedItem.Index = 1 And frmMain.opChangeJ = True Then
                gGuide.x2 = X
                gGuide.y2 = Y
            ElseIf EditButton = 3 And GetEditOption > 2 Then
                Select Case GetEditOption
                End Select
            Else
                If EditButton = 5 And InsideBoundingBox(StartX, StartY) = True Then
                    If frmMain.optGetCenter(1) = True Then RotateMode = 1
                    If frmMain.optGetCenter(3) = True Then RotateMode = 2
                    If frmMain.optGetCenter(2) = True Then RotateMode = 3
                    frmMain.sBar.Panels(2) = "Rotate to " & GetAngle(Snaped(StartX - X), Snaped(Y - StartY)) * 0.1 & " *"
                    DrawRotateGuide GetAngle(Snaped(StartX - X), Snaped(Y - StartY)), CenterX, CenterY, RotateMode
                Else
                    If StartBoxSelect = True Then
                        DrawBoxBand Int(X), Int(Y), Int(StartX), Int(StartY)
                    Else
                        If StartBoxSelect = False And EditButton = 4 And ObjectScaleMode > 0 Then
                            DrawScaleOutline ObjectScaleMode, Int(X), Int(Y)
                        Else
                            MousePointer = 5
                            AutoRedraw = False
                            Cls
                            If X < -25 Then sbBar(0) = sbBar(0) - 20: StartX = StartX + 20: AutoRedraw = True: Refresh: AutoRedraw = False
                            If X > ScaleWidth + 25 Then sbBar(0) = sbBar(0) + 20: StartX = StartX - 20: AutoRedraw = True: Refresh: AutoRedraw = False
                            If Y < -25 Then sbBar(1) = sbBar(1) - 20: StartY = StartY + 20: AutoRedraw = True: Refresh: AutoRedraw = False
                            If Y > ScaleHeight + 25 Then sbBar(1) = sbBar(1) + 20: StartY = StartY - 20: AutoRedraw = True: Refresh: AutoRedraw = False
                            DrawOutline Snaped(X - StartX), Snaped(Y - StartY)
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Public Sub DrawOutline(ByVal X As Single, ByVal Y As Single)
    'This draws the ouline of a shape, with the X and Y values offsetting the
    'outline by so much, so that you can drag the outline around the screen
    Dim Am As clsObject
    Cls
    X = X / ZoomLevel
    Y = Y / ZoomLevel
    DrawStyle = 2
    ForeColor = vbBlue
    If Model.Geometery.FirstSelectedObject + Model.Joint.FirstSelectedJoint <> 0 Then
        Select Case ViewMode
            Case 1: DrawBox Model.MinX + X, Model.MinZ + Y, Model.MaxX + X, Model.MaxZ + Y, 1
            Case 2: DrawBox Model.MinX + X, Model.MinY + Y, Model.MaxX + X, Model.MaxY + Y, 1
            Case 3: DrawBox Model.MinZ + X, Model.MinY + Y, Model.MaxZ + X, Model.MaxY + Y, 1
        End Select
    End If
    DrawStyle = 0
End Sub

Private Sub ShowScaleError()
    MsgBox amDivisionByZero, vbExclamation
    Refresh
End Sub


Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim VertexOver As Integer, ObjectOver As Integer, JOver As Integer, n As Integer, Sye As Single
    If pEnableEdit = True Then
        MousePointer = 0
        AutoRedraw = True
        Dim Am As clsObject, jKey As String, tkey As String, NFace As Integer
        Dim sX As Single, sY As Single, AbsX As Integer, AbsY As Integer
        Dim Cx As Single, Cy As Single, Cz As Single, Jm As clsJoint
        If Button = 1 Then
            If EditButton = 2 And frmMain.cmdSidebar.SelectedItem.Index = 2 Then
            ElseIf EditButton = 3 And GetEditOption > 2 Then
                Select Case GetEditOption
                    Case 13
                        Do
                            ObjectOver = ObjectHitTest(gGuide.x1, gGuide.y1, False)
                            If ObjectOver <> 0 Then
                                Select Case ViewMode
                                    Case 1: Cx = Snaped(X - gGuide.x1): Cz = Snaped(Y - gGuide.y1)
                                    Case 2: Cx = Snaped(X - gGuide.x1): Cy = Snaped(Y - gGuide.y1)
                                    Case 3: Cz = Snaped(X - gGuide.x1): Cx = Snaped(Y - gGuide.y1)
                                End Select
                                VertexOver = VertexHitTest(Model.Geometery(ObjectOver), gGuide.x1, gGuide.y1)
                                If VertexOver <> 0 Then
                                    Model.Geometery(ObjectOver).Vertex(VertexOver).Move Int(Cx), Int(Cy), Int(Cz)
                                    Model.Geometery(ObjectOver).FindObjectOutline
                                End If
                            End If
                        Loop Until ObjectOver = 0 Or Shift = 1
                        Model.FindModelOutline
                        Refresh
                    
                    
                    Case 11
                        ObjectOver = FaceObjectHitTest(gGuide.x1, gGuide.y1)
                        If ObjectOver <> 0 Then
                            VertexOver = FaceHitTest(Model.Geometery(ObjectOver), gGuide.x1, gGuide.y1)
                            If VertexOver <> 0 Then
                                NFace = Model.Geometery(ObjectOver).FragmentFace(VertexOver, 1)
                                Select Case ViewMode
                                    Case 1: Model.Geometery(ObjectOver).Vertex(NFace).Move (X - gGuide.x1) / frmMain.sldExtend(0), 0, (Y - gGuide.y1) / frmMain.sldExtend(0)
                                    Case 2: Model.Geometery(ObjectOver).Vertex(NFace).Move (X - gGuide.x1) / frmMain.sldExtend(0), (Y - gGuide.y1) / frmMain.sldExtend(0), 0
                                    Case 3: Model.Geometery(ObjectOver).Vertex(NFace).Move 0, (Y - gGuide.y1) / frmMain.sldExtend(0), (X - gGuide.x1) / frmMain.sldExtend(0)
                                End Select
                                Model.Geometery(ObjectOver).FindObjectOutline
                            End If
                        End If
                        Model.FindModelOutline
                        Refresh
                        Exit Sub
                    
                    
                    Case 12
                        ObjectOver = FaceObjectHitTest(gGuide.x1, gGuide.y1)
                        If ObjectOver <> 0 Then
                            VertexOver = FaceHitTest(Model.Geometery(ObjectOver), gGuide.x1, gGuide.y1)
                            If VertexOver <> 0 Then
                                NFace = VertexOver
                                For n = 1 To frmMain.sldExtend(0)
                                    
                                    Sye = (1 - (frmMain.sldExtend(1) / 60)) * (1 - (((n - (frmMain.sldExtend(0) / 2) - 0.5) / (frmMain.sldExtend(0) / 2) * frmMain.sldExtend(2))) / 10)
                                    NFace = Model.Geometery(ObjectOver).FragmentFace(NFace, 2, Sye)
                                    Select Case ViewMode
                                        Case 1: Model.Geometery(ObjectOver).MoveFace NFace, Snaped((X - gGuide.x1)) / (frmMain.sldExtend(0)), 0, Snaped((Y - gGuide.y1)) / (frmMain.sldExtend(0) + 1)
                                        Case 2: Model.Geometery(ObjectOver).MoveFace NFace, Snaped((X - gGuide.x1)) / (frmMain.sldExtend(0)), Snaped((Y - gGuide.y1)) / (frmMain.sldExtend(0) + 1), 0
                                        Case 3: Model.Geometery(ObjectOver).MoveFace NFace, 0, Snaped((Y - gGuide.y1)) / (frmMain.sldExtend(0)), Snaped((X - gGuide.x1)) / (frmMain.sldExtend(0) + 1)
                                    End Select
                                    Model.Geometery(ObjectOver).FindObjectOutline
                                Next n
                            End If
                        End If
                        Model.FindModelOutline
                        Refresh
                        Exit Sub
                    
                    Case 14
                        Do
                            ObjectOver = FaceObjectHitTest(gGuide.x1, gGuide.y1)
                            If ObjectOver <> 0 Then
                                Select Case ViewMode
                                    Case 1: Cx = Snaped(X - gGuide.x1): Cz = Snaped(Y - gGuide.y1)
                                    Case 2: Cx = Snaped(X - gGuide.x1): Cy = Snaped(Y - gGuide.y1)
                                    Case 3: Cz = Snaped(X - gGuide.x1): Cx = Snaped(Y - gGuide.y1)
                                End Select
                                VertexOver = FaceHitTest(Model.Geometery(ObjectOver), gGuide.x1, gGuide.y1)
                                If VertexOver <> 0 Then
                                    Model.Geometery(ObjectOver).MoveFace VertexOver, Int(Cx), Int(Cy), Int(Cz)
                                    Model.Geometery(ObjectOver).FindObjectOutline
                                End If
                            End If
                        Loop Until ObjectOver = 0 Or Shift = 1
                        Model.FindModelOutline
                        Refresh
                End Select
                Model.Saved = False
             
            ElseIf EditButton = 4 And ObjectScaleMode > 0 Then
                With Model
                    AbsX = AbsoluteX(X)
                    AbsY = AbsoluteY(Y)
                    Select Case ViewMode
                        Case 1
                            Select Case ObjectScaleMode
                                Case 1, 2, 3, 4: If .MaxX - .MinX = 0 Or .MaxZ - .MinZ = 0 Then ShowScaleError: Exit Sub
                                Case 5, 6: If .MaxX - .MinX = 0 Then ShowScaleError: Exit Sub
                                Case 7, 8: If .MaxZ - .MinZ = 0 Then ShowScaleError: Exit Sub
                            End Select
                            Select Case ObjectScaleMode
                                Case 1: sX = (.MaxX - AbsX) / (.MaxX - .MinX): sY = (.MaxZ - AbsY) / (.MaxZ - .MinZ): Cx = .MaxX: Cz = .MaxZ
                                Case 2: sX = (AbsX - .MinX) / (.MaxX - .MinX): sY = (.MaxZ - AbsY) / (.MaxZ - .MinZ): Cx = .MinX: Cz = .MaxZ
                                Case 3: sX = (AbsX - .MinX) / (.MaxX - .MinX): sY = (AbsY - .MinZ) / (.MaxZ - .MinZ): Cx = .MinX: Cz = .MinZ
                                Case 4: sX = (.MaxX - AbsX) / (.MaxX - .MinX):  sY = (AbsY - .MinZ) / (.MaxZ - .MinZ): Cx = .MaxX: Cz = .MinZ
                                Case 5:  sX = (.MaxX - AbsX) / (.MaxX - .MinX): sY = 1: Cy = 0: Cx = .MaxX
                                Case 6:  sX = (AbsX - .MinX) / (.MaxX - .MinX): sY = 1: Cy = 0: Cx = .MinX
                                Case 7:  sX = 1: sY = (.MaxZ - AbsoluteY(Y)) / (.MaxZ - .MinZ): Cz = .MaxZ: Cx = 0
                                Case 8:  sX = 1: sY = (AbsoluteY(Y) - .MinZ) / (.MaxZ - .MinZ): Cz = .MinZ: Cx = 0
                            End Select
    
                        Case 2
                            Select Case ObjectScaleMode
                                Case 1, 2, 3, 4: If .MaxX - .MinX = 0 Or .MaxY - .MinY = 0 Then ShowScaleError: Exit Sub
                                Case 5, 6: If .MaxX - .MinX = 0 Then ShowScaleError: Exit Sub
                                Case 7, 8: If .MaxY - .MinY = 0 Then ShowScaleError: Exit Sub
                            End Select
                            Select Case ObjectScaleMode
                                Case 1: sX = (.MaxX - AbsX) / (.MaxX - .MinX): sY = (.MaxY - AbsY) / (.MaxY - .MinY): Cx = .MaxX: Cy = .MaxY
                                Case 2: sX = (AbsX - .MinX) / (.MaxX - .MinX): sY = (.MaxY - AbsY) / (.MaxY - .MinY): Cx = .MinX: Cy = .MaxY
                                Case 3: sX = (AbsX - .MinX) / (.MaxX - .MinX): sY = (AbsY - .MinY) / (.MaxY - .MinY): Cx = .MinX: Cy = .MinY
                                Case 4: sX = (.MaxX - AbsX) / (.MaxX - .MinX):  sY = (AbsY - .MinY) / (.MaxY - .MinY): Cx = .MaxX: Cy = .MinY
                                Case 5:  sX = (.MaxX - AbsX) / (.MaxX - .MinX): sY = 1: Cy = 0: Cx = .MaxX
                                Case 6:  sX = (AbsX - .MinX) / (.MaxX - .MinX): sY = 1: Cy = 0: Cx = .MinX
                                Case 7:  sX = 1: sY = (.MaxY - AbsoluteY(Y)) / (.MaxY - .MinY): Cy = .MaxY: Cx = 0
                                Case 8:  sX = 1: sY = (AbsoluteY(Y) - .MinY) / (.MaxY - .MinY): Cy = .MinY: Cx = 0
                            End Select
    
                        Case 3
                            Select Case ObjectScaleMode
                                Case 1, 2, 3, 4: If .MaxY - .MinY = 0 Or .MaxZ - .MinZ = 0 Then ShowScaleError: Exit Sub
                                Case 5, 6: If .MaxZ - .MinZ = 0 Then ShowScaleError: Exit Sub
                                Case 7, 8: If .MaxY - .MinY = 0 Then ShowScaleError: Exit Sub
                            End Select
                            Select Case ObjectScaleMode
                                Case 1: sX = (.MaxZ - AbsX) / (.MaxZ - .MinZ): sY = (.MaxY - AbsY) / (.MaxY - .MinY): Cz = .MaxZ: Cy = .MaxY
                                Case 2: sX = (AbsX - .MinZ) / (.MaxZ - .MinZ): sY = (.MaxY - AbsY) / (.MaxY - .MinY): Cz = .MinZ: Cy = .MaxY
                                Case 3: sX = (AbsX - .MinZ) / (.MaxZ - .MinZ): sY = (AbsY - .MinY) / (.MaxY - .MinY): Cz = .MinZ: Cy = .MinY
                                Case 4: sX = (.MaxZ - AbsX) / (.MaxZ - .MinZ):  sY = (AbsY - .MinY) / (.MaxY - .MinY): Cz = .MaxZ: Cy = .MinY
                                Case 5:  sX = (.MaxZ - AbsX) / (.MaxZ - .MinZ): sY = 1: Cy = 0: Cz = .MaxZ
                                Case 6:  sX = (AbsX - .MinZ) / (.MaxZ - .MinZ): sY = 1: Cy = 0: Cz = .MinZ
                                Case 7:  sX = 1: sY = (.MaxY - AbsoluteY(Y)) / (.MaxY - .MinY): Cy = .MaxY: Cx = 0
                                Case 8:  sX = 1: sY = (AbsoluteY(Y) - .MinY) / (.MaxY - .MinY): Cy = .MinY: Cx = 0
                            End Select
                    End Select
                    
                    For Each Am In .Geometery
                        If Am.Selected = True Then
                            If sX < 0 Xor sY < 0 Then Am.ReverseFace
                            Select Case ViewMode
                                Case 1: Am.Grow sX, 1, sY, Cx, Cy, Cz
                                Case 2: Am.Grow sX, sY, 1, Cx, Cy, Cz
                                Case 3: Am.Grow 1, sY, sX, Cx, Cy, Cz
                            End Select
                            Am.FindObjectOutline
                        End If
                    Next Am
                    
                    For Each Jm In .Joint
                        If Jm.Selected = True Then
                            Select Case ViewMode
                                Case 1: Jm.Grow sX, 1, sY, Cx, Cy, Cz
                                Case 2: Jm.Grow sX, sY, 1, Cx, Cy, Cz
                                Case 3: Jm.Grow 1, sY, sX, Cx, Cy, Cz
                            End Select
                        End If
                    Next Jm
                    
                    Model.Saved = False
                    .FindModelOutline
                    Refresh
                End With
                Model.Saved = False
            
            ElseIf EditButton = 6 And frmMain.cmdSidebar.SelectedItem.Index = 1 And frmMain.opAddJ Then
                jKey = "Joint" & Timer & Rnd * 10
                If SelectedNode(frmMain.Joints) > 1 Then tkey = frmMain.Joints.Nodes(SelectedNode(frmMain.Joints)).Key
                Model.Joint.AddJoint jKey, tkey
                Model.Joint(jKey).Name = "New Joint " & Model.Joint.CountChildren
                Model.Joint.DisplayTreeInWindow frmMain.Joints
                With Model.Joint(jKey)
                    Select Case ViewMode
                        Case 1: .X = AbsoluteX(X): .Y = 0: .z = AbsoluteY(Y)
                        Case 2: .X = AbsoluteX(X): .z = 0: .Y = AbsoluteY(Y)
                        Case 3: .z = AbsoluteX(X): .X = 0: .Y = AbsoluteY(Y)
                    End Select
                    Model.Saved = False
                    Model.Scene.UpdateAllScenes
                    'Model.Joint.DeselectAll
                    If frmMain.chkSelect(6) = 0 Then
                        Model.Joint(jKey).Selected = True
                    End If
                    frmMain.Joints.Nodes(jKey).Selected = True
                    Model.FindModelOutline
                End With
                Refresh
            ElseIf EditButton = 6 And frmMain.cmdSidebar.SelectedItem.Index = 1 And frmMain.opChangeJ = True Then
                     
                Dim StartJoint As Integer, EndJoint As Integer
                StartJoint = JointHitTest(StartX, StartY)
                EndJoint = JointHitTest(X, Y)
                If StartJoint <> 0 Then
                    If EndJoint = 0 Then
                            Model.Joint(StartJoint).Target = ""
                        Else
                            Model.Joint(StartJoint).Target = Model.Joint(EndJoint).Key
                    End If
                        
                    Model.Joint.DisplayTreeInWindow frmMain.Joints
                    Refresh
                End If
                gGuide.Visible = False
                Model.Saved = False
    
            Else
                If InsideBoundingBox(StartX, StartY) = True And ObjectRotateMode = False Then
                    If frmMain.chkSelect(6) = 0 Or (frmMain.chkSelect(6) = 1 And Shift = 1) Then
                        Model.Geometery.MoveSelected Snaped(X - StartX) / ZoomLevel, Snaped(Y - StartY) / ZoomLevel, ViewMode
                        Model.Joint.MoveSelected Snaped(X - StartX) / ZoomLevel, Snaped(Y - StartY) / ZoomLevel, ViewMode
                        Model.FindModelOutline
                    End If
                End If
                If ObjectRotateMode = True Then
                    Model.RotateSelection GetAngle(Snaped(StartX - X), Snaped(Y - StartY)) * 0.1, ViewMode, 0
                End If
                If StartBoxSelect = True Then
                    BoxBandSelect X, Y, StartX, StartY, frmMain.chkSelect(3)
                    StartBoxSelect = False
                End If
                Model.FindModelOutline
                ObjectRotateMode = False
                Refresh
            End If
        Else
            JointOver = JointHitTest(X, Y)
            If JointOver <> "" Then
                If frmMain.chkSelect(6) = 1 Then
                    frmMain.ActiveForm.mnuEditPopup(11).Caption = "Attach Vertecies to " & Model.Joint(JointOver).Name
                Else
                    frmMain.ActiveForm.mnuEditPopup(11).Caption = "Attach Objects to " & Model.Joint(JointOver).Name
                End If
                frmMain.ActiveForm.mnuRightDrag(3).Caption = "Attach to " & Model.Joint(JointOver).Name
            Else
                frmMain.ActiveForm.mnuEditPopup(11).Caption = "No joint found"
                frmMain.ActiveForm.mnuRightDrag(3).Caption = "No joint found"
            End If
            If Almost(X, StartX, Y, StartY) Then
                PopupMenu frmMain.ActiveForm.menuEditPopup
                Else
                If Model.Geometery.CountSelected + Model.Joint.CountSelected <> 0 Then PopupMenu frmMain.ActiveForm.menuRightDrag
            End If
        End If
    End If
    'RaiseEvent SendMouseUp(X, Y, Button, Shift)
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    Xf = ScaleWidth / 2
    Yf = ScaleHeight / 2
    sbBar(1).Left = ScaleWidth - sbBar(1).Width
    sbBar(1).Height = ScaleHeight - sbBar(1).Width
    sbBar(0).Top = ScaleHeight - sbBar(0).Height
    sbBar(0).Width = ScaleWidth - sbBar(0).Height
    cmdBlock.Left = sbBar(0).Width
    cmdBlock.Top = sbBar(1).Height
End Sub

Public Function MoveTo(x1 As Integer, y1 As Integer)
    'This function uses the MoveTo command on its own, to set the position of the
    'cursor on the screen. By using the DrawTo command, you can draw from whereever
    'the cursor was to a new position, which is faster that seting the cursor position
    'every time
    MoveToEx hdc, (x1 - sbBar(0)) * ZoomLevel + Xf, (y1 + -sbBar(1)) * ZoomLevel + Yf, API
End Function

Public Function DrawTo(x1 As Integer, y1 As Integer)
    'This uses the LineTo command to draw from where ever the cursor was before, to
    'a new position, which is faster than having to set the cursor every time.
    LineTo hdc, (x1 - sbBar(0)) * ZoomLevel + Xf, (y1 + -sbBar(1)) * ZoomLevel + Yf
End Function

Public Function DrawBox(x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer, Optional Mode As Integer = 0)
    'This draws a square box on the screen, taking into acount zoom levels and the scroll bars
    MoveToEx hdc, (x1 - sbBar(0)) * ZoomLevel + Xf, (y1 + -sbBar(1)) * ZoomLevel + Yf, API
    LineTo hdc, (x2 - sbBar(0)) * ZoomLevel + Xf, (y1 + -sbBar(1)) * ZoomLevel + Yf
    LineTo hdc, (x2 - sbBar(0)) * ZoomLevel + Xf, (y2 + -sbBar(1)) * ZoomLevel + Yf
    LineTo hdc, (x1 - sbBar(0)) * ZoomLevel + Xf, (y2 + -sbBar(1)) * ZoomLevel + Yf
    LineTo hdc, (x1 - sbBar(0)) * ZoomLevel + Xf, (y1 + -sbBar(1)) * ZoomLevel + Yf
    If Mode = 1 Then
        DrawCircle x1, y1, 3, 1:       DrawCircle (x1 + x2) * 0.5, y1, 3, 1
        DrawCircle x2, y1, 3, 1:       DrawCircle (x1 + x2) * 0.5, y2, 3, 1
        DrawCircle x1, y2, 3, 1:       DrawCircle x1, (y1 + y2) * 0.5, 3, 1
        DrawCircle x2, y2, 3, 1:       DrawCircle x2, (y1 + y2) * 0.5, 3, 1
        DrawCircle (x1 + x2) * 0.5, (y1 + y2) * 0.5, 3, 1
    End If
End Function

Public Sub DrawStaticLine(x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer)
    'This draws a line onto the screen, taking into acount the scroll bars and
    'zoom factor. This will be called quite abit
    MoveToEx hdc, (x1), (y1), API: LineTo hdc, (x2), (y2)
End Sub

Public Sub DrawLine(x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer)
    'This draws a line onto the screen, taking into acount the scroll bars and
    'zoom factor. This will be called quite abit
    MoveToEx hdc, (x1 - sbBar(0)) * ZoomLevel + Xf, (y1 + -sbBar(1)) * ZoomLevel + Yf, API
    LineTo hdc, (x2 - sbBar(0)) * ZoomLevel + Xf, (y2 + -sbBar(1)) * ZoomLevel + Yf
End Sub

Public Sub DrawCircle(X As Integer, Y As Integer, iWidth As Integer, Optional Mode As Integer = 0)
    'This draws a circle onto the screen, taking into acount the zoom and scroll bars
    If Mode = 1 Then
        Ellipse hdc, (X - (iWidth / ZoomLevel) - sbBar(0)) * ZoomLevel + Xf, (Y - (iWidth / ZoomLevel) + -sbBar(1)) * ZoomLevel + ScaleHeight / 2, (X + (iWidth / ZoomLevel) - sbBar(0)) * ZoomLevel + Xf, (Y + (iWidth / ZoomLevel) + -sbBar(1)) * ZoomLevel + Yf
    Else
        Ellipse hdc, (X - iWidth - sbBar(0)) * ZoomLevel + Xf, (Y - iWidth + -sbBar(1)) * ZoomLevel + Yf, (X + iWidth - sbBar(0)) * ZoomLevel + Xf, (Y + iWidth + -sbBar(1)) * ZoomLevel + Yf
    End If
End Sub

Public Sub Refresh(Optional UpdateOthers As Boolean = True)
    If Not Model Is Nothing Then
        UserControl_Paint
        If UpdateOthers = True Then RaiseEvent UpdateOtherWindows
    End If
End Sub

Private Sub UserControl_Paint()
    ClearTablet
    DrawObject2D
End Sub

Public Sub ClearTablet()
    'This clears the window and draws the grid
    Dim n As Integer, hdc As Long
    Dim OfsX As Integer, OfsY As Integer
    Dim SmallGrid As Byte, LargeGrid As Byte
    Cls
    If pShowGrid = True Then
        LargeGrid = 50: SmallGrid = 10
        ForeColor = RGB(200, 200, 255): DrawStyle = 0
        OfsX = ((ScaleWidth / 2) - sbBar(0) * ZoomLevel) Mod LargeGrid * ZoomLevel
        OfsY = ((ScaleHeight / 2) - sbBar(1) * ZoomLevel) Mod LargeGrid * ZoomLevel
        For n = OfsX To ScaleWidth + OfsX Step LargeGrid * ZoomLevel: DrawStaticLine n, 0, n, ScaleHeight: Next n
        For n = OfsY To ScaleHeight + OfsY Step LargeGrid * ZoomLevel: DrawStaticLine 0, n, ScaleWidth, n: Next n
        ForeColor = 0
        n = (ScaleWidth / 2) - sbBar(0) * ZoomLevel: DrawStaticLine n, 0, n, ScaleHeight
        n = (ScaleHeight / 2) - sbBar(1) * ZoomLevel: DrawStaticLine 0, n, ScaleWidth, n
    End If
End Sub

Private Function DrawObject2D() As Boolean
    'This draws the object in 2D onto the tablet given in the parameter line. There are
    'several optional parameters to alter how it is drawn. ShowVerteceis highlights each veretx,
    'showface highlights each face, and ObjectSelected highlights the entire object.
    Dim Coord(60, 2) As Integer, n As Byte, m As Byte, RemoveHiddenFace As Boolean
    Dim FaceON As clsFace, VertexOn As clsEdge, CenX As Single, CenY As Single
    Dim Am As clsObject, Pm As clsJoint, Vm As clsVertex
    For Each Am In Model.Geometery
        
        If Model.Layers.Layer(Am.Layer).Selected = True And Am.Hidden = False Then
            If pShowVertecies And Am.Selected Then
                For Each Vm In Am.Vertex
                    Select Case ViewMode
                        Case 1
                            If Vm.Selected = True Then DrawWidth = 2
                            DrawCircle Vm.X, Vm.z, 4, 1
                            DrawWidth = 1
                        Case 2
                            If Vm.Selected = True Then DrawWidth = 2
                            DrawCircle Vm.X, Vm.Y, 4, 1
                            DrawWidth = 1
                        Case 3
                            If Vm.Selected = True Then DrawWidth = 2
                            DrawCircle Vm.z, Vm.Y, 4, 1
                            DrawWidth = 1
                    End Select
                Next Vm
            End If
        
            For Each FaceON In Am.Face
                n = 0: CenX = 0: CenY = 0
                ForeColor = Am.Colour
                For Each VertexOn In FaceON.Edge
                    n = n + 1
                    Select Case ViewMode
                        Case 1: Coord(n, 1) = Am.Vertex(VertexOn.Vertex).X: Coord(n, 2) = Am.Vertex(VertexOn.Vertex).z
                        Case 2: Coord(n, 1) = Am.Vertex(VertexOn.Vertex).X: Coord(n, 2) = Am.Vertex(VertexOn.Vertex).Y
                        Case 3: Coord(n, 1) = Am.Vertex(VertexOn.Vertex).z: Coord(n, 2) = Am.Vertex(VertexOn.Vertex).Y
                    End Select
                Next VertexOn
                MoveTo Coord(n, 1), Coord(n, 2)
                For m = 1 To n
                    DrawTo Coord(m, 1), Coord(m, 2)
                    CenX = CenX + Coord(m, 1)
                    CenY = CenY + Coord(m, 2)
                Next m
                If pShowFaces = True And Am.Selected = True Then
                    CenX = CenX / n
                    CenY = CenY / n
                    DrawCircle Int(CenX), Int(CenY), 4
                End If
            Next FaceON
            If Am.Selected = True Then
                ForeColor = vbRed
                Select Case ViewMode
                    Case 1:  DrawBox Am.MinX, Am.MinZ, Am.MaxX, Am.MaxZ
                    Case 2:  DrawBox Am.MinX, Am.MinY, Am.MaxX, Am.MaxY
                    Case 3:  DrawBox Am.MinZ, Am.MinY, Am.MaxZ, Am.MaxY
                End Select
            End If
        End If
    Next Am
    
    For Each Pm In Model.Joint
        If Pm.Hidden = False Then
            ForeColor = Pm.Colour
           ' If Pm.Grayed = True Then ForeColor = frmSettings.ColourBox(7).BackColor
            Select Case ViewMode
                Case 1
                    If Pm.Selected = True Then DrawWidth = 2
                    DrawCircle Pm.X, Pm.z, 4
                    DrawWidth = 1
                    If Pm.Target <> "" Then DrawLine Pm.X, Pm.z, Model.Joint(Pm.Target).X, Model.Joint(Pm.Target).z
    
                Case 2
                    If Pm.Selected = True Then DrawWidth = 2
                    DrawCircle Pm.X, Pm.Y, 4
                    DrawWidth = 1
                    If Pm.Target <> "" Then DrawLine Pm.X, Pm.Y, Model.Joint(Pm.Target).X, Model.Joint(Pm.Target).Y
    
                Case 3
                    If Pm.Selected = True Then DrawWidth = 2
                    DrawCircle Pm.z, Pm.Y, 4
                    DrawWidth = 1
                    If Pm.Target <> "" Then DrawLine Pm.z, Pm.Y, Model.Joint(Pm.Target).z, Model.Joint(Pm.Target).Y
            
            End Select
        End If
    Next Pm
    
    With Model
        If .Geometery.FirstSelectedObject + .Joint.FirstSelectedJoint <> 0 Then
            ForeColor = vbBlue
            Select Case ViewMode
                Case 1: DrawBox .MinX, .MinZ, .MaxX, .MaxZ, IntBo(Am8.HighLightSection)
                Case 2: DrawBox .MinX, .MinY, .MaxX, .MaxY, IntBo(Am8.HighLightSection)
                Case 3: DrawBox .MinZ, .MinY, .MaxZ, .MaxY, IntBo(Am8.HighLightSection)
            End Select
        End If
    End With
    
End Function


Public Function InsideBoundingBox(X As Single, Y As Single) As Boolean
    Dim Xx1 As Single, Yy1 As Single
    Dim Xx2 As Single, Yy2 As Single
    InsideBoundingBox = False
    If Model.Geometery.CountSelected + Model.Joint.CountChildren = 0 Then Exit Function
    With Model
        Select Case ViewMode
            Case 1
                Xx1 = ((.MinX - sbBar(0)) * ZoomLevel + ScaleWidth * 0.5) - 5
                Yy1 = ((.MinZ - sbBar(1)) * ZoomLevel + ScaleHeight * 0.5) - 5
                Xx2 = ((.MaxX - sbBar(0)) * ZoomLevel + ScaleWidth * 0.5) + 5
                Yy2 = ((.MaxZ - sbBar(1)) * ZoomLevel + ScaleHeight * 0.5) + 5
        
            Case 2
                Xx1 = ((.MinX - sbBar(0)) * ZoomLevel + ScaleWidth * 0.5) - 5
                Yy1 = ((.MinY - sbBar(1)) * ZoomLevel + ScaleHeight * 0.5) - 5
                Xx2 = ((.MaxX - sbBar(0)) * ZoomLevel + ScaleWidth * 0.5) + 5
                Yy2 = ((.MaxY - sbBar(1)) * ZoomLevel + ScaleHeight * 0.5) + 5
        
            Case 3
                Xx1 = ((.MinZ - sbBar(0)) * ZoomLevel + ScaleWidth * 0.5) - 5
                Yy1 = ((.MinY - sbBar(1)) * ZoomLevel + ScaleHeight * 0.5) - 5
                Xx2 = ((.MaxZ - sbBar(0)) * ZoomLevel + ScaleWidth * 0.5) + 5
                Yy2 = ((.MaxY - sbBar(1)) * ZoomLevel + ScaleHeight * 0.5) + 5
        
        End Select
    End With
    If X >= Xx1 And X <= Xx2 And Y >= Yy1 And Y <= Yy2 Then InsideBoundingBox = True
End Function

Public Function AbsoluteX(Value As Single) As Single
    AbsoluteX = Snaped(((Value - Xf) / ZoomLevel) + sbBar(0))
End Function

Public Function AbsoluteY(Value As Single) As Single
    AbsoluteY = Snaped(((Value - Yf) / ZoomLevel) + sbBar(1))
End Function

Public Function ObjectHitTest(ByVal X As Single, ByVal Y As Single, Optional SelectedOnly As Boolean = False) As Integer
    'This checks all objects to see if you have clicked on one. It returns
    'the number of the object that has a vertex under the mouse
    Dim Counter As Integer, Am As clsObject, Vertex As clsVertex, XX As Single, YY As Single
    X = ((X - Xf) / ZoomLevel) + sbBar(0)
    Y = ((Y - Yf) / ZoomLevel) + sbBar(1)
    For Each Am In Model.Geometery
        Counter = Counter + 1
        If Am.Hidden = False And Am.Locked = False And Model.Layers(Am.Layer).Selected = True And Model.Layers(Am.Layer).LayerLocked = False And (SelectedOnly = False Or Am.Selected = True) Then
            Select Case ViewMode
                Case 1
                    If X > Am.MinX - 4 And X < Am.MaxX + 4 And Y > Am.MinZ - 4 And Y < Am.MaxZ + 4 Then
                        For Each Vertex In Am.Vertex
                            If Almost(Int(Vertex.X), X, Int(Vertex.z), Y) = True Then ObjectHitTest = Counter: Exit Function
                        Next Vertex
                    End If
                        
                Case 2
                    If X > Am.MinX - 4 And X < Am.MaxX + 4 And Y > Am.MinY - 4 And Y < Am.MaxY + 4 Then
                        For Each Vertex In Am.Vertex
                            If Almost(Int(Vertex.X), X, Int(Vertex.Y), Y) = True Then ObjectHitTest = Counter: Exit Function
                        Next Vertex
                    End If
                
                Case 3
                    If X > Am.MinZ - 4 And X < Am.MaxZ + 4 And Y > Am.MinY - 4 And Y < Am.MaxY + 4 Then
                        For Each Vertex In Am.Vertex
                            If Almost(Int(Vertex.Y), Y, Int(Vertex.z), X) = True Then ObjectHitTest = Counter: Exit Function
                        Next Vertex
                    End If
            
            End Select
        End If
    Next Am
End Function

Public Function FaceObjectHitTest(ByVal X As Single, ByVal Y As Single) As Integer
    Dim Counter As Integer, Am As clsObject, VertexOn As clsEdge
    Dim FaceON As clsFace, CenX As Single, CenY As Single
    X = AbsoluteX(X)
    Y = AbsoluteY(Y)
    For Each Am In Model.Geometery
        Counter = Counter + 1
        For Each FaceON In Am.Face
            CenX = 0: CenY = 0
            For Each VertexOn In FaceON.Edge
                Select Case ViewMode
                    Case 1
                        CenX = CenX + Am.Vertex(VertexOn.Vertex).X
                        CenY = CenY + Am.Vertex(VertexOn.Vertex).z
                    Case 2
                        CenX = CenX + Am.Vertex(VertexOn.Vertex).X
                        CenY = CenY + Am.Vertex(VertexOn.Vertex).Y
                    Case 3
                        CenX = CenX + Am.Vertex(VertexOn.Vertex).z
                        CenY = CenY + Am.Vertex(VertexOn.Vertex).Y
                End Select
            Next VertexOn
            CenX = CenX / FaceON.EdgeCount
            CenY = CenY / FaceON.EdgeCount
            If Almost(Int(CenX), X, Int(CenY), Y) = True Then FaceObjectHitTest = Counter
        Next FaceON
    Next Am
End Function

Public Function FaceHitTest(iShape As clsObject, ByVal X As Single, ByVal Y As Single) As Integer
    Dim Counter As Integer, CenX As Single, CenY As Single
    Dim FaceON As clsFace, VertexOn As clsEdge
    X = AbsoluteX(X): Y = AbsoluteY(Y)
    For Each FaceON In iShape.Face
        Counter = Counter + 1
        CenX = 0: CenY = 0
        For Each VertexOn In FaceON.Edge
            Select Case ViewMode
                Case 1
                    CenX = CenX + iShape.Vertex(VertexOn.Vertex).X
                    CenY = CenY + iShape.Vertex(VertexOn.Vertex).z
                Case 2
                    CenX = CenX + iShape.Vertex(VertexOn.Vertex).X
                    CenY = CenY + iShape.Vertex(VertexOn.Vertex).Y
                Case 3
                    CenX = CenX + iShape.Vertex(VertexOn.Vertex).z
                    CenY = CenY + iShape.Vertex(VertexOn.Vertex).Y
            End Select
        Next VertexOn
        CenX = CenX / FaceON.EdgeCount
        CenY = CenY / FaceON.EdgeCount
        If Almost(Int(CenX), X, Int(CenY), Y) = True Then FaceHitTest = Counter
    Next FaceON
End Function

Public Function VertexHitTest(iShape As clsObject, X As Single, Y As Single) As Integer
    'This returns the vertex that the mouse is over. Again, you have to
    'tell it which object to varify.
    Dim Counter As Integer, Vertex As clsVertex, XX As Single, YY As Single
    For Each Vertex In iShape.Vertex
        Counter = Counter + 1
        If ViewMode = 1 Then
            XX = (Vertex.X - sbBar(0)) * ZoomLevel + Xf
            YY = (Vertex.z + -sbBar(1)) * ZoomLevel + Yf
        ElseIf ViewMode = 2 Then
            XX = (Vertex.X - sbBar(0)) * ZoomLevel + Xf
            YY = (Vertex.Y + -sbBar(1)) * ZoomLevel + Yf
        Else
            XX = (Vertex.z - sbBar(0)) * ZoomLevel + Xf
            YY = (Vertex.Y + -sbBar(1)) * ZoomLevel + Yf
        End If
        If Almost(Int(XX), X, Int(YY), Y) = True Then VertexHitTest = Counter: Exit Function
    Next Vertex
End Function

Public Function Snaped(Value) As Single
    'This snaps a value to the nearest grid line, looking at the value in
    'the editor settings window.
    If ViewMode > 3 Or frmMain.ActiveForm.mnuView(1).Checked = False Then
        Snaped = Value
    Else
        Snaped = ((CInt((Value) / Am8.SnapSize)) * Am8.SnapSize)
    End If
End Function

Public Sub DrawShapeGuide(ShapeClass As String, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single)
    Dim X As Integer, Y As Integer, NextRow As Integer, MidCirc As Integer
    Dim Xw As Integer, Yw As Integer, ReduceMe1 As Single
    Dim Ang As Integer, n As Integer, ReduceMe2 As Single
    Dim x3 As Integer, y3 As Integer, x4 As Integer, y4 As Integer
    Dim nX1 As Integer, nX2 As Integer, ny1 As Integer, ny2 As Integer
    x1 = AbsoluteX(x1):         y1 = AbsoluteY(y1)
    x2 = AbsoluteX(x2):         y2 = AbsoluteY(y2)
    X = (x1 + x2) * 0.5:        Y = (y1 + y2) * 0.5
    Xw = (x2 - x1) * 0.5:       Yw = (y2 - y1) * 0.5
    AutoRedraw = False
    Cls
    DrawStyle = 2
    Select Case ShapeClass

        Case "Cube", "Grid", "Rubix Cube"
            DrawBox Int(x1), Int(y1), Int(x2), Int(y2)

        Case "Torous"
            Ang = 180 / frmMain.ShpProp(1).Value + (frmMain.ShpProp(2).Value * 5)
            For n = 0 To 359 Step 360 / frmMain.ShpProp(1).Value
                NextRow = 360 / frmMain.ShpProp(1).Value
                x1 = Int(Sin((n + Ang) / Pie) * Xw) + X
                y1 = Int(Cos((n + Ang) / Pie) * Yw) + Y
                x2 = Int(Sin((n + Ang + NextRow) / Pie) * Xw) + X
                y2 = Int(Cos((n + Ang + NextRow) / Pie) * Yw) + Y
              '  If frmMain.Axis.Tag = "Y" Then
                    nX1 = x1 - ((x1 - 0) * 2)
                    nX2 = x2 - ((x2 - 0) * 2)
                    ny1 = y1
                    ny2 = y2
              '  ElseIf frmMain.Axis.Tag = "X" Then
               '     nX1 = x1
                '    nX2 = x2
                '    ny1 = y1 - ((y1 - frmMain.Axis.y1) * 2)
                 '   ny2 = y2 - ((y2 - frmMain.Axis.y1) * 2)
               ' End If
                DrawLine nX1, ny1, nX2, ny2
                DrawLine Int(x1), Int(y1), Int(nX1), Int(ny1)
                DrawLine Int(x1), Int(y1), Int(x2), Int(y2)
            Next n
            'frmMain.View.Line (frmMain.Guide.x1, frmMain.Guide.y1)-(frmMain.Guide.x2, frmMain.Guide.y2), Colours(3), B


        Case "Star"
            DrawBox Int(x1), Int(y1), Int(x2), Int(y2)
            Ang = 180 / (frmMain.ShpProp(1).Value * 2) + (frmMain.ShpProp(2).Value * 5)
            For n = 0 To 359 Step 360 / (frmMain.ShpProp(1).Value * 2)
                NextRow = 360 / (frmMain.ShpProp(1).Value * 2)
                If MidCirc = 0 Then
                    ReduceMe1 = 100 / (frmMain.ShpProp(3) * 5): ReduceMe2 = 100 / (frmMain.ShpProp(4) * 5): MidCirc = 1
                Else
                    ReduceMe1 = 100 / (frmMain.ShpProp(4) * 5): ReduceMe2 = 100 / (frmMain.ShpProp(3) * 5): MidCirc = 0
                End If
                x1 = (Sin((n + Ang) / Pie) * Xw) / ReduceMe1
                y1 = (Cos((n + Ang) / Pie) * Yw) / ReduceMe1
                x2 = (Sin((n + Ang + NextRow) / Pie) * Xw) / ReduceMe2
                y2 = (Cos((n + Ang + NextRow) / Pie) * Yw) / ReduceMe2
                DrawLine x1 + X, y1 + Y, x2 + X, y2 + Y
            Next n

        Case "Prism"
            DrawBox Int(x1), Int(y1), Int(x2), Int(y2)
            Ang = 180 / frmMain.ShpProp(1).Value + (frmMain.ShpProp(2).Value * 5)
            For n = 0 To 359 Step 360 / frmMain.ShpProp(1).Value
                NextRow = 360 / frmMain.ShpProp(1).Value
                ReduceMe1 = 100 / (frmMain.ShpProp(3) * 5)
                ReduceMe2 = 100 / (frmMain.ShpProp(4) * 5)
                x1 = (Sin((n + Ang) / Pie) * Xw) / ReduceMe1
                y1 = (Cos((n + Ang) / Pie) * Yw) / ReduceMe1
                x2 = (Sin((n + Ang + NextRow) / Pie) * Xw) / ReduceMe1
                y2 = (Cos((n + Ang + NextRow) / Pie) * Yw) / ReduceMe1
                x3 = (Sin((n + Ang) / Pie) * Xw) / ReduceMe2
                y3 = (Cos((n + Ang) / Pie) * Yw) / ReduceMe2
                x4 = (Sin((n + Ang + NextRow) / Pie) * Xw) / ReduceMe2
                y4 = (Cos((n + Ang + NextRow) / Pie) * Yw) / ReduceMe2
                DrawLine x1 + X, y1 + Y, x2 + X, y2 + Y
                DrawLine x1 + X, y1 + Y, x3 + X, y3 + Y
                DrawLine x3 + X, y3 + Y, x4 + X, y4 + Y
            Next n
        
        Case "Dimond", "Cone", "Face"
            DrawBox Int(x1), Int(y1), Int(x2), Int(y2)
            Ang = 180 / frmMain.ShpProp(1).Value + (frmMain.ShpProp(2).Value * 5)
            For n = 0 To 359 Step 360 / frmMain.ShpProp(1).Value
                NextRow = 360 / frmMain.ShpProp(1).Value
                x1 = (Sin((n + Ang) / Pie) * Xw)
                y1 = (Cos((n + Ang) / Pie) * Yw)
                x2 = (Sin((n + Ang + NextRow) / Pie) * Xw)
                y2 = (Cos((n + Ang + NextRow) / Pie) * Yw)
                DrawLine x1 + X, y1 + Y, x2 + X, y2 + Y
                If ShapeClass = "Cone" Or ShapeClass = "Dimond" Then
                    DrawLine x1 + X, y1 + Y, X, Y
                End If
            Next n

        Case "Wrap"
            If Model.Wraper.Count > 0 Then
                For n = 1 To Model.Wraper.Count - 1
                    x1 = Model.Wraper(n).X
                    y1 = Model.Wraper(n).Y
                    x2 = Model.Wraper(n + 1).X
                    y2 = Model.Wraper(n + 1).Y
                    nX1 = x1 - ((x1 - 0) * 2)
                    nX2 = x2 - ((x2 - 0) * 2)
                    ny1 = y1 '- ((y1 - 0) * 2)
                    ny2 = y2 '- ((y2 - 0) * 2)
                    DrawLine Int(x1), Int(y1), Int(x2), Int(y2)
                    DrawLine Int(nX1), Int(ny1), Int(nX2), Int(ny2)
                    DrawLine Int(nX1), Int(ny1), Int(x1), Int(y1)
                Next n
                If Model.Wraper.Count > 1 Then DrawLine Int(nX2), Int(ny2), Int(x2), Int(y2)
            End If

        Case Else
            Ang = -(180 / frmMain.ShpProp(1).Value + (frmMain.ShpProp(2).Value * 5))
            NextRow = 180 / frmMain.ShpProp(1).Value
            For n = NextRow To 180 Step 180 / frmMain.ShpProp(1).Value
                x1 = (Sin((n + Ang) / Pie) * Xw)
                y1 = (Cos((n + Ang) / Pie) * Yw)
                x2 = (Sin((n + Ang + NextRow) / Pie) * Xw)
                y2 = (Cos((n + Ang + NextRow) / Pie) * Yw)
                If n = NextRow Then
                    MidCirc = (Sin((NextRow + Ang) / Pie) * Xw)
                End If
                nX1 = x1 - ((x1 - MidCirc) * 2)
                ny1 = y1
                nX2 = x2 - ((x2 - MidCirc) * 2)
                ny2 = y2
                DrawLine nX1 + X, ny1 + Y, nX2 + X, ny2 + Y
                DrawLine nX1 + X, ny1 + Y, x1 + X, y1 + Y
                DrawLine x1 + X, y1 + Y, x2 + X, y2 + Y
            Next n

    End Select
End Sub

Public Function GetAngle(X, Y) As Single
    'This returns the angle from 0,0 to X,Y
    Dim An As Single
    If Y = 0 Then
        If X > 0 Then An = 90
        If X < 0 Then An = 270
    Else
        An = Atn(X / Y)
        An = An * Pie
        If X = 0 And Y < 0 Then An = 180
    End If
    If Y < 0 And X > 0 Then An = 180 - (Abs(An))
    If Y < 0 And X < 0 Then An = An + 180
    If Y > 0 And X < 0 Then An = 360 + An
    GetAngle = Int(An * 10)
End Function

Public Sub DrawRotateGuide(ByVal Angle As Single, ByVal X As Integer, ByVal Y As Integer, Mode As Byte)
    Dim Corner(4) As New clsVertex, n As Integer
    Dim CenterX As Integer, CenterY As Integer
    Angle = Angle Mod 3600: Cls
    With Model
        If Mode = 0 Or Mode = 1 Then
            If ViewMode = 1 Then CenterX = ((.MinX + .MaxX) / 2): CenterY = ((.MinZ + .MaxZ) / 2)
            If ViewMode = 2 Then CenterX = ((.MinX + .MaxX) / 2): CenterY = ((.MinY + .MaxY) / 2)
            If ViewMode = 3 Then CenterX = ((.MinY + .MaxY) / 2): CenterY = ((.MinZ + .MaxZ) / 2)
        End If
        If Mode = 2 Then
            CenterX = Int(StartX) + (((X - Xf) / ZoomLevel) + sbBar(0))
            CenterY = Int(StartY) + (((Y - Yf) / ZoomLevel) + sbBar(1))
        End If
        If Mode = 3 Then
            CenterX = 0
            CenterY = 0
        End If
        Select Case ViewMode
            Case 1
                Corner(1).X = .MinX: Corner(1).Y = .MinY: Corner(1).z = .MinZ
                Corner(2).X = .MaxX: Corner(2).Y = .MinY: Corner(2).z = .MinZ
                Corner(3).X = .MaxX: Corner(3).Y = .MinY: Corner(3).z = .MaxZ
                Corner(4).X = .MinX: Corner(4).Y = .MinY: Corner(4).z = .MaxZ
            
            Case 2
                Corner(1).X = .MinX: Corner(1).Y = .MinY: Corner(1).z = .MaxZ
                Corner(2).X = .MaxX: Corner(2).Y = .MinY: Corner(2).z = .MaxZ
                Corner(3).X = .MaxX: Corner(3).Y = .MaxY: Corner(3).z = .MaxZ
                Corner(4).X = .MinX: Corner(4).Y = .MaxY: Corner(4).z = .MaxZ
            
            Case 3
                Corner(1).X = .MinX: Corner(1).Y = .MinY: Corner(1).z = .MinZ
                Corner(2).X = .MaxX: Corner(2).Y = .MinY: Corner(2).z = .MaxZ
                Corner(3).X = .MaxX: Corner(3).Y = .MaxY: Corner(3).z = .MaxZ
                Corner(4).X = .MinX: Corner(4).Y = .MaxY: Corner(4).z = .MinZ
        
        End Select
    End With
    
    For n = 1 To 4
        Select Case ViewMode
            Case 1
                Corner(n).X = Corner(n).X - CenterX
                Corner(n).z = Corner(n).z - CenterY
                Set Corner(n) = modFunctions.RotatePoint(Corner(n), 0, Angle, 0)
            Case 2
                Corner(n).X = Corner(n).X - CenterX
                Corner(n).Y = Corner(n).Y - CenterY
                Set Corner(n) = modFunctions.RotatePoint(Corner(n), 0, 0, Angle)
            Case 3
                Corner(n).z = Corner(n).z - CenterX
                Corner(n).Y = Corner(n).Y - CenterY
                Set Corner(n) = modFunctions.RotatePoint(Corner(n), -Angle, 0, 0)
        End Select
    Next n
    
    DrawStyle = 2
    Select Case ViewMode
        Case 1
            For n = 1 To 3
                DrawLine Corner(n).X + CenterX, Corner(n).z + CenterY, Corner(n + 1).X + CenterX, Corner(n + 1).z + CenterY
            Next n
            DrawLine Corner(4).X + CenterX, Corner(4).z + CenterY, Corner(1).X + CenterX, Corner(1).z + CenterY
            
        Case 2
            For n = 1 To 3
                DrawLine Corner(n).X + CenterX, Corner(n).Y + CenterY, Corner(n + 1).X + CenterX, Corner(n + 1).Y + CenterY
            Next n
            DrawLine Corner(4).X + CenterX, Corner(4).Y + CenterY, Corner(1).X + CenterX, Corner(1).Y + CenterY
            
        Case 3
            For n = 1 To 3
                DrawLine Corner(n).z + CenterX, Corner(n).Y + CenterY, Corner(n + 1).z + CenterX, Corner(n + 1).Y + CenterY
            Next n
            DrawLine Corner(4).z + CenterX, Corner(4).Y + CenterY, Corner(1).z + CenterX, Corner(1).Y + CenterY
            
    End Select
    DrawStyle = 0
End Sub

Public Sub DrawBoxBand(x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer)
    Cls
    DrawStyle = 2: ForeColor = vbRed
    DrawBox ((x1 - Xf) / ZoomLevel) + sbBar(0), ((y1 - Yf) / ZoomLevel) + sbBar(1), ((x2 - Xf) / ZoomLevel) + sbBar(0), ((y2 - Yf) / ZoomLevel) + sbBar(1)
    DrawStyle = 0
End Sub

Public Sub DrawScaleOutline(ScaleMode As Byte, x1 As Integer, y1 As Integer)
    Dim Am As clsObject
    Cls
    DrawStyle = 2
    ForeColor = vbRed
    x1 = Snaped(((x1 - Xf) / ZoomLevel) + sbBar(0))
    y1 = Snaped(((y1 - Yf) / ZoomLevel) + sbBar(1))
    With Model
        Select Case ViewMode
            Case 1
                Select Case ScaleMode
                    Case 1:    DrawBox .MaxX, .MaxZ, x1, y1, 1
                    Case 2:    DrawBox .MinX, .MaxZ, x1, y1, 1
                    Case 3:    DrawBox .MinX, .MinZ, x1, y1, 1
                    Case 4:    DrawBox .MaxX, .MinZ, x1, y1, 1
                    Case 5:    DrawBox .MaxX, .MinZ, x1, .MaxZ, 1
                    Case 6:    DrawBox .MinX, .MinZ, x1, .MaxZ, 1
                    Case 7:    DrawBox .MinX, .MaxZ, .MaxX, y1, 1
                    Case 8:    DrawBox .MinX, .MinZ, .MaxX, y1, 1
                End Select
        
            Case 2
                Select Case ScaleMode
                    Case 1:    DrawBox .MaxX, .MaxY, x1, y1, 1
                    Case 2:    DrawBox .MinX, .MaxY, x1, y1, 1
                    Case 3:    DrawBox .MinX, .MinY, x1, y1, 1
                    Case 4:    DrawBox .MaxX, .MinY, x1, y1, 1
                    Case 5:    DrawBox .MaxX, .MinY, x1, .MaxY, 1
                    Case 6:    DrawBox .MinX, .MinY, x1, .MaxY, 1
                    Case 7:    DrawBox .MinX, .MaxY, .MaxX, y1, 1
                    Case 8:    DrawBox .MinX, .MinY, .MaxX, y1, 1
                End Select
        
            Case 3
                Select Case ScaleMode
                    Case 1:    DrawBox .MaxZ, .MaxY, x1, y1, 1
                    Case 2:    DrawBox .MinZ, .MaxY, x1, y1, 1
                    Case 3:    DrawBox .MinZ, .MinY, x1, y1, 1
                    Case 4:    DrawBox .MaxZ, .MinY, x1, y1, 1
                    Case 5:    DrawBox .MaxZ, .MinY, x1, .MaxY, 1
                    Case 6:    DrawBox .MinZ, .MinY, x1, .MaxY, 1
                    Case 7:    DrawBox .MinZ, .MaxY, .MaxZ, y1, 1
                    Case 8:    DrawBox .MinZ, .MinY, .MaxZ, y1, 1
                End Select
    
        End Select
    End With
    DrawStyle = 0
End Sub

Public Function JointHitTest(ByVal X As Single, ByVal Y As Single) As String
    Dim Jm As clsJoint, XX As Single, YY As Single
    For Each Jm In Model.Joint
        Select Case ViewMode
            Case 1: XX = (Jm.X - sbBar(0)) * ZoomLevel + Xf: YY = (Jm.z + -sbBar(1)) * ZoomLevel + Yf
            Case 2: XX = (Jm.X - sbBar(0)) * ZoomLevel + Xf: YY = (Jm.Y + -sbBar(1)) * ZoomLevel + Yf
            Case 3: XX = (Jm.z - sbBar(0)) * ZoomLevel + Xf: YY = (Jm.Y + -sbBar(1)) * ZoomLevel + Yf
        End Select
        If Almost(XX, X, YY, Y) = True Then JointHitTest = Jm.Key: Exit Function
    Next Jm
End Function

Public Function BoxBandSelect(ByVal x1 As Integer, ByVal y1 As Integer, ByVal x2 As Integer, ByVal y2 As Integer, FullSelect As Integer)
    Dim Am As clsObject, Pm As clsVertex, ObjectSelected As Boolean
    Dim fm As clsObject, Jm As clsJoint, SelectedVertex As Integer
    x1 = ((x1 - Xf) / ZoomLevel) + sbBar(0)
    y1 = ((y1 - Yf) / ZoomLevel) + sbBar(1)
    x2 = ((x2 - Xf) / ZoomLevel) + sbBar(0)
    y2 = ((y2 - Yf) / ZoomLevel) + sbBar(1)
    If x2 > x1 Then Swap x1, x2
    If y2 > y1 Then Swap y1, y2
    If frmMain.chkSelect(1) = 1 Then
        For Each Am In Model.Geometery
            If Am.Hidden = False And Am.Locked = False And Model.Layers(Am.Layer).Selected = True And Model.Layers(Am.Layer).LayerLocked = False Then
                ObjectSelected = False
                SelectedVertex = 0
                If frmMain.chkSelect(6) = 0 Then
                    For Each Pm In Am.Vertex
                        Select Case ViewMode
                            Case 1: If Pm.X < x1 And x2 < Pm.X And Pm.z < y1 And y2 < Pm.z Then SelectedVertex = SelectedVertex + 1
                            Case 2: If Pm.X < x1 And x2 < Pm.X And Pm.Y < y1 And y2 < Pm.Y Then SelectedVertex = SelectedVertex + 1
                            Case 3: If Pm.z < x1 And x2 < Pm.z And Pm.Y < y1 And y2 < Pm.Y Then SelectedVertex = SelectedVertex + 1
                        End Select
                    Next Pm
                Else
                    For Each Pm In Am.Vertex
                        Select Case ViewMode
                            Case 1: If Pm.X < x1 And x2 < Pm.X And Pm.z < y1 And y2 < Pm.z Then Pm.Selected = True
                            Case 2: If Pm.X < x1 And x2 < Pm.X And Pm.Y < y1 And y2 < Pm.Y Then Pm.Selected = True
                            Case 3: If Pm.z < x1 And x2 < Pm.z And Pm.Y < y1 And y2 < Pm.Y Then Pm.Selected = True
                        End Select
                    Next Pm
                End If
                If (FullSelect = 0 And SelectedVertex <> 0) Or SelectedVertex = Am.Vertex.Count Then Am.Selected = True: ObjectSelected = True
                If ObjectSelected = True Then
                    If Am.Group.Count > 0 Then
                        For Each fm In Model.Geometery
                            If fm.Group.Count > 0 Then
                                If Am.Group(1).GroupID = fm.Group(1).GroupID Then fm.Selected = True
                            End If
                        Next fm
                    End If
                End If
            End If
        Next Am
    End If
    If frmMain.chkSelect(2) = 1 Then
        For Each Jm In Model.Joint
            Select Case ViewMode
                Case 1: If Jm.X < x1 And x2 < Jm.X And Jm.z < y1 And y2 < Jm.z Then Jm.Selected = True
                Case 2: If Jm.X < x1 And x2 < Jm.X And Jm.Y < y1 And y2 < Jm.Y Then Jm.Selected = True
                Case 3: If Jm.z < x1 And x2 < Jm.z And Jm.Y < y1 And y2 < Jm.Y Then Jm.Selected = True
            End Select
        Next Jm
    End If
End Function

Public Function SelectionOutlineHittest(ByVal X As Single, ByVal Y As Single) As Integer
    X = ((X - Xf) / ZoomLevel) + sbBar(0)
    Y = ((Y - Yf) / ZoomLevel) + sbBar(1)
    With Model
        Select Case ViewMode
            Case 1
                If Almost(X, .MinX, Y, .MinZ) = True Then SelectionOutlineHittest = 1
                If Almost(X, .MaxX, Y, .MinZ) = True Then SelectionOutlineHittest = 2
                If Almost(X, .MaxX, Y, .MaxZ) = True Then SelectionOutlineHittest = 3
                If Almost(X, .MinX, Y, .MaxZ) = True Then SelectionOutlineHittest = 4
                If Almost(X, .MinX, Y, (.MinZ + .MaxZ) * 0.5) = True Then SelectionOutlineHittest = 5
                If Almost(X, .MaxX, Y, (.MinZ + .MaxZ) * 0.5) = True Then SelectionOutlineHittest = 6
                If Almost(X, (.MinX + .MaxX) * 0.5, Y, .MinZ) = True Then SelectionOutlineHittest = 7
                If Almost(X, (.MinX + .MaxX) * 0.5, Y, .MaxZ) = True Then SelectionOutlineHittest = 8
    
            Case 2
                If Almost(X, .MinX, Y, .MinY) = True Then SelectionOutlineHittest = 1
                If Almost(X, .MaxX, Y, .MinY) = True Then SelectionOutlineHittest = 2
                If Almost(X, .MaxX, Y, .MaxY) = True Then SelectionOutlineHittest = 3
                If Almost(X, .MinX, Y, .MaxY) = True Then SelectionOutlineHittest = 4
                If Almost(X, .MinX, Y, (.MinY + .MaxY) * 0.5) = True Then SelectionOutlineHittest = 5
                If Almost(X, .MaxX, Y, (.MinY + .MaxY) * 0.5) = True Then SelectionOutlineHittest = 6
                If Almost(X, (.MinX + .MaxX) * 0.5, Y, .MinY) = True Then SelectionOutlineHittest = 7
                If Almost(X, (.MinX + .MaxX) * 0.5, Y, .MaxY) = True Then SelectionOutlineHittest = 8
    
            Case 3
                If Almost(X, .MinZ, Y, .MinY) = True Then SelectionOutlineHittest = 1
                If Almost(X, .MaxZ, Y, .MinY) = True Then SelectionOutlineHittest = 2
                If Almost(X, .MaxZ, Y, .MaxY) = True Then SelectionOutlineHittest = 3
                If Almost(X, .MinZ, Y, .MaxY) = True Then SelectionOutlineHittest = 4
                If Almost(X, .MinZ, Y, (.MinY + .MaxY) * 0.5) = True Then SelectionOutlineHittest = 5
                If Almost(X, .MaxZ, Y, (.MinY + .MaxY) * 0.5) = True Then SelectionOutlineHittest = 6
                If Almost(X, (.MinZ + .MaxZ) * 0.5, Y, .MinY) = True Then SelectionOutlineHittest = 7
                If Almost(X, (.MinZ + .MaxZ) * 0.5, Y, .MaxY) = True Then SelectionOutlineHittest = 8
    
        End Select
    End With
End Function

Public Function CenterView()
    Dim XX As Integer, YY As Integer
    Select Case ViewMode
        Case 1: XX = (Model.MinX + Model.MaxX) * 0.5: YY = (Model.MinZ + Model.MaxZ) * 0.5
        Case 2: XX = (Model.MinX + Model.MaxX) * 0.5: YY = (Model.MinY + Model.MaxY) * 0.5
        Case 3: XX = (Model.MinZ + Model.MaxZ) * 0.5: YY = (Model.MinY + Model.MaxY) * 0.5
    End Select
    sbBar(0) = XX: sbBar(1) = YY
End Function

Public Function ZoomToSelected() As Single
    'This returns the size of zoom required to make the selected objects fill up the entire screen.
    If Model.Geometery.CountSelected + Model.Joint.CountSelected = 0 Then ZoomToSelected = ZoomLevel: Exit Function
    Dim X As Single, Y As Single, z As Single
    With Model
        X = (.MaxX - .MinX) * 1.1
        Y = (.MaxY - .MinY) * 1.1
        z = (.MaxZ - .MinZ) * 1.1
    End With
    If X + Y + z = 0 Then ZoomToSelected = ZoomLevel: Exit Function
    Select Case ViewMode
        Case 1
            If Abs(ScaleWidth - z) < Abs(ScaleHeight - X) Then
                If Abs(ScaleHeight - X) = 0 Then
                    ZoomToSelected = ZoomLevel
                Else
                    ZoomToSelected = (ScaleHeight / X)
                End If
            Else
                If Abs(ScaleWidth - z) = 0 Then
                    ZoomToSelected = ZoomLevel
                Else
                    ZoomToSelected = (ScaleWidth / z)
                End If
            End If

            
        Case 2
            If Abs(ScaleWidth - X) > Abs(ScaleHeight - Y) Then
                If Abs(ScaleWidth - X) = 0 Then
                    ZoomToSelected = ZoomLevel
                Else
                    ZoomToSelected = (ScaleWidth / X)
                End If
            Else
                If Abs(ScaleHeight - Y) = 0 Then
                    ZoomToSelected = ZoomLevel
                Else
                    ZoomToSelected = (ScaleHeight / Y)
                End If
            End If
            
            
        Case 3
            If Abs(ScaleHeight - z) > Abs(ScaleWidth - Y) Then
                If Abs(ScaleWidth - z) = 0 Then
                    ZoomToSelected = ZoomLevel
                Else
                    ZoomToSelected = (ScaleWidth / z)
                End If
            Else
                If Abs(ScaleHeight - Y) = 0 Then
                    ZoomToSelected = ZoomLevel
                Else
                    ZoomToSelected = (ScaleHeight / Y)
                End If
            End If
    End Select
End Function






Public Sub SetDragDropStyle(Style As Integer)
    OLEDropMode = Style
End Sub

Public Sub SetBorderStyle(Style As Integer)
    Appearance = Style
End Sub

Public Sub SetScrollBarStyle(Style As Integer)
    sbBar(0).Visible = Style
    sbBar(1).Visible = Style
    cmdBlock.Visible = Style
End Sub

Public Property Get FileKey() As String
    FileKey = Model.Key
End Property

Public Sub AssignTabletTo(AssignedModel As clsFile)
    Set Model = AssignedModel
End Sub

Private Sub sbBar_Change(Index As Integer)
    UserControl_Paint
End Sub

Private Sub sbBar_Scroll(Index As Integer)
    UserControl_Paint
    DoEvents
End Sub

Private Sub UserControl_Initialize()
    ZoomLevel = 1
End Sub

Public Function CancelCreateObject()
    If ShapeX1 + ShapeX2 + ShapeY1 + ShapeY2 <> 0 Then
        ShapeX1 = 0:        ShapeX2 = 0
        ShapeY1 = 0:        ShapeY2 = 0
        Set Model.Wraper = New colWrap
        AutoRedraw = True
        Refresh
    End If
End Function

