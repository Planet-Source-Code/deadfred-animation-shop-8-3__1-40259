VERSION 5.00
Begin VB.UserControl Engine 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ForeColor       =   &H00FFFFFF&
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.Timer tmAnimate 
      Left            =   2160
      Top             =   1680
   End
End
Attribute VB_Name = "Engine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'#####################################################################
'#                                                                   #
'#  This control displays a file object in 3D using software code    #
'#  rather than external libararys. It will always run corretly,     #
'#  but is some what slower than directX render methods.             #
'#                                                                   #
'#####################################################################

Const zeye As Single = 800


Private Model As clsFile

Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long) As Long


Private Type POINTAPI
    X As Long
    y As Long
End Type
Private API As POINTAPI


Private Type typVertex
    XX As Integer           '3D X position of vertex
    YY As Integer           '3D Y position of vertex
    Zz As Integer           '3D Z position of vertex
    X As Integer            '2D X position of vertex
    y As Integer            '2D Y position of vertex
    TexXX As Integer        '2D X position of vertex on the texture map
    TexYY As Integer        '2D Y position of vertex on the texture map
End Type


Private Type typLine
    StartEntered As Integer 'Sets whether the start details of the line have been found
    StartX As Integer       '2D Start X position of the line
    StartY As Integer       '2D Start Y position of the line
    EndX As Integer         '2D End X position of the line
    EndY As Integer         '2D End Y position of the line
    XXStart As Single       '3D Start X position of the line
    YYStart As Single       '3D Start Y position of the line
    ZZStart As Single       '3D Start Z position of the line
    XXEnd As Single         '3D End X position of the line
    YYEnd As Single         '3D End Y position of the line
    ZZEnd As Single         '3D End Z position of the line
    XXStartTx As Single     '2D Start X position of the line on the texture map
    YYStartTx As Single     '2D Start Y position of the line on the texture map
    XXEndTx As Single       '2D Start X position of the line on the texture map
    YYEndTx As Single       '2D Start Y position of the line on the texture map
End Type


Public Event MouseDown(X As Single, y As Single, Button As Integer, Shift As Integer)
Public Event MouseMove(X As Single, y As Single, Button As Integer, Shift As Integer)
Public Event MouseUp(X As Single, y As Single, Button As Integer, Shift As Integer)

Private ZBuffer() As Long
Private IBuffer() As Long
Private TBuffer() As Byte
Private SBuffer() As Integer

Private DrawingShadow As Integer

Public pAutoZoom As Boolean
Public pAllFace As Boolean
Public pHighlightFace As Boolean
Public pHightlightVertex As Boolean
Public pDrawObjects As Boolean
Public pPerspecitve As Boolean
Public pDrawJoints As Boolean
Public pLabelJoints As Boolean
Public pRenderSolid As Boolean
Public pRenderFX As Boolean
Public pRenderShadow As Boolean
Public pRenderTexture As Boolean
Public ShapeFX As Integer
Public pAutoRotate As Boolean
Public pPaintMap As Boolean
Public pSelectedOnly As Boolean
Public pFacePreView As Boolean
Public pDrawEdgePreview As Boolean
Public pDrawSkeliton As Boolean
Public pNameJoints As Boolean
Public pDrawOrigin As Boolean
Public pClipFaces As Boolean
Public pClipLine As Integer
Public pQuickDraw As Boolean

Public FaceOver As Integer
Public VertexOver As Integer


Public PicketAnimate As Integer
Public PicketAnimateOver As Integer

Private lDiffuse As Integer
Private lGrain As Integer
Private lTransparant As Integer
Private lBrushRed As Integer
Private lBrushGreen As Integer
Private lBrushBlue As Integer

Public ZoomLevel As Single
Public Angle1 As Single, Angle2 As Single, Angle3 As Single
Private MouseX As Single, MouseY As Single, PaintColour As Long
Private Xf As Integer, Yf As Integer
Private Tx(75, 75, 3) As Integer

Private Type FaceDis
    vert(35) As clsVertex
    EdgeCount As Integer
End Type

Public AnimationFrame As Integer
Public AnimationSceneName As String



Private Sub tmAnimate_Timer()
    tmAnimate.Tag = tmAnimate.Tag - 1
    If tmAnimate.Tag = 0 Then tmAnimate.Interval = 0
    Model.Scene.MoveAnimation
    If frmMain.ActiveForm.mnuTools(11).Checked = True Then SavePicture Image, AnimationSceneName & ThreeLength(AnimationFrame) & ".bmp"
    AnimationFrame = AnimationFrame + 1
    RefreshView
End Sub




Private Function SliceFace(Face As FaceDis, SliceLine As Integer) As FaceDis
    Dim Cliped As FaceDis, NewVert As Byte, Start As Byte, n As Byte
    Dim Xx1 As Integer, Xx2 As Integer, Yy1 As Integer, Yy2 As Integer
    Dim Perc As Single, Zz1 As Integer, Zz2 As Integer
    For n = 1 To 35: Set Cliped.vert(n) = New clsVertex: Next n
    For n = 1 To 35: Set SliceFace.vert(n) = New clsVertex: Next n
    NewVert = 1: Start = 0
    For n = 1 To Face.EdgeCount
        If Face.vert(n).y <= SliceLine Then Start = n: Exit For
    Next n
    If Start = 0 Then Exit Function
    If Start <> 1 Then
        If Face.vert(Face.EdgeCount).y < SliceLine Then
            Xx1 = Face.vert(Face.EdgeCount).X
            Yy1 = Face.vert(Face.EdgeCount).y
            Zz1 = Face.vert(Face.EdgeCount).z
            Xx2 = Face.vert(1).X
            Yy2 = Face.vert(1).y
            Zz2 = Face.vert(1).z
            If Yy2 <> SliceLine Then Perc = (Yy2 - SliceLine) / (Yy2 - Yy1)
            Cliped.vert(NewVert).X = Xx2 - ((Xx2 - Xx1) * Perc)
            Cliped.vert(NewVert).y = SliceLine
            Cliped.vert(NewVert).z = Zz2 - ((Zz2 - Zz1) * Perc)
            NewVert = NewVert + 1
        End If
        Xx1 = Face.vert(Start - 1).X
        Yy1 = Face.vert(Start - 1).y
        Zz1 = Face.vert(Start - 1).z
        Xx2 = Face.vert(Start).X
        Yy2 = Face.vert(Start).y
        Zz2 = Face.vert(Start).z
        If Yy2 <> SliceLine Then Perc = (Yy2 - SliceLine) / (Yy2 - Yy1)
        Cliped.vert(NewVert).X = Xx2 - ((Xx2 - Xx1) * Perc)
        Cliped.vert(NewVert).y = SliceLine
        Cliped.vert(NewVert).z = Zz2 - ((Zz2 - Zz1) * Perc)
        NewVert = NewVert + 1
    End If
    Cliped.vert(NewVert).X = Face.vert(Start).X
    Cliped.vert(NewVert).y = Face.vert(Start).y
    Cliped.vert(NewVert).z = Face.vert(Start).z
    For n = Start To Face.EdgeCount
        If n = Face.EdgeCount Then
            Xx1 = Face.vert(n).X
            Yy1 = Face.vert(n).y
            Zz1 = Face.vert(n).z
            Xx2 = Face.vert(Start).X
            Yy2 = Face.vert(Start).y
            Zz2 = Face.vert(Start).z
        Else
            Xx1 = Face.vert(n).X
            Yy1 = Face.vert(n).y
            Zz1 = Face.vert(n).z
            Xx2 = Face.vert(n + 1).X
            Yy2 = Face.vert(n + 1).y
            Zz2 = Face.vert(n + 1).z
        End If
        If Yy1 <= SliceLine And Yy2 <= SliceLine Then
            NewVert = NewVert + 1
            Cliped.vert(NewVert).X = Xx2
            Cliped.vert(NewVert).y = Yy2
            Cliped.vert(NewVert).z = Zz2
        ElseIf Yy1 >= SliceLine And Yy2 <= SliceLine Then
            If Yy2 <> SliceLine Then Perc = (Yy2 - SliceLine) / (Yy2 - Yy1)
            NewVert = NewVert + 1
            Cliped.vert(NewVert).X = Xx2 - ((Xx2 - Xx1) * Perc)
            Cliped.vert(NewVert).y = SliceLine
            Cliped.vert(NewVert).z = Zz2 - ((Zz2 - Zz1) * Perc)
            NewVert = NewVert + 1
            Cliped.vert(NewVert).X = Xx2
            Cliped.vert(NewVert).y = Yy2
            Cliped.vert(NewVert).z = Zz2
        ElseIf Yy1 <= SliceLine And Yy2 >= SliceLine Then
            If Yy2 <> SliceLine Then Perc = (Yy2 - SliceLine) / (Yy2 - Yy1)
            NewVert = NewVert + 1
            Cliped.vert(NewVert).X = Xx2 - ((Xx2 - Xx1) * Perc)
            Cliped.vert(NewVert).y = SliceLine
            Cliped.vert(NewVert).z = Zz2 - ((Zz2 - Zz1) * Perc)
        End If
    Next n
    SliceFace.EdgeCount = NewVert - 1
    For n = 1 To NewVert - 1
        SliceFace.vert(n).X = Cliped.vert(n).X
        SliceFace.vert(n).y = Cliped.vert(n).y
        SliceFace.vert(n).z = Cliped.vert(n).z
    Next n
End Function




Private Function ClipFace(Face As FaceDis) As Boolean
    Dim Cliped As FaceDis, NewVert As Byte, Start As Byte, n As Byte
    Dim Xx1 As Integer, Xx2 As Integer, Yy1 As Integer, Yy2 As Integer
    Dim Perc As Single, Zz1 As Integer, Zz2 As Integer
    For n = 1 To 35: Set Cliped.vert(n) = New clsVertex: Next n
    NewVert = 1: Start = 0
    For n = 1 To Face.EdgeCount
        If Face.vert(n).z <= pClipLine Then Start = n: Exit For
    Next n
    If Start = 0 Then Exit Function
    If Start <> 1 Then
        If Face.vert(Face.EdgeCount).z < pClipLine Then
            Xx1 = Face.vert(Face.EdgeCount).X
            Yy1 = Face.vert(Face.EdgeCount).y
            Zz1 = Face.vert(Face.EdgeCount).z
            Xx2 = Face.vert(1).X
            Yy2 = Face.vert(1).y
            Zz2 = Face.vert(1).z
            If Zz2 <> pClipLine Then Perc = (Zz2 - pClipLine) / (Zz2 - Zz1)
            Cliped.vert(NewVert).X = Xx2 - ((Xx2 - Xx1) * Perc)
            Cliped.vert(NewVert).y = Yy2 - ((Yy2 - Yy1) * Perc)
            Cliped.vert(NewVert).z = pClipLine
            NewVert = NewVert + 1
        End If
        Xx1 = Face.vert(Start - 1).X
        Yy1 = Face.vert(Start - 1).y
        Zz1 = Face.vert(Start - 1).z
        Xx2 = Face.vert(Start).X
        Yy2 = Face.vert(Start).y
        Zz2 = Face.vert(Start).z
        If Zz2 <> pClipLine Then Perc = (Zz2 - pClipLine) / (Zz2 - Zz1)
        Cliped.vert(NewVert).X = Xx2 - ((Xx2 - Xx1) * Perc)
        Cliped.vert(NewVert).y = Yy2 - ((Yy2 - Yy1) * Perc)
        Cliped.vert(NewVert).z = pClipLine
        NewVert = NewVert + 1
    End If
    Cliped.vert(NewVert).X = Face.vert(Start).X
    Cliped.vert(NewVert).y = Face.vert(Start).y
    Cliped.vert(NewVert).z = Face.vert(Start).z
    For n = Start To Face.EdgeCount
        If n = Face.EdgeCount Then
            Xx1 = Face.vert(n).X
            Yy1 = Face.vert(n).y
            Zz1 = Face.vert(n).z
            Xx2 = Face.vert(Start).X
            Yy2 = Face.vert(Start).y
            Zz2 = Face.vert(Start).z
        Else
            Xx1 = Face.vert(n).X
            Yy1 = Face.vert(n).y
            Zz1 = Face.vert(n).z
            Xx2 = Face.vert(n + 1).X
            Yy2 = Face.vert(n + 1).y
            Zz2 = Face.vert(n + 1).z
        End If
        If Zz1 <= pClipLine And Zz2 <= pClipLine Then
            NewVert = NewVert + 1
            Cliped.vert(NewVert).X = Xx2
            Cliped.vert(NewVert).y = Yy2
            Cliped.vert(NewVert).z = Zz2
        ElseIf Zz1 >= pClipLine And Zz2 <= pClipLine Then
            If Zz2 <> pClipLine Then Perc = (Zz2 - pClipLine) / (Zz2 - Zz1)
            NewVert = NewVert + 1
            Cliped.vert(NewVert).X = Xx2 - ((Xx2 - Xx1) * Perc)
            Cliped.vert(NewVert).y = Yy2 - ((Yy2 - Yy1) * Perc)
            Cliped.vert(NewVert).z = pClipLine
            NewVert = NewVert + 1
            Cliped.vert(NewVert).X = Xx2
            Cliped.vert(NewVert).y = Yy2
            Cliped.vert(NewVert).z = Zz2
        ElseIf Zz1 <= pClipLine And Zz2 >= pClipLine Then
            If Zz2 <> pClipLine Then Perc = (Zz2 - pClipLine) / (Zz2 - Zz1)
            NewVert = NewVert + 1
            Cliped.vert(NewVert).X = Xx2 - ((Xx2 - Xx1) * Perc)
            Cliped.vert(NewVert).y = Yy2 - ((Yy2 - Yy1) * Perc)
            Cliped.vert(NewVert).z = pClipLine
        End If
    Next n
    Dim Ner(30) As typVertex
    For n = 1 To NewVert - 1
        If pPerspecitve Then
            Ner(n).X = (Xf + Cliped.vert(n).X * (zeye / (zeye - Cliped.vert(n).z)))
            Ner(n).y = (Yf + Cliped.vert(n).y * (zeye / (zeye - Cliped.vert(n).z)))
        Else
            Ner(n).X = (Xf + Cliped.vert(n).X)
            Ner(n).y = (Yf + Cliped.vert(n).y)
        End If
    Next n
    If FaceNormal(Ner(), NewVert - 1) < 0 And pAllFace = True Then Exit Function
    MoveTo Ner(1).X, Ner(1).y
    For n = 1 To NewVert - 1
        DrawTo Ner(n).X, Ner(n).y
    Next n: DrawTo Ner(1).X, Ner(1).y
End Function














Public Sub BeginRotate(X As Integer, y As Integer)
    MouseX = X
    MouseY = y
End Sub


Public Sub BeginPaint(X As Integer, y As Integer, Colour As Long)
    ShapeFX = 6
    MouseX = X
    MouseY = y
    PaintColour = Colour
    DrawObject3D
    ShapeFX = 2
    UserControl_Paint
End Sub


Public Property Get FileKey() As String
    FileKey = Model.Key
End Property



Public Sub AssignEngineTo(AssignedModel As clsFile)
    'This sets the model instance to point to an existing file class
    Set Model = AssignedModel
End Sub



Public Function DrawObject3D() As Boolean
    'This draws the object in 3D onto the tablet given in the parameter line. There are
    'several optional parameters to alter how it is drawn. ShowVerteceis highlights each veretx,
    'showface highlights each face, and ObjectSelected highlights the entire object.
    Dim Coord(25) As typVertex, n As Integer, m As Integer, NotRotated As clsVertex
    Dim FaceON As clsFace, VertexOn As clsEdge, CenX As Single, CenY As Single
    Dim Am As clsObject, Rotated As clsVertex, FaceCounter As Integer, Pm As clsJoint
    Dim Nx As Single, Ny As Single, Vm As clsVertex, XX As Integer, YY As Integer
    Dim NewFace As FaceDis, CenZ As Single
    For m = 1 To 35: Set NewFace.vert(m) = New clsVertex: Next m
    

    Model.MorphSkeliton "BaseFrame", "Animate"
    
    If pRenderSolid = True Then
        For n = 0 To ScaleHeight
            For m = 0 To ScaleWidth
                ZBuffer(n, m) = 0
                SBuffer(n, m) = 0
                TBuffer(n, m) = 0
                IBuffer(n, m) = 0
            Next m
        Next n
    End If
    
    
    
    If pDrawObjects = True Then
        For Each Am In Model.Geometery
            If Am.Selected = True Or pSelectedOnly = False And Model.Layers(Am.Layer).Selected = True Then
                If pHightlightVertex = True Then
                    For Each Vm In Am.Vertex
                        n = n + 1
                        Set Rotated = RotatePoint(Vm, Angle1, Angle2, Angle3)
                        If pPerspecitve = True Then
                            XX = (Xf + (Rotated.X * ZoomLevel) * (zeye / (zeye - Rotated.z)))
                            YY = (Yf + (Rotated.y * ZoomLevel) * (zeye / (zeye - Rotated.z)))
                        Else
                            XX = Xf - (Rotated.X * ZoomLevel)
                            YY = Yf - (Rotated.y * ZoomLevel)
                        End If
                        If Almost(MouseX, CSng(XX), MouseY, CSng(YY)) = True Then VertexOver = n
                        If Vm.Selected = True Then
                            ForeColor = vbRed
                            DrawCircle XX, YY, 4
                            DrawCircle XX, YY, 3
                        Else
                            ForeColor = vbWhite
                            DrawCircle XX, YY, 4
                        End If
                    Next Vm
                End If
                n = 0
                If Am.Colour = 0 Then ForeColor = vbWhite Else ForeColor = Am.Colour
                FaceCounter = 0
                
                
                
                For Each FaceON In Am.Face
                    FaceCounter = FaceCounter + 1
                    n = 0: CenX = 0: CenY = 0: CenZ = 0
                    For Each VertexOn In FaceON.Edge
                        n = n + 1
                        If Am.Key = "GroundPlain" Then
                            Set Rotated = RotatePoint(Am.Vertex(VertexOn.Vertex), 0, 0, 0)
                        ElseIf Am.Vertex(VertexOn.Vertex).TargetName <> "" Then
                            With Model.Joint(Am.Vertex(VertexOn.Vertex).TargetName)
                                Set Rotated = RotatePoint(Am.Vertex(VertexOn.Vertex), 0, 0, 0, 0, 0, 0)
                                Rotated.X = Rotated.X - (.X - .NewPositX)
                                Rotated.y = Rotated.y - (.y - .NewPositY)
                                Rotated.z = Rotated.z - (.z - .NewPositZ)
                                Set Rotated = RotatePoint(Rotated, .AngleX, .AngleY, .AngleZ, .NewPositX, .NewPositY, .NewPositZ)
                                Set Rotated = RotatePoint(Rotated, Angle1, Angle2, Angle3)
                            End With
                        Else
                            Set Rotated = RotatePoint(Am.Vertex(VertexOn.Vertex), Angle1, Angle2, Angle3)
                        End If
                        NewFace.vert(n).X = (Rotated.X * ZoomLevel)
                        NewFace.vert(n).y = (Rotated.y * ZoomLevel)
                        NewFace.vert(n).z = (Rotated.z * ZoomLevel)
                        If Am.TexVert.Count >= VertexOn.TexVertex And VertexOn.TexVertex <> 0 Then
                            Coord(n).TexXX = Am.TexVert(VertexOn.TexVertex).X
                            Coord(n).TexYY = Am.TexVert(VertexOn.TexVertex).y
                        End If
                    Next VertexOn

                    
                    NewFace.EdgeCount = n
                    For m = 1 To n
                        CenX = CenX + NewFace.vert(m).X
                        CenY = CenY + NewFace.vert(m).y
                        CenZ = CenZ + NewFace.vert(m).z
                    Next m


                    If pRenderSolid = True Then
                        For m = 1 To n
                            Coord(m).X = (Xf + NewFace.vert(m).X * (zeye / (zeye - NewFace.vert(m).z)))
                            Coord(m).y = (Yf + NewFace.vert(m).y * (zeye / (zeye - NewFace.vert(m).z)))
                            Coord(m).XX = NewFace.vert(m).X
                            Coord(m).YY = NewFace.vert(m).y
                            Coord(m).Zz = NewFace.vert(m).z
                        Next m
                        If FaceNormal(Coord(), FaceON.Edge.Count) > 0 Or Am.ForceShowFace = True Then
                            ReDim LBuffer(ScaleHeight) As typLine
                            lTransparant = Am.Transparancy
                            lGrain = Am.grain
                            lDiffuse = Am.Diffusion
                            If Am.Colour = 0 Then
                                lBrushRed = 255: lBrushGreen = 255: lBrushBlue = 255
                            Else
                                lBrushRed = Am.Colour And 255
                                lBrushGreen = (Am.Colour And 65280) / 256
                                lBrushBlue = (Am.Colour And 16711680) / 65536
                            End If
                            If Am.IsShadow = True Then DrawingShadow = Sgn(FaceNormal(Coord(), FaceON.Edge.Count)) Else DrawingShadow = 0
                            For n = 1 To NewFace.EdgeCount
                                If n = NewFace.EdgeCount Then ScanEdge LBuffer(), Coord(n), Coord(1) Else ScanEdge LBuffer(), Coord(n), Coord(n + 1)
                            Next n
                        End If
                    Else
                        If pClipFaces = True Then ClipFace SliceFace(NewFace, 50) Else ClipFace NewFace
                    End If
                    
                    
                    If pHighlightFace = True Then
                        CenX = CenX / n
                        CenY = CenY / n
                        CenZ = CenZ / n
                        CenX = (CenX * (zeye / (zeye - CenZ)))
                        CenY = (CenY * (zeye / (zeye - CenZ)))
                        DrawCircle Int(CenX) + Xf, Int(CenY) + Yf, 4
                        If Almost(Int(CenX + Xf), MouseX, Int(CenY + Yf), MouseY) = True Then FaceOver = FaceCounter
                    End If
                    
                    
                Next FaceON
                
                
                If pDrawEdgePreview = True Then
                    If frmObject.lstFaceOrder.ListCount <> 0 Then
                        For n = 1 To frmObject.lstFaceOrder.ListCount
                            Set Rotated = RotatePoint(Am.Vertex(frmObject.lstFaceOrder.List(n - 1)), Angle1, Angle2, Angle3)
                            If pPerspecitve = True Then
                                Coord(n).X = (Xf + (Rotated.X * ZoomLevel) * (zeye / (zeye - Rotated.z)))
                                Coord(n).y = (Yf + (Rotated.y * ZoomLevel) * (zeye / (zeye - Rotated.z)))
                            Else
                                Coord(n).X = Xf + (Rotated.X * ZoomLevel)
                                Coord(n).y = Yf + (Rotated.y * ZoomLevel)
                            End If
                        Next n
                        If FaceNormal(Coord(), frmObject.lstFaceOrder.ListCount) < 0 Then ForeColor = vbRed Else ForeColor = vbGreen
                        MoveTo Coord(frmObject.lstFaceOrder.ListCount).X, Coord(frmObject.lstFaceOrder.ListCount).y
                        For n = 1 To frmObject.lstFaceOrder.ListCount
                            DrawTo Coord(n).X, Coord(n).y
                        Next n
                    End If
                End If
                
            End If
        Next Am
    End If
    
    
    
    If pDrawSkeliton = True Then
        'If Am8(Active).Geometery.CountSelected <> 0 And ShowSelection = True Then Outline3DSelection Angle1, Angle2, Angle3
        Set NotRotated = New clsVertex
        For Each Pm In Model.Joint
            ReDim Vertex(1) As clsVertex
            Set Vertex(1) = New clsVertex
            Vertex(1).X = Pm.NewPositX: Vertex(1).y = Pm.NewPositY: Vertex(1).z = Pm.NewPositZ
            Set Vertex(1) = RotatePoint(Vertex(1), Angle1, Angle2, Angle3)
            If pPerspecitve = True Then
                XX = (Xf + Vertex(1).X * (800 / (800 - Vertex(1).z)))
                YY = (Yf + Vertex(1).y * (800 / (800 - Vertex(1).z)))
            Else
                XX = (Xf + Vertex(1).X)
                YY = (Yf + Vertex(1).y)
            End If
            DrawCircle XX, YY, 4
            If pLabelJoints = True Then PrintText XX, YY, Pm.Name
            If Pm.Target <> "" Then
                Vertex(1).X = Model.Joint(Pm.Target).NewPositX
                Vertex(1).y = Model.Joint(Pm.Target).NewPositY
                Vertex(1).z = Model.Joint(Pm.Target).NewPositZ
                Set Vertex(1) = RotatePoint(Vertex(1), Angle1, Angle2, Angle3)
                If pPerspecitve = True Then
                    Nx = (Xf + Vertex(1).X * (800 / (800 - Vertex(1).z)))
                    Ny = (Yf + Vertex(1).y * (800 / (800 - Vertex(1).z)))
                Else
                    Nx = (Xf + Vertex(1).X)
                    Ny = (Yf + Vertex(1).y)
                End If
                DrawLine XX, YY, Int(Nx), Int(Ny)
            End If
        Next Pm
    End If
    
    
    
'    If pRenderSolid = True Then
'        For n = 0 To ScaleHeight
'            For m = 0 To ScaleWidth
'                If SBuffer(n, m) = 0 Then
'                    SetPixel hdc, m, n, IBuffer(n, m)
'                Else
'                    If IBuffer(n, m) <> 0 Then
'                        SetPixel hdc, m, n, vbRed
'                    End If
'                End If
'            Next m
'        Next n
'    End If
    

    
    
End Function


Public Sub PrintText(XX As Integer, YY As Integer, Text As String)
    CurrentX = XX
    CurrentY = YY
    Print Text
End Sub

Public Sub RefreshView()
    UserControl_Paint
End Sub


Public Sub SetTimer(Tag As Integer, Interval As Integer)
    tmAnimate.Tag = Tag
    tmAnimate.Interval = Interval
End Sub




Private Sub UserControl_Initialize()
    ZoomLevel = 1
    pClipLine = 700
End Sub



Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    RaiseEvent MouseDown(X, y, Button, Shift)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    'This automates the rotation of the model whenever you drag the mouse across
    'the control
    Static OldX As Integer, OldY As Integer, TempDR As Boolean
    RaiseEvent MouseMove(X, y, Button, Shift)
    If Button = 1 And pAutoRotate = True Then
        Angle1 = Angle1 - (y - OldY) * 10
        Angle2 = Angle2 - (X - OldX) * 10
        Angle1 = Angle1 Mod 3600
        Angle2 = Angle2 Mod 3600
        Angle3 = Angle3 Mod 3600
        If pQuickDraw = True Then
            TempDR = pRenderSolid
            pRenderSolid = False
            UserControl_Paint
            pRenderSolid = TempDR
        Else
            UserControl_Paint
        End If
            
    End If
    If Button = 2 And pAutoZoom = True Then
        Me.ZoomLevel = ZoomLevel - ((y - OldY) / 100)
        If ZoomLevel < 0.25 Then ZoomLevel = 0.25
        If ZoomLevel > 8 Then ZoomLevel = 8
        UserControl_Paint
    End If
    
    OldX = X: OldY = y
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    RaiseEvent MouseUp(X, y, Button, Shift)
    UserControl_Paint
End Sub




Private Sub UserControl_Paint()
    'This redraws the model whenever the control is repainted
    Cls
    DrawObject3D
    If pDrawOrigin = True Then DrawGuides
    Refresh
End Sub



Public Sub DrawTo(x1 As Integer, y1 As Integer)
    'This uses the LineTo command to draw from where ever the cursor was before, to
    'a new position, which is faster than having to set the cursor every time.
    LineTo hdc, x1, y1
End Sub



Public Sub MoveTo(x1 As Integer, y1 As Integer)
    'This function uses the MoveTo command on its own, to set the position of the
    'cursor on the screen. By using the DrawTo command, you can draw from whereever
    'the cursor was to a new position, which is faster that seting the cursor position
    'every time
    MoveToEx hdc, x1, y1, API
End Sub



Private Function FaceNormal(Ner() As typVertex, Edges As Byte) As Double
    'This takes all the corners of a face, and calculates the FaceNormal.
    'Instad of just working on the first 3 points, it takes points from
    'all around the face to give a more acurate answer.
    On Error Resume Next
    Select Case Edges
        Case 3, 4
            FaceNormal = (CLng((Ner(1).y - Ner(3).y)) * CLng((Ner(2).X - Ner(1).X))) - (CLng((Ner(1).X - Ner(3).X)) * CLng((Ner(2).y - Ner(1).y)))
        Case 5, 6, 7, 8, 9, 10
            FaceNormal = (CLng((Ner(1).y - Ner(5).y)) * CLng((Ner(3).X - Ner(1).X))) - (CLng((Ner(1).X - Ner(5).X)) * CLng((Ner(3).y - Ner(1).y)))
        Case Is > 10
            FaceNormal = (CLng((Ner(1).y - Ner(9).y)) * CLng((Ner(5).X - Ner(1).X))) - (CLng((Ner(1).X - Ner(9).X)) * CLng((Ner(5).y - Ner(1).y)))
    End Select
End Function



Public Sub DrawLine(x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer)
    'This draws a line onto the screen, taking into acount the scroll bars and
    'zoom factor. This will be called quite abit
    MoveToEx hdc, x1, y1, API
    LineTo hdc, x2, y2
End Sub



Public Sub DrawCircle(X As Integer, y As Integer, iWidth As Integer)
    'This draws a circle onto the screen, taking into acount the zoom and scroll bars
    Ellipse hdc, X - iWidth, y - iWidth, X + iWidth, y + iWidth
End Sub



Private Sub ScanEdge(ByRef LBuffer() As typLine, Vert1 As typVertex, Vert2 As typVertex)
    'This is the next level of the render process. It takes the start and end point of a line, and
    'fills in the values between the start and end points. It works out the 3D and 2D locations at each
    'pixel along the line, and a value between 0 and 1 showing the percentage of the distance along the
    'line that each pixel is at
    
    Dim HoriLine As Integer, Fragmenter As Single
    Dim CurrentX As Single, XShift As Single
    Dim CurrentY As Single, YShift As Single
    Dim CurrentXX As Single, XXShift As Single
    Dim CurrentYY As Single, YYShift As Single
    Dim CurrentZZ As Single, ZZShift As Single
    Dim CurrentTx As Single, TYShift As Single
    Dim CurrentTy As Single, TXShift As Single
    
    'If the edge being scanned is horizontal then just ignore the line completely
    If Vert1.y = Vert2.y Then Exit Sub
    
    'This calculates a value to find the number of vertical steps it takes to get from Y1 to Y2.
    'It is used repeatedly to move variables smoothly from one value to another. The 1/ means that
    'I can use a multiply later on, and redunce the number of devides used, as deviders are slower
    Fragmenter = 1 / Abs(Vert1.y - Vert2.y)
    
    'This is a big list of all the variables that need to be blened from one value to another
    'over the vertical distance of the face. Each variable is set at a start value. The ??Shift
    'variable in each line holds the value that needs to be added to the start value to get it to
    'the required end value by the time the complete edge has been scanned.
    CurrentX = Vert1.X: XShift = (Vert2.X - Vert1.X) * Fragmenter
    CurrentY = Vert1.y: YShift = (Vert2.y - Vert1.y) * Fragmenter
    CurrentXX = Vert1.XX: XXShift = (Vert2.XX - Vert1.XX) * Fragmenter
    CurrentYY = Vert1.YY: YYShift = (Vert2.YY - Vert1.YY) * Fragmenter
    CurrentZZ = Vert1.Zz: ZZShift = (Vert2.Zz - Vert1.Zz) * Fragmenter
    CurrentTx = Vert1.TexXX: TXShift = (Vert2.TexXX - Vert1.TexXX) * Fragmenter
    CurrentTy = Vert1.TexYY: TYShift = (Vert2.TexYY - Vert1.TexYY) * Fragmenter

    For HoriLine = Vert1.y To Vert2.y Step Sgn(Vert2.y - Vert1.y)
        If HoriLine > 0 And HoriLine < ScaleHeight Then
            If LBuffer(HoriLine).StartEntered = 0 Then
                'If the StartEntered variable is false, then the details of the start of the line have not yet
                'been filled in, so the current variables and stored in the line buffer
                LBuffer(HoriLine).StartX = CurrentX
                LBuffer(HoriLine).XXStart = CurrentXX
                LBuffer(HoriLine).YYStart = CurrentYY
                LBuffer(HoriLine).ZZStart = CurrentZZ
                LBuffer(HoriLine).XXStartTx = CurrentTx
                LBuffer(HoriLine).YYStartTx = CurrentTy
                LBuffer(HoriLine).StartEntered = 1
                
            ElseIf LBuffer(HoriLine).StartEntered = 1 And Int(LBuffer(HoriLine).XXStart) <> Int(CurrentXX) Then
                'If the StartEntered variable is true, then the details of the start of the line are already
                'there, so theend details are filledin, and the line is scanned
                LBuffer(HoriLine).EndX = CurrentX
                LBuffer(HoriLine).XXEnd = CurrentXX
                LBuffer(HoriLine).YYEnd = CurrentYY
                LBuffer(HoriLine).ZZEnd = CurrentZZ
                LBuffer(HoriLine).XXEndTx = CurrentTx
                LBuffer(HoriLine).YYEndTx = CurrentTy
                If ShapeFX <> 6 Or MouseY = HoriLine Then ScanLine LBuffer(HoriLine), HoriLine
                LBuffer(HoriLine).StartX = CurrentX
            End If
        End If
        CurrentX = CurrentX + XShift
        CurrentY = CurrentY + YShift
        CurrentXX = CurrentXX + XXShift
        CurrentYY = CurrentYY + YYShift
        CurrentZZ = CurrentZZ + ZZShift
        CurrentTx = CurrentTx + TXShift
        CurrentTy = CurrentTy + TYShift
    Next HoriLine
End Sub



Private Sub ScanLine(LineDetail As typLine, y)
    'This is the last level of the rendering process. When the start and end points of a scan line have been
    'found, that line can be drawn. It also has to blend the values from the start of the scan line into
    'those at the end of the scan line as it moves from one to the other.
    On Error Resume Next
    Dim Colour As Long
    Dim DifX As Integer, DifY As Integer
    Dim X As Integer, Devider As Single
    Dim CurrentXX As Single, XXShift As Single
    Dim CurrentYY As Single, YYShift As Single
    Dim CurrentZZ As Single, ZZShift As Single
    Dim CurrentTx As Single, XXShiftTx As Single
    Dim CurrentTy As Single, YYShiftTx As Single
    Dim Red As Integer, Blue As Integer, Green As Integer
    
    If LineDetail.EndX = LineDetail.StartX Then Exit Sub
    Devider = 1 / Abs(LineDetail.EndX - LineDetail.StartX)
    CurrentXX = LineDetail.XXStart: XXShift = (LineDetail.XXEnd - LineDetail.XXStart) * Devider
    CurrentYY = LineDetail.YYStart: YYShift = (LineDetail.YYEnd - LineDetail.YYStart) * Devider
    CurrentZZ = LineDetail.ZZStart: ZZShift = (LineDetail.ZZEnd - LineDetail.ZZStart) * Devider
    CurrentTx = LineDetail.XXStartTx: XXShiftTx = (LineDetail.XXEndTx - LineDetail.XXStartTx) * Devider
    CurrentTy = LineDetail.YYStartTx: YYShiftTx = (LineDetail.YYEndTx - LineDetail.YYStartTx) * Devider
    
    
    Dim Rhdc As Long
    Rhdc = frmMain.ActiveForm.TexMap.GetCurrentHDC

        For X = LineDetail.StartX + 1 To LineDetail.EndX + 1 Step Sgn(LineDetail.EndX - LineDetail.StartX)
            If X > 0 And X < ScaleWidth Then
                If DrawingShadow = 0 Then
                

                
                    Select Case ShapeFX
                        Case 1
                            '-----------------------------------------------------------------
                            '2D checker board with 3D Depth Shading
                            If Abs(X) Mod 30 >= 15 Then
                                If Abs(y) Mod 30 >= 15 Then Colour = RGB(100 + (CurrentZZ * 0.5), 0, 0) Else Colour = RGB(0, 100 + (CurrentZZ * 0.5), 0)
                            Else
                                If Abs(y) Mod 30 >= 15 Then Colour = RGB(0, 100 + (CurrentZZ * 0.5), 0) Else Colour = RGB(100 + (CurrentZZ * 0.5), 0, 0)
                            End If
                        
                        
                        Case 2
                            '-----------------------------------------------------------------
                            'Bitmaps distorted onto faces
                            'Red = Tx(CurrentTx, CurrentTy, 1) + (CurrentZZ * 0.5) '+ 100
                            'Green = Tx(CurrentTx, CurrentTy, 2) + (CurrentZZ * 0.5) '+ 100
                            'Blue = Tx(CurrentTx, CurrentTy, 3) + (CurrentZZ * 0.5) ' + 100
                            Colour = GetPixel(Rhdc, CurrentTy, CurrentTx)
                            
    Red = Colour And 255
    Green = (Colour And 65280) / 256
    Blue = (Colour And 16711680) / 65536
                            
                            
    Red = -100 + Red + (CurrentZZ * 0.5) + (Rnd * lGrain)
    Green = -100 + Green + (CurrentZZ * 0.5) + (Rnd * lGrain)
    Blue = -100 + Blue + (CurrentZZ * 0.5) + (Rnd * lGrain)
    

    
    If Red < 0 Then Red = 0
    If Green < 0 Then Green = 0
    If Blue < 0 Then Blue = 0
    
    Colour = RGB(Red, Green, Blue)
        
                        
                        
                        Case 3
                            '-----------------------------------------------------------------
                            '3D checkerboard with depth shading
                            If Abs(CurrentXX + 10000) Mod 30 >= 15 Then
                                If Abs(CurrentYY + 10000) Mod 30 >= 15 Then Colour = RGB(100 + (CurrentZZ * 0.5), 0, 0) Else Colour = RGB(0, 100 + (CurrentZZ * 0.5), 0)
                            Else
                                If Abs(CurrentYY + 10000) Mod 30 >= 15 Then Colour = RGB(0, 100 + (CurrentZZ * 0.5), 0) Else Colour = RGB(100 + (CurrentZZ * 0.5), 0, 0)
                            End If
                        
                        
                        Case 4
                            '-----------------------------------------------------------------
                            '3D checkerboard with depth shading
                            If Abs(CurrentXX + 10000) Mod 30 >= 15 Then
                                If Abs(CurrentZZ + 10000) Mod 30 >= 15 Then Colour = RGB(100 + (CurrentZZ * 0.5), 0, 0) Else Colour = RGB(0, 100 + (CurrentZZ * 0.5), 0)
                            Else
                                If Abs(CurrentZZ + 10000) Mod 30 >= 15 Then Colour = RGB(0, 100 + (CurrentZZ * 0.5), 0) Else Colour = RGB(100 + (CurrentZZ * 0.5), 0, 0)
                            End If
                        
                        
                        Case 5
                            '-----------------------------------------------------------------
                            'Solid colour, with depth shading
                            Red = -100 + lBrushRed + (CurrentZZ * 0.5) + (Rnd * lGrain)
                            Green = -100 + lBrushGreen + (CurrentZZ * 0.5) + (Rnd * lGrain)
                            Blue = -100 + lBrushBlue + (CurrentZZ * 0.5) + (Rnd * lGrain)
                            

                            
                            If Red < 0 Then Red = 0
                            If Green < 0 Then Green = 0
                            If Blue < 0 Then Blue = 0
                            
                            Colour = RGB(Red, Green, Blue)
                            
                        
                        Case 6
                            If X = MouseX And y = MouseY Then
                                'Red = PaintColour And 255
                                'Green = (PaintColour And (256 ^ 2 - 256)) / 256
                                'Blue = (PaintColour And (256 ^ 3 - 65536)) / (256 ^ 2)
                                frmMain.ActiveForm.TexMap.DoPset CurrentTx, CurrentTy, PaintColour
                                'SetPixel Rhdc, CurrentTx, CurrentTy, PaintColour
                                'Tx(CurrentTx, CurrentTy, 1) = Red
                                'Tx(CurrentTx, CurrentTy, 2) = Green
                                'Tx(CurrentTx, CurrentTy, 3) = Blue
                                Exit Sub
                            End If
                    End Select
                End If
    
                
                DifX = (Rnd * lDiffuse)
                DifY = (Rnd * lDiffuse)
                If ZBuffer(y + DifY, X + DifX) <= CurrentZZ + 5000 Then
                    ZBuffer(y + DifY, X + DifX) = CurrentZZ + 5000
'                    If PicketAnimate = 0 Then
                        SetPixel hdc, X + DifX, y + DifY, Colour
'                    Else
'                        If Int((x + DifX) * 0.25) Mod PicketAnimate = PicketAnimateOver Then
'                            SetPixel hdc, x + DifX, y + DifY, Colour
'                        End If
'                    End If
                End If
               
               
               
                CurrentXX = CurrentXX + XXShift
                CurrentYY = CurrentYY + YYShift
                CurrentZZ = CurrentZZ + ZZShift
                CurrentTx = CurrentTx + XXShiftTx
                CurrentTy = CurrentTy + YYShiftTx
            End If
        Next X
    

End Sub

Private Function CombineColour(ColourA As Long, ColourB As Long, PercentOfA As Single) As Long

    Dim RedA As Byte, GreenA As Byte, BlueA As Byte
    Dim RedB As Byte, GreenB As Byte, BlueB As Byte
    Dim RedC As Byte, GreenC As Byte, BlueC As Byte

    RedA = ColourA And 255
    GreenA = (ColourA And 65280) / 256
    BlueA = (ColourA And 16711680) / 65536

    RedB = ColourB And 255
    GreenB = (ColourB And 65280) / 256
    BlueB = (ColourB And 16711680) / 65536
    
    RedC = (RedB * PercentOfA) + (RedA * (1 - PercentOfA))
    BlueC = (BlueB * PercentOfA) + (BlueA * (1 - PercentOfA))
    GreenC = (GreenB * PercentOfA) + (GreenA * (1 - PercentOfA))

    CombineColour = RGB(RedC, GreenC, BlueC)

End Function


Private Sub DrawGuides()
    'This draws the origin mark, which points out which axis is which.
    'Its kind of an alternative to the ground plane. Its split in two parts
    'to make the code nicer. This part just calls the next part...
    Draw3DLine 70, 0, 0, 0, 0, 0, "X", RGB(255, 0, 0)
    Draw3DLine 0, 70, 0, 0, 0, 0, "Y", RGB(0, 255, 0)
    Draw3DLine 0, 0, 70, 0, 0, 0, "Z", RGB(0, 0, 255)
End Sub


Private Sub Draw3DLine(x1 As Integer, y1 As Integer, z1 As Integer, x2 As Integer, y2 As Integer, z2 As Integer, Message As String, col As Long)
    'This is the second part, which rotates the values given in the first half,
    'and draws a line from that point to the center, and writes a message on
    'to the screen at the end of the line.
    Dim Xx1 As Integer, Yy1 As Integer, Xx2 As Integer, Yy2 As Integer
    Dim Rotated As clsVertex, zeye As Integer
    zeye = 800:    ForeColor = col:    Set Rotated = New clsVertex
    Rotated.X = x1:    Rotated.y = y1:    Rotated.z = z1
    Set Rotated = RotatePoint(Rotated, Angle1, Angle2, Angle3)
    If pPerspecitve = True Then
        Xx1 = Xf + Int(Rotated.X * (zeye / (zeye - Rotated.z)))
        Yy1 = Yf + Int(Rotated.y * (zeye / (zeye - Rotated.z)))
    Else
        Xx1 = Xf + Int(Rotated.X)
        Yy1 = Yf + Int(Rotated.y)
    End If
    Rotated.X = x2:    Rotated.y = y2:    Rotated.z = z2
    Set Rotated = RotatePoint(Rotated, Angle1, Angle2, Angle3)
    If pPerspecitve = True Then
        Xx2 = Xf + Int(Rotated.X * (zeye / (zeye - Rotated.z)))
        Yy2 = Yf + Int(Rotated.y * (zeye / (zeye - Rotated.z)))
    Else
        Xx2 = Xf + Int(Rotated.X)
        Yy2 = Yf + Int(Rotated.y)
    End If
    DrawLine Xx1, Yy1, Xx2, Yy2
    CurrentX = Xx1
    CurrentY = Yy1
    Print Message
    ForeColor = 0
End Sub


Private Sub UserControl_Resize()
    'This sets the Xf and Yf variables point to the centre of
    'the screen, so that the model is centered correctly
    Xf = ScaleWidth / 2
    Yf = ScaleHeight / 2
    ReDim ZBuffer(ScaleHeight, ScaleWidth) As Long
    ReDim SBuffer(ScaleHeight, ScaleWidth) As Integer
    ReDim TBuffer(ScaleHeight, ScaleWidth) As Byte
    ReDim IBuffer(ScaleHeight, ScaleWidth) As Long
End Sub
