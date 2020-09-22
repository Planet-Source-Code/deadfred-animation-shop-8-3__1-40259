VERSION 5.00
Begin VB.UserControl DXEngine 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.Timer tmaLightStyle 
      Interval        =   100
      Left            =   1440
      Top             =   1080
   End
End
Attribute VB_Name = "DXEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'#####################################################################
'#                                                                   #
'#  This control displays a file object in 3D using Direct X to      #
'#  speed up the rendering process. You'll need DirectX setup        #
'#  on your PC, so windowsNT users may not be able to use it         #
'#                                                                   #
'#####################################################################

'Declare objects
Dim g_dx As New DirectX7
Dim m_dd As DirectDraw7
Dim m_ddClipper As DirectDrawClipper
Dim m_rm As Direct3DRM3
'Declare devices
Dim m_rmDevice As Direct3DRMDevice3
'declare viewports
Dim m_rmViewport As Direct3DRMViewport2
'Declare frames
Dim m_rootFrame As Direct3DRMFrame3
Dim m_lightFrame As Direct3DRMFrame3
Dim m_cameraFrame As Direct3DRMFrame3
Dim m_objectFrame As Direct3DRMFrame3
'Declare meshes
Dim m_meshBuilder As Direct3DRMMeshBuilder3
'Delare lights
Dim m_light As Direct3DRMLight
Dim m_ambientLight As Direct3DRMLight
'Viewport sizes
Dim m_width As Long
Dim m_height As Long
'Mousedown positions for rotation
Public m_LastX As Integer
Public m_LastY As Integer
Dim VertArray() As D3DVECTOR 'Vertices for the object
Dim SideFaces() As Long      'Vertices making the sides of each face

Private Model As clsFile
Public pLightPattern As String
Public pSelectedOnly As Boolean
Public pCenterModel As Boolean
Public pShowSkeliton As Boolean

Public Property Get FileKey() As String
    FileKey = Model.Key
End Property

Public Sub ClearWindow()
    If InitRM Then
        InitScene
        RenderScene
        Am8.pShowNoDX = True
    Else
        If DirectXNotAvaliable = False And Am8.pShowNoDX = False Then
            If MsgBox(amShowNoDirectX, vbInformation + vbYesNo) = vbNo Then Am8.pShowNoDX = True
        End If
        DirectXNotAvaliable = True
    End If
End Sub

Public Sub AssignDXEngineTo(AssignedModel As clsFile)
    If AssignedModel Is Nothing Then Else Set Model = AssignedModel
End Sub

Public Sub RefreshModel()
    PlaceModelInWindow
End Sub

Public Sub RenderScene()
    'This calls the viewport render methods, and draws the image to the screen.
    On Local Error Resume Next
    m_rmViewport.Clear D3DRMCLEAR_ALL
    m_rmViewport.Render m_rootFrame
    m_rmDevice.Update
End Sub

Public Sub SetLights(Ambiant, Spot)
    'This sets the ambiant and spot light levels. It takes a value between 1
    'and 100 for each. It seems to be able to take coloured light values, but
    'it doesn't have any effect. I've probebly done somthing wrong...
    If m_light Is Nothing Then Exit Sub
    m_light.SetColorRGB Spot / 100, Spot / 100, Spot / 100
    m_ambientLight.SetColorRGB Ambiant / 100, Ambiant / 100, Ambiant / 100
    RenderScene
End Sub

Public Sub SetMode(Mode As Integer)
    'This changes the render method of the device rmDevice, where all the
    'objects are. The higher Mode is, the better quality method used..
    Select Case Mode
        Case 1: m_rmDevice.SetQuality D3DRMFILL_POINTS
        Case 2: m_rmDevice.SetQuality D3DRMRENDER_WIREFRAME
        Case 3: m_rmDevice.SetQuality D3DRMRENDER_FLAT
        Case 4: m_rmDevice.SetQuality D3DRMRENDER_PHONG
    End Select
    RenderScene
End Sub

Public Function InitRM() As Boolean
    'This sets up a drawing plane in a picture box. If this fails then
    'its most likly that DirectX isn't installed, so you can't use this
    'feature at all. NT dosn't have DX as standard.
    On Error GoTo FailedDX_Must_Be_An_NT_User
    Set m_dd = g_dx.DirectDrawCreate("")
    Set m_ddClipper = m_dd.CreateClipper(0)
    m_ddClipper.SetHWnd hWnd
    m_width = ScaleWidth
    m_height = ScaleHeight
    Set m_rm = g_dx.Direct3DRMCreate()
    Set m_rmDevice = m_rm.CreateDeviceFromClipper(m_ddClipper, "", m_width, m_height)
    SetMode frmMain.sldShade.ListIndex + 1
    InitRM = True
FailedDX_Must_Be_An_NT_User:
End Function

Public Sub PlaceModelInWindow()
    Dim Am As clsObject, Vm As clsVertex, Fm As clsFace, Em As clsEdge, Rotated As clsVertex
    Dim VertArray() As D3DVECTOR, SideFaces() As Long, VertexOn As Integer, FaceON As Integer
    Dim C1 As Long, Green As Byte, Blue As Byte, Red As Byte, Jm As clsJoint
    Dim XX As Single, YY As Single, Zz As Single
    If Model Is Nothing Then Exit Sub
    If Not InitRM Then Enabled = False: DirectXNotAvaliable = True: Exit Sub
    InitScene
    Set Rotated = New clsVertex
    If pCenterModel = True Then
        Model.SelectAll
        Model.FindModelOutline
        Model.DeselectAll
        XX = (Model.MinX + Model.MaxX) * 0.7
        YY = (Model.MinY + Model.MaxY) * 0.7
        Zz = (Model.MinZ + Model.MaxZ) * 0.7
    End If
    For Each Am In Model.Geometery
        If Am.Selected = True Or pSelectedOnly = False Then
            ReDim VertArray(Am.Vertex.Count) As D3DVECTOR
            ReDim SideFaces(Am.EdgeFaceCount) As Long
            VertexOn = 0
            For Each Vm In Am.Vertex
                If Vm.TargetName <> "" Then
                    With Model.Joint(Vm.TargetName)
                        Set Rotated = RotatePoint(Vm, 0, 0, 0, 0, 0, 0)
                        Rotated.x = Rotated.x - (.x - .NewPositX)
                        Rotated.y = Rotated.y - (.y - .NewPositY)
                        Rotated.z = Rotated.z - (.z - .NewPositZ)
                        Set Rotated = RotatePoint(Rotated, .AngleX, .AngleY, .AngleZ, .NewPositX, .NewPositY, .NewPositZ)
                    End With
                Else
                    Set Rotated = RotatePoint(Vm, 0, 0, 0)
                End If
                VertArray(VertexOn).x = (XX + Rotated.x) * 0.16
                VertArray(VertexOn).y = (YY + -Rotated.y) * 0.16
                VertArray(VertexOn).z = (Zz + Rotated.z) * 0.16
                VertexOn = VertexOn + 1
            Next Vm
            FaceON = 0
            For Each Fm In Am.Face
                SideFaces(FaceON) = Fm.EdgeCount
                FaceON = FaceON + 1
                For Each Em In Fm.Edge
                    SideFaces(FaceON) = Em.Vertex - 1
                    FaceON = FaceON + 1
                Next Em
            Next Fm
            C1 = Am.Colour
            Red = C1 And 255
            Green = (C1 And (256 ^ 2 - 256)) / 256
            Blue = (C1 And (256 ^ 3 - 65536)) / (256 ^ 2)
            If Red = 0 And Blue = 0 And Green = 0 Then Red = 255: Green = 255: Blue = 255
            AddShapeDX Am.Vertex.Count, VertArray, SideFaces, Red, Green, Blue
        End If
    Next Am
    If pShowSkeliton = True Then
        For Each Jm In Model.Joint
            If Jm.Target <> "" Then
                ReDim VertArray(6) As D3DVECTOR
                VertArray(0).x = Jm.x: VertArray(0).y = Jm.y: VertArray(0).z = Jm.z
                VertArray(1).x = Jm.x + JointWid: VertArray(1).y = Jm.y + JointWid: VertArray(1).z = Jm.z + JointWid
                VertArray(2).x = Jm.x - JointWid: VertArray(2).y = Jm.y + JointWid: VertArray(2).z = Jm.z - JointWid
                With Model
                    VertArray(3).x = .Joint(Jm.Target).x: VertArray(3).y = .Joint(Jm.Target).y: VertArray(3).z = .Joint(Jm.Target).z
                    VertArray(4).x = .Joint(Jm.Target).x + JointWid: VertArray(4).y = .Joint(Jm.Target).y + JointWid: VertArray(4).z = .Joint(Jm.Target).z + JointWid
                    VertArray(5).x = .Joint(Jm.Target).x - JointWid: VertArray(5).y = .Joint(Jm.Target).y + JointWid: VertArray(5).z = .Joint(Jm.Target).z - JointWid
                End With
                For VertexOn = 0 To 5
                    VertArray(VertexOn).x = VertArray(VertexOn).x * 0.16
                    VertArray(VertexOn).y = VertArray(VertexOn).y * -0.16
                    VertArray(VertexOn).z = VertArray(VertexOn).z * 0.16
                Next VertexOn
                ReDim SideFaces(15) As Long
                SideFaces(0) = 4:   SideFaces(1) = 0:  SideFaces(2) = 1:  SideFaces(3) = 4:  SideFaces(4) = 3
                SideFaces(5) = 4:   SideFaces(6) = 1:  SideFaces(7) = 2:  SideFaces(8) = 5:  SideFaces(9) = 4
                SideFaces(10) = 4: SideFaces(11) = 2: SideFaces(12) = 0: SideFaces(13) = 3: SideFaces(14) = 5
                AddShapeDX 6, VertArray, SideFaces, 0, 0, 255
            End If
        Next Jm
    End If
    RenderScene
End Sub

Private Sub tmaLightStyle_Timer()
    Static LightPosition As Integer, LightValue As Integer
    If pLightPattern <> "" Then
        LightPosition = LightPosition + 1
        If LightPosition > Len(pLightPattern) Then LightPosition = 1
        LightValue = 100 - (Asc(Mid(pLightPattern, LightPosition, 1)) - 65) * 4
        SetLights frmMain.sldLight(0), LightValue
        RenderScene
    End If
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    Select Case Chr$(KeyAscii)
        Case "8": MovePlayerForward (1)
        Case "5": MovePlayerForward (-1)
        Case "4": RotatePlayer (-4 / Pie)
        Case "6": RotatePlayer (4 / Pie)
        Case "1": MovePlayerUp (1)
        Case "0": MovePlayerUp (-1)
    End Select
    RenderScene
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    m_LastX = x
    m_LastY = y
End Sub

Public Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 0 Then
        RotateTrackBall CInt(x), CInt(y)
        RenderScene
        DoEvents
    End If
End Sub

Private Sub UserControl_Paint()
    RenderScene
End Sub

Public Sub InitScene()
    'This sets up the Direct 3D frames and lights and such. If the Init_RM
    'command failed, it dosn't run this bit, because this bit won't run either
    Set m_rootFrame = m_rm.CreateFrame(Nothing)         'All objects and lights are inside
    Set m_cameraFrame = m_rm.CreateFrame(m_rootFrame)   'frames. When a frame moves, everything
    Set m_lightFrame = m_rm.CreateFrame(m_rootFrame)    'in the frame also moves. These frame
    Set m_objectFrame = m_rm.CreateFrame(m_rootFrame)   'hold the object and lights
    m_cameraFrame.SetPosition Nothing, 0, 0, -100 'Set the camera start position
    Set m_rmViewport = m_rm.CreateViewport(m_rmDevice, m_cameraFrame, 0, 0, m_width, m_height)
    m_rmViewport.SetBack (1000) 'Set the back cliper. The bigger it is, the farther you can see
    Set m_light = m_rm.CreateLight(D3DRMLIGHT_DIRECTIONAL, &HFFFFFFFF) 'Sets the spot light
    m_lightFrame.AddLight m_light 'Adds the spot light to the frame
    Set m_ambientLight = m_rm.CreateLightRGB(D3DRMLIGHT_AMBIENT, 0.5, 0.5, 0.5) 'Sets the ambiant light
    m_lightFrame.AddLight m_ambientLight 'Adds the ambiant (background) light to te frame
End Sub

Public Sub RotateTrackBall(x As Integer, y As Integer)
    'this function taken from MS engine sample in the VB sample that
    'comes with MS DirectX7 SDK. It works as follows:
    'select point on screen interpret as though selecting
    'point on sphere. as new point is passed when mouse is
    'moved rotate in coresponding direction(s) on sphere.
    If Model Is Nothing Then Exit Sub
    Dim delta_x As Single, delta_y As Single
    Dim delta_r As Single, radius As Single, denom As Single, Angle As Single
    ' rotation axis in camcoords, worldcoords, sframecoords
    Dim axisC As D3DVECTOR
    Dim wc As D3DVECTOR
    Dim axisS As D3DVECTOR
    Dim base As D3DVECTOR
    Dim Origin As D3DVECTOR
    delta_x = x - m_LastX
    delta_y = y - m_LastY
    m_LastX = x
    m_LastY = y
    delta_r = Sqr(delta_x * delta_x + delta_y * delta_y)
    radius = 50
    denom = Sqr(radius * radius + delta_r * delta_r)
    If (delta_r = 0 Or denom = 0) Then Exit Sub
    Angle = (delta_r / denom)
    axisC.x = (-delta_y / delta_r)
    axisC.y = (-delta_x / delta_r)
    axisC.z = 0
    m_cameraFrame.Transform wc, axisC
    m_objectFrame.InverseTransform axisS, wc
    m_cameraFrame.Transform wc, Origin
    m_objectFrame.InverseTransform base, wc
    axisS.x = axisS.x - base.x
    axisS.y = axisS.y - base.y
    axisS.z = axisS.z - base.z
    m_objectFrame.AddRotation D3DRMCOMBINE_BEFORE, axisS.x, axisS.y, axisS.z, Angle
End Sub

Public Sub AddShapeDX(NumVerts As Integer, VertArray() As D3DVECTOR, SideFaces() As Long, Red As Byte, Green As Byte, Blue As Byte)
    'This takes a list of vertecies and edges, and puts them into the
    'meshbuilder, which then adds it to the frame.
    Set m_meshBuilder = m_rm.CreateMeshBuilder()
    Dim NormArray(0) As D3DVECTOR
    m_meshBuilder.AddFaces NumVerts, VertArray, 0, NormArray, SideFaces
    m_meshBuilder.SetColorRGB Red / 255, Green / 255, Blue / 255
    m_objectFrame.AddVisual m_meshBuilder
End Sub

Public Sub RotatePlayer(Angle)
    'This rotates the player left or right by the selected amount
    m_cameraFrame.AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, Angle
End Sub

Public Sub TiltPlayer(Angle)
    'This tilts the player to look up or down by a set number of degrees
    m_cameraFrame.AddRotation D3DRMCOMBINE_BEFORE, 0, 0, 1, Angle
End Sub

Public Sub MovePlayerForward(Distance)
    'This moves the player forawrd, in the direction they are looking
    'by a set value. Negatives make you go backwards
    m_cameraFrame.AddTranslation D3DRMCOMBINE_BEFORE, 0, 0, Distance
End Sub

Public Sub MovePlayerUp(Distance)
    'This moves the camara up or down by a set value
    m_cameraFrame.AddTranslation D3DRMCOMBINE_BEFORE, 0, Distance * 0.5, 0
End Sub

Private Sub UserControl_Terminate()
    'When messing with DX, its best to clean up after yourself, or things
    'stop working :o(
    Set m_light = Nothing
    Set m_ambientLight = Nothing
    Set m_meshBuilder = Nothing
    Set m_rmViewport = Nothing
    Set m_lightFrame = Nothing
    Set m_cameraFrame = Nothing
    Set m_objectFrame = Nothing
    Set m_rootFrame = Nothing
    Set m_rmDevice = Nothing
    Set m_ddClipper = Nothing
    Set m_rm = Nothing
    Set m_dd = Nothing
End Sub
