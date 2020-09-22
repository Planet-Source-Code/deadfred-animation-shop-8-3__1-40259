VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public Key As String
Public Saved As Boolean
Public ModelName As String
Public NonEditableFile As Boolean

Public Geometery As colObject
Public Joint As colJoint
Public Scene As colScene
Public Layers As colLayer

Public CurrentFilePath As String
Public ModelDate As Date
Public MinX As Single, MaxX As Single
Public MinY As Single, MaxY As Single
Public MinZ As Single, MaxZ As Single
Public Notes As String
Public DisplayJoints As Boolean

Public Wraper As colWrap


Private Sub Class_Initialize()
    Set Geometery = New colObject
    Set Joint = New colJoint
    Set Scene = New colScene
    Set Layers = New colLayer
    Set Wraper = New colWrap
End Sub

Public Sub LockSelected()
    Dim Am As clsObject
    For Each Am In Geometery: If Am.Selected = True Then Am.Locked = True
    Next Am
End Sub

Public Sub HideSelected()
    Dim Am As clsObject
    For Each Am In Geometery: If Am.Selected = True Then Am.Hidden = True
    Next Am
End Sub

Public Sub GreySelected()
    Dim Am As clsObject
    For Each Am In Geometery: If Am.Selected = True Then Am.Grayed = True
    Next Am
End Sub

Public Sub UnLockAll()
    Dim Am As clsObject
    For Each Am In Geometery: If Am.Hidden = False Then Am.Locked = False
    Next Am
End Sub

Public Sub UnGreyAll()
    Dim Am As clsObject
    For Each Am In Geometery: Am.Grayed = False: Next Am
End Sub

Public Sub UnHideAll()
    Dim Am As clsObject
    For Each Am In Geometery: If Am.Hidden = True Then Am.Hidden = False: Am.Locked = False
    Next Am
End Sub


Public Sub RotateSelection(Angle As Single, ViewMode As Integer, CenterMode As Integer)
    Dim Cx As Single, Cy As Single, Cz As Single
    Dim Am As clsObject, Jm As clsJoint
    Saved = False
    If CenterMode = 0 Then
        Cx = (MinX + MaxX) * 0.5
        Cy = (MinY + MaxY) * 0.5
        Cz = (MinZ + MaxZ) * 0.5
    End If
    If CenterMode = 2 Then
    End If
    For Each Am In Geometery
        If Am.Selected = True Then
            If CenterMode = 1 Then
                Cx = (Am.MinX + Am.MaxX) * 0.5
                Cy = (Am.MinY + Am.MaxY) * 0.5
                Cz = (Am.MinZ + Am.MaxZ) * 0.5
            End If
            Select Case ViewMode
                Case 1: Am.Rotate 0, Angle * 10, 0, Cx, Cy, Cz
                Case 3: Am.Rotate -Angle * 10, 0, 0, Cx, Cy, Cz
                Case 2: Am.Rotate 0, 0, Angle * 10, Cx, Cy, Cz
            End Select
        End If
    Next Am
    For Each Jm In Joint
        If Jm.Selected = True Then
            Select Case ViewMode
                Case 1: Jm.Rotate 0, Angle * 10, 0, Cx, Cy, Cz
                Case 3: Jm.Rotate -Angle * 10, 0, 0, Cx, Cy, Cz
                Case 2: Jm.Rotate 0, 0, Angle * 10, Cx, Cy, Cz
            End Select
        End If
    Next Jm
End Sub


Public Sub DuplicateSelection(ViewMode As Integer)
    Saved = False
    SaveToFile App.Path & "\data\duplicate", 1
    Select Case ViewMode
        Case 1:  LoadFromFile App.Path & "\data\duplicate", 1, 10, , 10
        Case 2:  LoadFromFile App.Path & "\data\duplicate", 1, 10, 10
        Case 3:  LoadFromFile App.Path & "\data\duplicate", 1, , 10, 10
    End Select
End Sub


Public Function DeleteSelected()
    'This removes all the selected objects from the geometery collection
    Dim Am As clsObject, Jm As clsJoint, Jmm As clsJoint, Vm As clsVertex
    For Each Am In Geometery
        If Am.Selected = True Then
            Geometery.RemoveObject Am.Key
        End If
    Next Am
    For Each Jm In Joint
        If Jm.Selected = True Then
            For Each Jmm In Joint
                If Jmm.Target = Jm.Key Then Jmm.Target = ""
            Next Jmm
            For Each Am In Geometery
                For Each Vm In Am.Vertex
                    If Vm.TargetName = Jm.Key Then Vm.TargetName = ""
                Next Vm
            Next Am
            Joint.Remove Jm.Key
        End If
    Next Jm
    FindModelOutline
    Joint.DisplayTreeInWindow frmMain.Joints
End Function


Public Sub FindModelOutline()
    'This finds the size of the selected objects bounding box; Ie. the smallest posible that
    'can suround all of the selected objects
    MinX = 0: MinY = 0: MinZ = 0: MaxX = 0: MaxY = 0: MaxZ = 0
    Dim Am As clsObject, FirstObject As Integer, Jm As clsJoint, FirstJoint As Integer
    FirstObject = Geometery.FirstSelectedObject
    FirstJoint = Joint.FirstSelectedJoint
    If FirstObject <> 0 Then
        MinX = Geometery(FirstObject).MinX: MaxX = Geometery(FirstObject).MaxX
        MinY = Geometery(FirstObject).MinY: MaxY = Geometery(FirstObject).MaxY
        MinZ = Geometery(FirstObject).MinZ: MaxZ = Geometery(FirstObject).MaxZ
        For Each Am In Geometery
            If Am.Selected = True Then
                If Am.MaxX > MaxX Then MaxX = Am.MaxX
                If Am.MaxY > MaxY Then MaxY = Am.MaxY
                If Am.MaxZ > MaxZ Then MaxZ = Am.MaxZ
                If Am.MinX < MinX Then MinX = Am.MinX
                If Am.MinY < MinY Then MinY = Am.MinY
                If Am.MinZ < MinZ Then MinZ = Am.MinZ
            End If
        Next Am
    End If
    If FirstObject = 0 And FirstJoint <> 0 Then
        MinX = Joint(FirstJoint).X
        MinY = Joint(FirstJoint).y
        MinZ = Joint(FirstJoint).z
        MaxX = Joint(FirstJoint).X
        MaxY = Joint(FirstJoint).y
        MaxZ = Joint(FirstJoint).z
    End If
    For Each Jm In Joint
        If Jm.Selected = True Then
            If Jm.X > MaxX Then MaxX = Jm.X
            If Jm.y > MaxY Then MaxY = Jm.y
            If Jm.z > MaxZ Then MaxZ = Jm.z
            If Jm.X < MinX Then MinX = Jm.X
            If Jm.y < MinY Then MinY = Jm.y
            If Jm.z < MinZ Then MinZ = Jm.z
        End If
    Next Jm
    If FirstJoint <> 0 Then
        If MinX = MaxX Then MinX = MinX - 4: MaxX = MaxX + 4
        If MinY = MaxY Then MinY = MinY - 4: MaxY = MaxY + 4
        If MinZ = MaxZ Then MinZ = MinZ - 4: MaxZ = MaxZ + 4
    End If
End Sub

Public Sub DeselectAll()
    Dim Am As clsObject, Jm As clsJoint
    For Each Am In Geometery: Am.Selected = False: Next Am
    For Each Jm In Joint: Jm.Selected = False: Next Jm
    If ActiveFile <> "" Then Am8(ActiveFile).FindModelOutline
End Sub

Public Sub SelectAll()
    Dim Am As clsObject, Jm As clsJoint
    For Each Am In Geometery: Am.Selected = True: Next Am
    For Each Jm In Joint: Jm.Selected = True: Next Jm
    FindModelOutline
End Sub

Public Function SaveToFile(FileName As String, Optional Mode As Integer = 0, Optional XShift As Single = 0, Optional YShift As Single = 0, Optional ZShift As Single = 0)
    'This saves the file object to the file name given. This also does the
    'copy and paste function, so that both Save and Copy use the same code and
    'have the same format. When the mode is 1 (for Copy), certain peices of
    'data are not included in the save, but apart from that, Save and Copy are
    'the same. Obviously, only selected objects and joints are included in the
    'Copy function
    Dim Am As clsObject, Vm As clsVertex, Fm As clsFace, Em As clsEdge, Gm As Integer
    Dim Jm As clsJoint, Sm As clsScene, Dm As clsFrame, Rm As clsJointRow
    Dim Lm As clsLayer, Tm As clsTexVertex, n As Integer
    On Error GoTo FailedToSave
    Open FileName For Output As #1
    If Mode = 0 Then
        'When saving an object, this data is included, but it is not included
        'when copying data to the clipboard
        Print #1, "MD_NT:"; 1
        Print #1, Notes
        Print #1, "MD_NT:"; 1
        CurrentFilePath = FileName
        Print #1, "MD_FP: "; CurrentFilePath
        For Each Lm In Layers
            Print #1, "LY_LK: "; Lm.LayerKey
            Print #1, "LY_LN: "; Lm.LayerName
            Print #1, "LY_CL: "; Lm.LayerColour
            Print #1, "LY_GR: "; Lm.LayerGrayed
            Print #1, "LY_HD: "; Lm.LayerHidden
            Print #1, "LY_LL: "; Lm.LayerLocked
            Print #1, "LY_SL: "; Lm.Selected
        Next Lm
        Saved = True
    End If
    If Mode = 0 Then Print #1, "JTCNT: "; Joint.CountChildren Else Print #1, "JTCNT: "; Joint.CountSelected
    If Mode = 0 Then Print #1, "OBCNT: "; Geometery.CountObjects Else Print #1, "OBCNT: "; Geometery.CountSelected
    For Each Am In Geometery
        If Mode = 0 Or Am.Selected = True Then
            Print #1, "OBNEW: "; Am.Key
            Print #1, "OB_EN: "; Am.EntityName
            Print #1, "OB_ED: "; Am.EntityDetails
            Print #1, "OB_CL: "; Am.Colour
            Print #1, "OB_DF: "; Am.Diffusion
            Print #1, "OB_TC: "; Am.Transparancy
            Print #1, "OB_TM: "; Am.TextureMode
            Print #1, "OB_GN: "; Am.grain
            Print #1, "OB_TX: "; Am.Texture
            Print #1, "OB_LY: "; Am.Layer
            Print #1, "OB_LK: "; Am.Locked
            Print #1, "OB_HD: "; Am.Hidden
            Print #1, "OB_GY: "; Am.Grayed
            For Gm = Am.Group.Count To 1 Step -1
                Print #1, "OB_GP: "; Am.Group(Gm).GroupID
            Next Gm
            For Each Tm In Am.TexVert
                Print #1, "OB_TV:"; Tm.Selected
                Print #1, Tm.X
                Print #1, Tm.y
            Next Tm
            For Each Vm In Am.Vertex
                Print #1, "OB_VT:"; Vm.Selected
                Print #1, Vm.TargetName
                Print #1, Vm.X - XShift
                Print #1, Vm.y - YShift
                Print #1, Vm.z - ZShift
            Next Vm
            For Each Fm In Am.Face
                Print #1, "OB2FC:"; Fm.EdgeCount
                For Each Em In Fm.Edge
                    Print #1, Em.Vertex
                    Print #1, Em.TexVertex
                Next Em
            Next Fm
        End If
    Next Am
    Print #1, "OBEND:"; False
    For Each Jm In Joint
        If Mode = 0 Or Jm.Selected = True Then
            Print #1, "JTNEW: "; Jm.Key
            Print #1, "JT_CL: "; Jm.Colour
            Print #1, "JT_GY: "; Jm.Grayed
            Print #1, "JT_HD: "; Jm.Hidden
            Print #1, "JT_LK: "; Jm.Locked
            Print #1, "JT_NM: "; Jm.Name
            Print #1, "JT_TG: "; Jm.Target
            Print #1, "JT_OX: "; Jm.X - XShift
            Print #1, "JT_OY: "; Jm.y - YShift
            Print #1, "JT_OZ: "; Jm.z - ZShift
            Print #1, "JT_PF: "; Jm.JointProfileIndex
            Print #1, "JT_PV: "; Jm.JProf(0)
            For n = 1 To 5
                Print #1, Jm.JProf(n)
            Next n
        End If
    Next Jm
    Print #1, "JTEND:"; False
    Print #1, "UDTJT:"; 0
    If Mode = 0 Then
        'This is data for Scenes, which dosn't get saved when you Copy data
        For Each Sm In Scene
            Print #1, "SC_KY: "; Sm.Key
            Print #1, "SC_NM: "; Sm.Name
            For Each Dm In Sm
                Print #1, "FM_KY: "; Dm.Key
                Print #1, "FM_NM: "; Dm.Name
                Print #1, "FM_SM: "; Dm.Smooth
                Print #1, "FM_SP: "; Dm.Speed
                For Each Rm In Dm
                    Print #1, "NEWRW: "; Rm.Key
                    Print #1, Rm.Name
                    Print #1, Rm.AngleX
                    Print #1, Rm.AngleY
                    Print #1, Rm.AngleZ
                    Print #1, Rm.OriginX
                    Print #1, Rm.OriginY
                    Print #1, Rm.OriginZ
                    Print #1, Rm.ScaleX
                    Print #1, Rm.ScaleY
                    Print #1, Rm.ScaleZ
                Next Rm
            Next Dm
        Next Sm
    End If
    Print #1, "AM_QT:"; 0
    Close
    If Mode = 0 Then Am8.FileHistory.AddHistory FileName
Exit Function

FailedToSave:
    MsgBox "Failed to save file" & vbNewLine & vbNewLine & Err.Description, vbCritical
End Function

Public Function LoadFromFile(FileName As String, Optional Mode As Integer = 0, Optional XShift As Single = 0, Optional YShift As Single = 0, Optional ZShift As Single = 0) As Boolean
    'This loads the file given, into the current file object. To create a new file, you first
    'have to create a new file object, and then load the filename into the file oject
    'On Error GoTo FailedToLoadFile
    Dim Funct As String, Value As String, Comand As String, ObjectKey As String
    Dim JointKey As String, SceneKey As String, FrameKey As String, Tmp As Variant, t As Integer
    Dim XX As Integer, YY As Integer, Zz As Integer, Edge As Integer, n As Integer, GetjtKEY As Integer
    Dim NoteLine As String, LayerKey As String, TexEdge As Integer, m As Integer, Vm As clsVertex
    Dim JointList() As String, ObjectList() As String, TargetName As String, Am As clsObject
    Dim GETObKEY As Integer, JointON As Integer, AddedLayer As Boolean, NewValue As String
    GETObKEY = -1: GetjtKEY = -1
    DeselectAll
    If Mode = 0 Then Layers.ClearLayers: Scene.EraseScenes
    Open FileName For Input As #1
        Do
        Line Input #1, Comand
        Funct = Mid(Comand, 1, InStr(1, Comand, ":") - 1)
        Value = Mid(Comand, InStr(1, Comand, ":") + 2, Len(Comand))
        Select Case Funct
            Case "MD_FP": CurrentFilePath = Value 'Sets the file Path
            Case "LY_LK": LayerKey = Value
            Case "LY_CL": Layers(LayerKey).LayerColour = Val(Value)
            Case "LY_GR": Layers(LayerKey).LayerGrayed = Val(Value)
            Case "LY_HD": Layers(LayerKey).LayerHidden = Value
            Case "LY_LL": Layers(LayerKey).LayerLocked = Value
            Case "LY_LN": Layers.AddLayer Value, LayerKey
            Case "LY_SL": Layers(LayerKey).Selected = Value
            Case "SC_KY": SceneKey = Value 'Store the new scene key
            Case "SC_NM": Scene.CreateScene SceneKey, Value 'Load the scene name, and create the scene with the scene key
            Case "FM_NM": Scene(SceneKey)(FrameKey).Name = Value 'load the name of the current frame
            Case "FM_SM": Scene(SceneKey)(FrameKey).Smooth = Val(Value) 'Load the frame smoothness property
            Case "FM_SP": Scene(SceneKey)(FrameKey).Speed = Val(Value) 'Load the frame Speed property
            Case "JT_CL": Joint(JointKey).Colour = Val(Value) 'Load the joint colour
            Case "JT_GY": Joint(JointKey).Grayed = Val(Value) 'Is the joint Grayed or not?
            Case "JT_HD": Joint(JointKey).Hidden = Val(Value) 'Is the joint Hidden or not?
            Case "JT_LK": Joint(JointKey).Locked = Val(Value) 'Is the joint Locked or not?
            Case "JT_NM": Joint(JointKey).Name = Value 'Load the joint name
            Case "JT_TG": Joint(JointKey).Target = Value 'Load the joint target
            Case "JT_OX": Joint(JointKey).X = Val(Value) + XShift 'Load the x position of the joint
            Case "JT_OY": Joint(JointKey).y = Val(Value) + YShift 'Load the y position of the joint
            Case "JT_OZ": Joint(JointKey).z = Val(Value) + ZShift 'Load the z position of the joint
            Case "OB_GP": Geometery(ObjectKey).Group.AddGroup Value 'Create a new group level for this object
            Case "OB_LK": Geometery(ObjectKey).Locked = Value 'Sets weather the object is locked
            Case "OB_HD": Geometery(ObjectKey).Hidden = Value '..or hidden...
            Case "OB_GY": Geometery(ObjectKey).Grayed = Value '..or grayed out
            Case "OB_CL": Geometery(ObjectKey).Colour = Value 'Set the colour of the object
            Case "OB_DF": Geometery(ObjectKey).Diffusion = Value 'Set the diffusion level of the object
            Case "OB_TC": Geometery(ObjectKey).Transparancy = Value 'Set the transparancy of the object
            Case "OB_TM": Geometery(ObjectKey).TextureMode = Value 'Set the object texture mode
            Case "OB_GN": Geometery(ObjectKey).grain = Value 'Set the garininess of the object
            Case "OB_TX": Geometery(ObjectKey).Texture = Value 'Set the texture filename
            Case "OB_EN": Geometery(ObjectKey).EntityName = Value 'Set the entity name
            Case "OB_ED": Geometery(ObjectKey).EntityDetails = Value 'Set the entity details

            
            Case "OBEND": If ObjectKey <> "" Then Geometery(ObjectKey).FindObjectOutline
            Case "JT_PF": Joint(JointKey).JointProfileIndex = Val(Value)
            Case "OB_LY"
                If Layers.CountLayers = 0 Then Layers.AddLayer "Main", "Main": Layers(1).Default = True
                If Mode = 2 Then Geometery(ObjectKey).Layer = Layers.Default Else Geometery(ObjectKey).Layer = Value
                  
            Case "JTCNT"
                ReDim JointList(Val(Value), 2) As String
                For n = 1 To Value: JointList(n, 1) = "Load" & Timer & n & Rnd: Next n
                GetjtKEY = 0

            Case "OBCNT"
                ReDim ObjectList(Val(Value), 2) As String
                For n = 1 To Value: ObjectList(n, 1) = "Load" & Timer & n & Rnd: Next n
                GETObKEY = 0
            
            Case "OBNEW"
                'Create a new object
                If GETObKEY = -1 Then
                    If ObjectKey <> "" Then Geometery(ObjectKey).FindObjectOutline
                    ObjectKey = "Loaded" & Geometery.CountObjects
                    Geometery.CreateObject ObjectKey
                    If Mode = 1 Then Geometery(ObjectKey).Selected = True
                Else
                    GETObKEY = GETObKEY + 1
                    If ObjectKey <> "" Then Geometery(ObjectKey).FindObjectOutline
                    ObjectKey = ObjectList(GETObKEY, 1)
                    ObjectList(GETObKEY, 2) = Value
                    Geometery.CreateObject ObjectKey
                    If Mode = 1 Then Geometery(ObjectKey).Selected = True
                End If

            Case "JTNEW"
                'Create a new Joint
                If GetjtKEY = -1 Then
                    JointKey = Value
                    Joint.AddJoint JointKey
                Else
                    GetjtKEY = GetjtKEY + 1
                    JointKey = JointList(GetjtKEY, 1)
                    JointList(GetjtKEY, 2) = Value
                    Joint.AddJoint JointKey
                    If Mode = 1 Then Joint(JointKey).Selected = True
                End If
                
            Case "FM_KY"
                FrameKey = Value
                Scene(SceneKey).CreateFrame FrameKey 'Load the frame key, and create the new frame
                
            Case "UDTJT"
                If GetjtKEY <> -1 Then
                    For n = 1 To GetjtKEY
                        For t = 1 To GetjtKEY
                            If JointList(t, 2) = Joint(JointList(n, 1)).Target Then Joint(JointList(n, 1)).Target = JointList(t, 1)
                        Next t
                    Next n
                End If
                
            Case "NEWRW"
                'Load the actual frame contents for the joint row named
                For n = 1 To GetjtKEY: If JointList(n, 2) = Value Then NewValue = JointList(n, 1)
                Next n: Scene(SceneKey)(FrameKey).UpdateFrameContents
                Input #1, Tmp: Scene(SceneKey)(FrameKey).Joint(NewValue).Name = Tmp
                Input #1, Tmp: Scene(SceneKey)(FrameKey).Joint(NewValue).AngleX = Tmp
                Input #1, Tmp: Scene(SceneKey)(FrameKey).Joint(NewValue).AngleY = Tmp
                Input #1, Tmp: Scene(SceneKey)(FrameKey).Joint(NewValue).AngleZ = Tmp
                Input #1, Tmp: Scene(SceneKey)(FrameKey).Joint(NewValue).OriginX = Tmp
                Input #1, Tmp: Scene(SceneKey)(FrameKey).Joint(NewValue).OriginY = Tmp
                Input #1, Tmp: Scene(SceneKey)(FrameKey).Joint(NewValue).OriginZ = Tmp
                Input #1, Tmp: Scene(SceneKey)(FrameKey).Joint(NewValue).ScaleX = Tmp
                Input #1, Tmp: Scene(SceneKey)(FrameKey).Joint(NewValue).ScaleY = Tmp
                Input #1, Tmp: Scene(SceneKey)(FrameKey).Joint(NewValue).ScaleZ = Tmp
                
            Case "MD_NT"
                'Sets the model notes
                Do
                    Input #1, NoteLine: If Mid(NoteLine, 1, 6) <> "MD_NT:" Then Notes = Notes & NoteLine
                Loop Until Mid(NoteLine, 1, 6) = "MD_NT:"
            
            Case "OB_TV"
                'Load a new vertex into this object
                Input #1, XX, YY: Geometery(ObjectKey).TexVert.Add XX, YY
            
            Case "OB_VT"
                'Load a new vertex into this object
                Input #1, TargetName, XX, YY, Zz
                Geometery(ObjectKey).Vertex.Add XX + XShift, YY + YShift, Zz + ZShift
                Geometery(ObjectKey).Vertex(Geometery(ObjectKey).Vertex.Count).Selected = True
                Geometery(ObjectKey).Vertex(Geometery(ObjectKey).Vertex.Count).TargetName = TargetName
        
            Case "OB_FC"
                'Load a new face into this object
                Geometery(ObjectKey).Face.Add Val(Value)
                For n = 1 To Val(Value): Input #1, Edge: Geometery(ObjectKey).Face(Geometery(ObjectKey).Face.Count).Edge.Add Edge
                Next n
            
            Case "OB2FC"
                'Load a new face into this object
                Geometery(ObjectKey).Face.Add Val(Value)
                For n = 1 To Val(Value): Input #1, Edge, TexEdge
                    Geometery(ObjectKey).Face(Geometery(ObjectKey).Face.Count).Edge.Add Edge, TexEdge
                Next n
            
            Case "JT_PV"
                Joint(JointKey).JProf(n) = Val(Value)
                For m = 1 To 5: Line Input #1, Value: Joint(JointKey).JProf(m) = Val(Mid(Value, 7))
                Next m
        
        End Select
    Loop While Funct <> "AM_QT"
    Close
    
    
    For Each Am In Geometery: For Each Vm In Am.Vertex: For t = 1 To GetjtKEY
        If JointList(t, 2) = Vm.TargetName Then Vm.TargetName = JointList(t, 1)
    Next t, Vm, Am
    If GETObKEY <> -1 Then
        For n = 1 To GETObKEY: For m = 1 To Geometery(ObjectList(n, 1)).Group.Count: For t = 1 To GETObKEY
            If ObjectList(t, 2) = Geometery(ObjectList(n, 1)).Group(Int(m)).GroupID Then
                Geometery(ObjectList(n, 1)).Group(Int(m)).GroupID = ObjectList(t, 1)
            End If
        Next t, m, n
    End If
    
    
    
    
    If Mode = 2 Then
        ActiveFile = Key
        Layers.ClearLayers
        Layers.AddLayer "Main", "Main"
        Scene.CreateScene "BaseFrame", "BaseFrame"
        Scene("BaseFrame").CreateFrame "BaseFrame"
        Scene("BaseFrame").CreateFrame "Animate"
        Scene("BaseFrame").CreateFrame "Inc"
        Scene("BaseFrame")("BaseFrame").UpdateFrameContents
        
    Else
        Scene.UpdateAllScenes
        MorphSkeliton "BaseFrame", "BaseFrame"
        If Mode = 0 Then Am8.FileHistory.AddHistory FileName
        If Scene.SceneCount = 0 Then Scene.CreateScene "BaseFrame", "BaseFrame": Scene("BaseFrame").CreateFrame "BaseFrame"
        Scene.ListScenesInWindow frmMain.cmbScenes
        If frmMain.cmbScenes.ListCount <> 0 Then frmMain.cmbScenes.ListIndex = 0
    End If
    
    
    If Scene("BaseFrame").FrameCount = 1 Then Scene("BaseFrame").CreateFrame "Animate": Scene("BaseFrame").CreateFrame "Inc"
    
    For Each Am In Geometery
        If Am.Layer = "" Then Am.Layer = Layers.Default
    Next Am
    
    If ModelName = "" Then ModelName = RightClip(Mid(FileName, InStrRev(FileName, "\") + 1), 4)
'   If Layers.CountLayers = 0 Then Layers.AddLayer "Main", "Main"
    
    FindModelOutline
    'Scene.UpdateAllScenes
    
    LoadFromFile = True
    
    
    
    Scene.UpdateAllScenes
    Scene.CopyToAnimate "BaseFrame", "BaseFrame"
    
Exit Function
FailedToLoadFile:
    MsgBox FileName & vbNewLine & vbNewLine & amFailedToLoadFile & ". " & Err.Description, vbExclamation
End Function

Public Sub MorphSkeliton(SceneON As String, FrameON As String)
    Dim Jm As clsJoint, Count As Integer
    If FrameON = "" Then Exit Sub
    For Each Jm In Joint
        Count = Count + 1
        If Jm.Target = "" Then
            Jm.AngleX = Scene(SceneON).Frame(FrameON).Joint(Count).AngleX
            Jm.AngleY = Scene(SceneON).Frame(FrameON).Joint(Count).AngleY
            Jm.AngleZ = Scene(SceneON).Frame(FrameON).Joint(Count).AngleZ
            Jm.NewPositX = Jm.X + Jm.PositionX
            Jm.NewPositY = Jm.y + Jm.PositionY
            Jm.NewPositZ = Jm.z + Jm.PositionZ
            WarpJoint Jm.Key, SceneON, FrameON
        End If
    Next Jm
End Sub

Private Function WarpJoint(JointName As String, SceneON As String, FrameON As String)
    Dim XX As Single, YY As Single, Zz As Single, Count As Integer
    Dim Jm As clsJoint, Rotated As clsVertex
    Set Rotated = New clsVertex
    For Each Jm In Joint
        Count = Count + 1
        If Jm.Target = JointName Then
            With Joint(Jm.Target)
                Rotated.X = Jm.X - .X
                Rotated.y = Jm.y - .y
                Rotated.z = Jm.z - .z
                Set Rotated = RotatePoint(Rotated, .AngleX, .AngleY, .AngleZ)
                Jm.NewPositX = Rotated.X + .NewPositX + Scene(SceneON).Frame(FrameON).Joint(Count).OriginX
                Jm.NewPositY = Rotated.y + .NewPositY + Scene(SceneON).Frame(FrameON).Joint(Count).OriginY
                Jm.NewPositZ = Rotated.z + .NewPositZ + Scene(SceneON).Frame(FrameON).Joint(Count).OriginZ
                Jm.AngleX = .AngleX + Scene(SceneON).Frame(FrameON).Joint(Count).AngleX
                Jm.AngleY = .AngleY + Scene(SceneON).Frame(FrameON).Joint(Count).AngleY
                Jm.AngleZ = .AngleZ + Scene(SceneON).Frame(FrameON).Joint(Count).AngleZ
                WarpJoint Jm.Key, SceneON, FrameON
            End With
        End If
    Next Jm
End Function
