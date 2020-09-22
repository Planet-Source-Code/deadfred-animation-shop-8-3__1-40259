VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' ############################################################################
' #                                                                          #
' #  This is the object class, which contains the faces, vertecies and       #
' #  texture vertecies. This class contains some of the edit code, such as   #
' #              compress, fragment, seperate and so on...                   #
' #                                                                          #
' ############################################################################

Public Vertex As colVertex
Public TexVert As colTexVertex
Public Face As colFace
Public Group As colGroup

Public Key As String
Public Layer As String
Public IsShadow As Boolean

Public Selected As Boolean
Public Locked As Boolean
Public Hidden As Boolean
Public Grayed As Boolean
Public Colour As Long
Public Transparancy As Byte
Public grain As Byte
Public Diffusion As Byte
Public ForceShowFace As Boolean
Public Texture As String
Public TextureMode As Integer

Public EntityName As String
Public EntityDetails As String

Public MinX As Single, MaxX As Single
Public MinY As Single, MaxY As Single
Public MinZ As Single, MaxZ As Single


Public Function FragmentFace(FaceID As Integer, FragMode As Integer, Optional skl As Single = 0) As Integer
    Dim CenX As Single, CenY As Single, CenZ As Integer, n As Integer, Faces As Integer
    Dim X As Integer, y As Integer, z As Integer, NewFace As Integer
    
    For n = 1 To Face(FaceID).EdgeCount
        CenX = CenX + Vertex(Face(FaceID).Edge(n).Vertex).X
        CenY = CenY + Vertex(Face(FaceID).Edge(n).Vertex).y
        CenZ = CenZ + Vertex(Face(FaceID).Edge(n).Vertex).z
    Next n
    CenX = CenX / Face(FaceID).EdgeCount
    CenY = CenY / Face(FaceID).EdgeCount
    CenZ = CenZ / Face(FaceID).EdgeCount
    
    Select Case FragMode
        Case 0
            For n = 1 To Face(FaceID).EdgeCount - 2
                Face.Add 3, Face(FaceID).Edge(n).Vertex, Face(FaceID).Edge(n + 1).Vertex, Face(FaceID).Edge(Face(FaceID).EdgeCount).Vertex
            Next n
            Face.Remove FaceID
            
        Case 1
            Vertex.Add Int(CenX), Int(CenY), Int(CenZ)
            FragmentFace = Vertex.Count
            For n = 1 To Face(FaceID).EdgeCount - 1
                Face.Add 3, Face(FaceID).Edge(n).Vertex, Face(FaceID).Edge(n + 1).Vertex, FragmentFace
            Next n
            Face.Add 3, Face(FaceID).Edge(n).Vertex, Face(FaceID).Edge(1).Vertex, FragmentFace
            Face.Remove FaceID
        
        Case 2
            Face.Add Face(FaceID).EdgeCount
            For n = 1 To Face(FaceID).EdgeCount
                X = ((Vertex(Face(FaceID).Edge(n).Vertex).X - CenX) * skl) + CenX
                y = ((Vertex(Face(FaceID).Edge(n).Vertex).y - CenY) * skl) + CenY
                z = ((Vertex(Face(FaceID).Edge(n).Vertex).z - CenZ) * skl) + CenZ
                Vertex.Add X, y, z
                Face(Face.Count).Edge.Add Vertex.Count
            Next n
            FragmentFace = Face.Count - 1
            For n = 1 To Face(FaceID).EdgeCount - 1
                Face.Add 4, Face(FaceID).Edge(n).Vertex, Face(FaceID).Edge(n + 1).Vertex, Face(FragmentFace + 1).Edge(n + 1).Vertex, Face(FragmentFace + 1).Edge(n).Vertex
            Next n
            Face.Add 4, Face(FaceID).Edge(n).Vertex, Face(FaceID).Edge(1).Vertex, Face(FragmentFace + 1).Edge(1).Vertex, Face(FragmentFace + 1).Edge(n).Vertex
            Face.Remove FaceID
    
    End Select
End Function


Public Sub CompressObject()
    'This function compresses an object by finding all the repeated vertecies and removing them.
    'It makes sure that any faces that use a vertex that is repeated, uses the remaining vertex at that point
    Dim Vm As clsVertex, VertexString As String, n As Integer, FaceCount As Integer
    Dim VertexList() As clsVertex, Stn  As String, Em As clsEdge
    Dim XX As Integer, Xx2 As Integer, Xx3 As Integer, Xx4 As Integer
    frmMain.List1.Clear
    For Each Vm In Vertex
        VertexString = Vm.X & ", " & Vm.y & ", " & Vm.z & ", " & Vm.TargetName
        frmMain.List1.AddItem VertexString
    Next Vm
    Do
        n = n + 1
        If frmMain.List1.List(n) = frmMain.List1.List(n + 1) Then frmMain.List1.RemoveItem n + 1: n = n - 1
    Loop Until n = frmMain.List1.ListCount - 1
    
    ReDim VertexList(frmMain.List1.ListCount - 1) As clsVertex
    For n = 0 To frmMain.List1.ListCount - 1
        Set VertexList(n) = New clsVertex
        Stn = frmMain.List1.List(n)
        XX = InStr(1, Stn, ",")
        Xx2 = InStr(XX + 1, Stn, ",")
        Xx3 = InStr(Xx2 + 1, Stn, ",")
        Xx4 = InStr(Xx3, Stn, ",")
        VertexList(n).X = Val(Mid(Stn, 1, XX - 1))
        VertexList(n).y = Val(Mid(Stn, XX + 1, Xx2 - 1))
        VertexList(n).z = Val(Mid(Stn, Xx2 + 1, Xx3 - 1))
        'VertexList(n).targetname = Val(Mid(Stn, 1, Xx - 1))
    Next n
    For FaceCount = Face.Count To 1 Step -1
        For Each Em In Face(FaceCount).Edge
            VertexString = Vertex(Em.Vertex).X & ", " & Vertex(Em.Vertex).y & ", " & Vertex(Em.Vertex).z & ", " & Vertex(Em.Vertex).TargetName
            For n = 0 To frmMain.List1.ListCount - 1
                If frmMain.List1.List(n) = VertexString Then Em.Vertex = n + 1
            Next n
        Next Em
    Next FaceCount
    For n = 1 To Vertex.Count: Vertex.Remove 1: Next n
    For n = 0 To frmMain.List1.ListCount - 1
        Vertex.Add VertexList(Int(n)).X, VertexList(Int(n)).y, VertexList(Int(n)).z
        Vertex(Int(n + 1)).TargetName = VertexList(Int(n)).TargetName
    Next n
    frmMain.List1.Clear
End Sub


Public Sub DeleteVertecies(Optional VertexID As Integer)
    'This goes through each vertex in the object and removes any selected ones, and the faces that use them, or
    'removes a specified vertex to remove, if one is given
    Dim Fm As clsFace, Em As clsEdge
    Dim FaceCount As Integer, VertexCount As Integer
    For FaceCount = Face.Count To 1 Step -1
        For Each Em In Face(FaceCount).Edge
            If (Vertex(Em.Vertex).Selected = True And VertexID = 0) Or Em.Vertex = VertexID Then Face.Remove FaceCount: Exit For
        Next Em
    Next FaceCount
    For VertexCount = Vertex.Count To 1 Step -1
        If (Vertex(VertexCount).Selected = True And VertexID = 0) Or VertexID = VertexCount Then
            Vertex.Remove VertexCount
            For FaceCount = Face.Count To 1 Step -1
                For Each Em In Face(FaceCount).Edge
                    If Em.Vertex > VertexCount Then Em.Vertex = Em.Vertex - 1
                Next Em
            Next FaceCount
        End If
    Next VertexCount
    FindObjectOutline
End Sub


Public Sub SeperateVertecies()
    'This takes a single object, and turns it into two objects. All the selected vertecies go into one, and the rest
    'go into the other object
    Dim Vm As clsVertex, NewKey As String, FaceSelected As Boolean
    Dim XShift As Integer, YShift As Integer, ZShift As Integer, Nm As clsVertex
    Dim Fm As clsFace, Em As clsEdge, FaceON As Integer, VertexOn As Integer, VertexFound As Integer
    NewKey = "Split" & Timer & Rnd * 1000
    Am8(ActiveFile).Geometery.CreateObject NewKey
    With Am8(ActiveFile).Geometery(NewKey)
        For Each Fm In Face
            FaceSelected = False
            For Each Em In Fm.Edge
                If Vertex(Em.Vertex).Selected = True Then FaceSelected = True
            Next Em
            If FaceSelected = True Then
                .Face.Add Fm.EdgeCount
                FaceON = FaceON + 1
                For Each Em In Fm.Edge
                    VertexOn = 0
                    VertexFound = 0
                    For Each Nm In .Vertex
                        VertexOn = VertexOn + 1
                        If Nm.X = Vertex(Em.Vertex).X + XShift And Nm.y = Vertex(Em.Vertex).y + YShift And Nm.z = Vertex(Em.Vertex).z + ZShift Then VertexFound = VertexOn
                    Next Nm
                    If VertexFound > 0 Then
                        .Face(FaceON).Edge.Add VertexFound
                    Else
                        .Vertex.Add Vertex(Em.Vertex).X + XShift, Vertex(Em.Vertex).y + YShift, Vertex(Em.Vertex).z + ZShift
                        .Face(FaceON).Edge.Add .Vertex.Count
                    End If
                Next Em
            End If
        Next Fm
        .FindObjectOutline
        .Layer = Layer
        .grain = grain
        .Diffusion = Diffusion
        .Transparancy = Transparancy
        .Selected = True
    End With
    DeleteVertecies
End Sub


Public Sub FindObjectOutline()
    'This finds the size of the objects bounding box; Ie. the smallest posible that
    'can soround the entire object
    Dim Am As clsVertex
    MinX = Vertex(1).X:    MinY = Vertex(1).y
    MinZ = Vertex(1).z:    MaxX = Vertex(1).X
    MaxY = Vertex(1).y:    MaxZ = Vertex(1).z
    For Each Am In Vertex
        If Am.X > MaxX Then MaxX = Am.X
        If Am.y > MaxY Then MaxY = Am.y
        If Am.z > MaxZ Then MaxZ = Am.z
        If Am.X < MinX Then MinX = Am.X
        If Am.y < MinY Then MinY = Am.y
        If Am.z < MinZ Then MinZ = Am.z
    Next Am
End Sub


Public Sub FlipHorizontal(OnX As Boolean, OnY As Boolean, OnZ As Boolean)
    'This flips the object around the given axis or axies
    Dim Am As clsVertex
    For Each Am In Vertex
        If OnX = True Then Am.X = -Am.X
        If OnY = True Then Am.y = -Am.y
        If OnZ = True Then Am.z = -Am.z
    Next Am
End Sub


Public Sub AttachObjectTo(JointKey As String)
    'This sets all the vertecies in the object to attach to the given joint
    Dim Vm As clsVertex
    For Each Vm In Vertex
        If Vm.Selected = True Or frmMain.chkSelect(6) = 0 Then Vm.TargetName = JointKey
    Next Vm
End Sub


Public Sub Randomize(ViewMode As Byte)
    'This moves the vertecies in the object by a randome amount
    Dim Am As clsVertex
    For Each Am In Vertex
        If ViewMode = 1 Then Am.y = Am.y + (Rnd * 5) - 2
        If ViewMode = 3 Then Am.X = Am.X + (Rnd * 5) - 2
        If ViewMode = 2 Then Am.z = Am.z + (Rnd * 5) - 2
    Next Am
    FindObjectOutline
End Sub


Public Sub Rotate(Angle1 As Single, Angle2 As Single, Angle3 As Single, Cx As Single, Cy As Single, Cz As Single)
    'This does a proper rotates on the selected object, given the
    'angles around each axis, and the center of rotation
    Dim Am As clsVertex, Rotated As clsVertex
    For Each Am In Vertex
        If Am.Selected = True Or frmMain.chkSelect(6) = 0 Then
            Set Rotated = RotatePoint(Am, Angle1, Angle2, Angle3, Int(Cx), Int(Cy), Int(Cz))
            Am.X = Rotated.X: Am.y = Rotated.y: Am.z = Rotated.z
        End If
    Next Am
    FindObjectOutline
End Sub


Public Sub CreateObject(Class As String, ViewMode As Integer, Optional O1 As Integer = 0, Optional O2 As Integer = 0, Optional O3 As Integer = 0, Optional O4 As Integer = 0, Optional O5 As Integer = 0, Optional O6 As Integer = 0, Optional O7 As Integer = 0, Optional O8 As Integer = 0, Optional O9 As Integer = 0, Optional O10 As Integer = 0)
    'This creates a brand new 3D shape in this object. You supply several
    'parameters to choose which shape to create. Class would be a name
    'such as 'Cube', 'Sphere' etc. Each of the remaining optional parameters
    'are used to specify other attributes of the new object. Depending on the
    'class of shape, there will be different numbers of atributes, ueed for
    'different purposes.
    Create3DObject Me, ViewMode, Class, O1, O2, O3, O4, O5, O6, O7, O8, O9, O10
    FindObjectOutline
End Sub


Public Property Get EdgeFaceCount() As Integer
    'This returns the number of edges and faces combined
    EdgeFaceCount = EdgeCount + Face.Count
End Property


Private Function FaceNormal(Ner() As Integer, Edges As Byte) As Double
    'This takes all the corners of a face, and calculates the FaceNormal.
    'Instad of just working on the first 3 points, it takes points from
    'all around the face to give a more acurate answer.
    On Error Resume Next
    Select Case Edges
        Case 3, 4: FaceNormal = (CLng((Ner(1, 2) - Ner(3, 2))) * CLng((Ner(2, 1) - Ner(1, 1)))) - (CLng((Ner(1, 1) - Ner(3, 1))) * CLng((Ner(2, 2) - Ner(1, 2))))
        Case 5, 6, 7, 8, 9, 10: FaceNormal = (CLng((Ner(1, 2) - Ner(5, 2))) * CLng((Ner(3, 1) - Ner(1, 1)))) - (CLng((Ner(1, 1) - Ner(5, 1))) * CLng((Ner(3, 2) - Ner(1, 2))))
        Case Is > 10: FaceNormal = (CLng((Ner(1, 2) - Ner(9, 2))) * CLng((Ner(5, 1) - Ner(1, 1)))) - (CLng((Ner(1, 1) - Ner(9, 1))) * CLng((Ner(5, 2) - Ner(1, 2))))
    End Select
End Function


Public Sub DrawOutline(Window As Tablet, ViewMode As Byte, X As Single, y As Single)
    'This draws the outline for this object. The parameters X and Y allow the
    'outline to be moved around the screen Ie. for dragging a selection
    'Window.ForeColour = vbRed
    Select Case ViewMode
        Case 1:  Window.DrawBox MinX + X, MinZ + y, MaxX + X, MaxZ + y
        Case 2:  Window.DrawBox MinX + X, MinY + y, MaxX + X, MaxY + y
        Case 3:  Window.DrawBox MinZ + X, MinY + y, MaxZ + X, MaxY + y
    End Select
End Sub


Public Function SplitFace(iFace As Integer, Edge As Integer) As Integer
    'This functin inserts a new vertex into the side of a face
    'It can be used to turn a square into an octogon, by
    'increasing the number of sides.
    Dim x1 As Integer, y1 As Integer, z1 As Integer
    If Edge = Face(iFace).Edge.Count Then
        x1 = (Vertex(Face(iFace).Edge(Edge).Vertex).X + Vertex(Face(iFace).Edge(1).Vertex).X) * 0.5
        y1 = (Vertex(Face(iFace).Edge(Edge).Vertex).y + Vertex(Face(iFace).Edge(1).Vertex).y) * 0.5
        z1 = (Vertex(Face(iFace).Edge(Edge).Vertex).z + Vertex(Face(iFace).Edge(1).Vertex).z) * 0.5
        Vertex.Add x1, y1, z1
        Face(iFace).Edge.Add Vertex.Count
    Else
        x1 = (Vertex(Face(iFace).Edge(Edge).Vertex).X + Vertex(Face(iFace).Edge(Edge + 1).Vertex).X) * 0.5
        y1 = (Vertex(Face(iFace).Edge(Edge).Vertex).y + Vertex(Face(iFace).Edge(Edge + 1).Vertex).y) * 0.5
        z1 = (Vertex(Face(iFace).Edge(Edge).Vertex).z + Vertex(Face(iFace).Edge(Edge + 1).Vertex).z) * 0.5
        Vertex.Add x1, y1, z1
        Face(iFace).Edge.Insert Edge + 1, Vertex.Count
    End If
    SplitFace = Vertex.Count
End Function


Public Function Grow(X As Single, y As Single, z As Single, Cx As Single, Cy As Single, Cz As Single) As Boolean
    'This is the scale function, but Scale is a reserved word. You suppply the
    'amount to enlarge or shrink, and the center point of the change
    Dim Am As clsVertex
    If frmMain.chkSelect(6) = 0 Then
        For Each Am In Vertex
            Am.X = Am.X - Cx: Am.y = Am.y - Cy: Am.z = Am.z - Cz
            Am.X = Am.X * X: Am.y = Am.y * y: Am.z = Am.z * z
            Am.X = Am.X + Cx: Am.y = Am.y + Cy: Am.z = Am.z + Cz
        Next Am
    Else
        For Each Am In Vertex
            If Am.Selected = True Then
                Am.X = Am.X - Cx: Am.y = Am.y - Cy: Am.z = Am.z - Cz
                Am.X = Am.X * X: Am.y = Am.y * y: Am.z = Am.z * z
                Am.X = Am.X + Cx: Am.y = Am.y + Cy: Am.z = Am.z + Cz
            End If
        Next Am
    End If
End Function


Public Sub MoveFace(Face As Integer, iX As Single, iY As Single, iZ As Integer)
    'This moves a single face within the object. You have to specify the number
    'of the face that you want to move, and the distances to move it along each axis
    Dim Am As clsEdge
    For Each Am In Me.Face(Face).Edge
        Vertex(Am.Vertex).X = Vertex(Am.Vertex).X + iX
        Vertex(Am.Vertex).y = Vertex(Am.Vertex).y + iY
        Vertex(Am.Vertex).z = Vertex(Am.Vertex).z + iZ
    Next Am
End Sub


Public Sub ReverseFace(Optional FaceNumber As Integer = 0)
    'This reverses all the faces in the object, unless a specific face
    'number is given, in which case, only that face is reversed
    Dim Fm As clsFace
    If FaceNumber = 0 Then
        For Each Fm In Face: Fm.ReverseFace: Next Fm
    Else
        Face(FaceNumber).ReverseFace
    End If
End Sub


Public Function Move(X As Integer, y As Integer, z As Integer) As Boolean
    'This moves the object through all three axis
    Dim Am As clsVertex
    If frmMain.chkSelect(6) = 0 Then
        For Each Am In Vertex: Am.X = Am.X + X: Am.y = Am.y + y: Am.z = Am.z + z: Next Am
    Else
        For Each Am In Vertex
            If Am.Selected = True Then Am.X = Am.X + X: Am.y = Am.y + y: Am.z = Am.z + z
        Next Am
    End If
End Function


Public Property Get EdgeCount() As Integer
    'This counts the number of edges in this object. A square has 4 edges, a cube has 4*6 = 24 edges
    Dim Am As clsFace, CountEdges As Integer
    For Each Am In Face: CountEdges = CountEdges + Am.EdgeCount: Next Am
    EdgeCount = CountEdges
End Property


Private Sub Class_Initialize()
    'This sets the collections to their correct type when the
    'class is created
    Set Vertex = New colVertex
    Set TexVert = New colTexVertex
    Set Face = New colFace
    Set Group = New colGroup
End Sub


Private Sub Class_Terminate()
    'Set all the collections to nothing. Save memory :-)
    Set Group = Nothing
    Set Vertex = Nothing
    Set Face = Nothing
End Sub

