VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' ############################################################################
' #                                                                          #
' #  This is the object collection. There is one collection for each model.  #
' #   It contains all the different objects that make up a model. Each       #
' # object can have many vertecies, and many faces, and can be as complex    #
' #           as nessessary, or as simple as a single face                   #
' #                                                                          #
' ############################################################################

Private AllObjects As Collection


Public Function CombineObject()
    Dim First As clsObject, FirstOne As Boolean, Am As clsObject
    Dim Fm As clsFace, Em As clsEdge, Vm As clsVertex, n As Integer, m As Integer
    Set First = Object(FirstSelectedObject)
    For Each Am In Me
        If Am.Selected = True Then
            If FirstOne = True Then
                For n = 1 To Am.Face.Count
                    First.Face.Add Am.Face(n).EdgeCount
                    For m = 1 To Am.Face(n).EdgeCount
                        First.Face(First.Face.Count).Edge.Add Am.Face(n).Edge(m).Vertex + First.Vertex.Count, Am.Face(n).Edge(m).TexVertex
                    Next m
                Next n
                For Each Vm In Am.Vertex
                    First.Vertex.Add Vm.x, Vm.y, Vm.z
                Next Vm
                RemoveObject Am.Key
            End If
            FirstOne = True
        End If
    Next Am
    First.FindObjectOutline
End Function


Public Function Align(Direction As Integer, ViewMode As Byte)
    Dim Am As clsObject
    With Am8(ActiveFile)
        For Each Am In Me
            If Am.Selected = True Then
                Select Case ViewMode
                    Case 1
                        If Direction = 1 Then Am.Move 0, 0, -(Am.MinZ - .MinZ)
                        If Direction = 2 Then Am.Move 0, 0, (((.MinZ + .MaxZ) / 2) - ((Am.MinZ + Am.MaxZ) / 2))
                        If Direction = 3 Then Am.Move 0, 0, (.MaxZ - Am.MaxZ)
                        If Direction = 5 Then Am.Move -(Am.MinX - .MinX), 0, 0
                        If Direction = 6 Then Am.Move (((.MinX + .MaxX) / 2) - ((Am.MinX + Am.MaxX) / 2)), 0, 0
                        If Direction = 7 Then Am.Move (.MaxX - Am.MaxX), 0, 0
                    
                    Case 2
                        If Direction = 1 Then Am.Move 0, -(Am.MinY - .MinY), 0
                        If Direction = 2 Then Am.Move 0, (((.MinY + .MaxY) / 2) - ((Am.MinY + Am.MaxY) / 2)), 0
                        If Direction = 3 Then Am.Move 0, (.MaxY - Am.MaxY), 0
                        If Direction = 5 Then Am.Move -(Am.MinX - .MinX), 0, 0
                        If Direction = 6 Then Am.Move (((.MinX + .MaxX) / 2) - ((Am.MinX + Am.MaxX) / 2)), 0, 0
                        If Direction = 7 Then Am.Move (.MaxX - Am.MaxX), 0, 0
                    
                    Case 3
                        If Direction = 1 Then Am.Move 0, -(Am.MinY - .MinY), 0
                        If Direction = 2 Then Am.Move 0, (((.MinY + .MaxY) / 2) - ((Am.MinY + Am.MaxY) / 2)), 0
                        If Direction = 3 Then Am.Move 0, (.MaxY - Am.MaxY), 0
                        If Direction = 5 Then Am.Move 0, 0, -(Am.MinZ - .MinZ)
                        If Direction = 6 Then Am.Move 0, 0, (((.MinZ + .MaxZ) / 2) - ((Am.MinZ + Am.MaxZ) / 2))
                        If Direction = 7 Then Am.Move 0, 0, (.MaxZ - Am.MaxZ)
                End Select
                Am.FindObjectOutline
            End If
        Next Am
    End With
End Function


Public Function FlipSelected(Direction As Byte) As Boolean
    Dim Am As clsObject
    With Am8(ActiveFile)
        For Each Am In Me
            If Am.Selected = True Then
                Select Case frmMain.ActiveForm.Tablet.ViewMode
                    Case 1
                        Am.Move -(.MinX + .MaxX) / 2, 0, -(.MinZ + .MaxZ) / 2
                        If Direction = 1 Then Am.FlipHorizontal True, False, False
                        If Direction = 2 Then Am.FlipHorizontal False, False, True
                        Am.Move (.MinX + .MaxX) / 2, 0, (.MinZ + .MaxZ) / 2
                        
                    Case 2
                        Am.Move -(.MinX + .MaxX) / 2, -(.MinY + .MaxY) / 2, 0
                        If Direction = 1 Then Am.FlipHorizontal True, False, False
                        If Direction = 2 Then Am.FlipHorizontal False, True, False
                        Am.Move (.MinX + .MaxX) / 2, (.MinY + .MaxY) / 2, 0
                    
                    Case 3
                        Am.Move 0, -(.MinY + .MaxY) / 2, -(.MinZ + .MaxZ) / 2
                        If Direction = 1 Then Am.FlipHorizontal False, False, True
                        If Direction = 2 Then Am.FlipHorizontal False, True, False
                        Am.Move 0, (.MinY + .MaxY) / 2, (.MinZ + .MaxZ) / 2
    
                End Select
                Am.ReverseFace
                Am.FindObjectOutline
            End If
        Next Am
    .FindModelOutline
    End With
End Function


Public Function DeselectAllVertecies()
    'This deselects all objects
    Dim Am As clsObject, Vm As clsVertex
    For Each Am In Me: For Each Vm In Am.Vertex: Vm.Selected = False: Next Vm: Next Am
End Function


Public Function FirstSelectedObject() As Integer
    'This returns the number of the first selected object
    Dim Am As clsObject, Count As Integer
    For Each Am In AllObjects
        Count = Count + 1
        If Am.Selected = True Then FirstSelectedObject = Count: Exit Function
    Next Am
End Function


Public Function CreateObject(sKey As String) As clsObject
    'This creates a new object. You have to specify a unique name to identify the
    'object. To start, the object will have no vertecies or faces, and will be invisible
    On Error GoTo CouldNotCreateObject
    Dim NewObject As clsObject
    Set NewObject = New clsObject
    NewObject.Key = sKey
    AllObjects.Add NewObject, sKey
    Set NewObject = Nothing
    Exit Function
CouldNotCreateObject:
    MsgBox amShapeError & vbNewLine & vbNewLine & Err.Description
End Function


Public Property Get Object(vntIndexKey As Variant) As clsObject
Attribute Object.VB_UserMemId = 0
    'Collection type stuff
    Set Object = AllObjects(vntIndexKey)
End Property


Public Sub RemoveObject(vntIndexKey As Variant)
    'This removes the specified object
    On Error Resume Next
    AllObjects.Remove vntIndexKey
End Sub


Public Function CountSelected() As Integer
    'This counts and returns the number of selected objects
    Dim Am As clsObject, n As Integer
    For Each Am In Me
        If Am.Selected = True Then CountSelected = CountSelected + 1
    Next Am
End Function


Public Function CountVertecies() As Integer
    'This counts the number of vertecies in all the objects
    Dim Am As clsObject, n As Integer
    For Each Am In Me: CountVertecies = CountVertecies + Am.Vertex.Count: Next Am
End Function


Public Function CountFaces() As Integer
    'This counts the number of faces in all the objects
    Dim Am As clsObject, n As Integer
    For Each Am In Me: CountFaces = CountFaces + Am.Face.Count: Next Am
End Function


Public Function CountObjects() As Integer
    'This counts the total number of objects in the model
    CountObjects = AllObjects.Count
End Function


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    'Used for the For Each ..In syntax. Remember to set the property to -4 in the
    'procedure propery window
    Set NewEnum = AllObjects.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'This sets the AllObjects variable to be a collection
    Set AllObjects = New Collection
End Sub


Private Sub Class_Terminate()
    'This removes the AllObjects collection
    Set AllObjects = Nothing
End Sub


Public Function MoveSelected(x As Single, y As Single, ViewMode As Integer)
    'This moves the selected object by the distance specified with X and Y. It checks
    'to see which view you are in (side view, top view, front view), and moves you
    'reletive to that view
    Dim NowON As Integer, Ox As Integer, Oy As Integer, oZ As Integer, n As Integer
    Dim Am As clsObject
    If ViewMode = 1 Then Ox = x: Oy = 0: oZ = y
    If ViewMode = 2 Then Ox = x: Oy = y: oZ = 0
    If ViewMode = 3 Then Ox = 0: Oy = y: oZ = x
    For Each Am In Me
        If Am.Selected = True Then Am.Move Ox, Oy, oZ: Am.FindObjectOutline
    Next Am
    Am8(ActiveFile).FindModelOutline
End Function


Public Function GroupSelected()
    'This goes through each selected object and creates another group level for those objects
    Dim Am As clsObject
    For Each Am In Me
        If Am.Selected = True Then Am.Group.AddGroup Am8(ActiveFile).Geometery(Am8(ActiveFile).Geometery.FirstSelectedObject).Key
    Next Am
End Function


Public Function UngroupSelected()
    'This removes a level of grouping from the selected objects
    Dim Am As clsObject
    For Each Am In Me
        If Am.Selected = True Then If Am.Group.Count > 0 Then Am.Group.RemoveGroup
    Next Am
End Function

