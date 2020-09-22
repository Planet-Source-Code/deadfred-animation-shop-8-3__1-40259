VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colFace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' ############################################################################
' #                                                                          #
' #  This is the colletion of faces. Each object has one collection of faces #
' #  and each face can contain many edges. A cube will have a collection of  #
' #    6 faces, a single plain will have just one. You can add faces and     #
' #  insert the vertecies tat make up the face in one go, or add a face and  #
' #                then add the vertecies to the face later                  #
' #                                                                          #
' ############################################################################

Private AllFace As Collection


Public Function Insert(Before As Integer, EdgeCount As Integer, Optional f1 As Integer = 0, Optional F2 As Integer = 0, Optional F3 As Integer = 0, Optional F4 As Integer = 0, Optional F5 As Integer = 0, Optional F6 As Integer = 0, Optional F7 As Integer = 0, Optional F8 As Integer = 0, Optional F9 As Integer = 0, Optional F10 As Integer = 0, Optional F11 As Integer = 0, Optional F12 As Integer = 0) As clsFace
    'This inserts a new face within the existing list of faces. You have to specify where in
    'this list you want it to appear, and the number of edges the face will have. You
    'can also define upto the first 12 vertecies directly. If you want to create a face with
    'more than 12 edges, you will have to use the edge.add function
    Dim NewObject As clsFace
    Set NewObject = New clsFace
    If f1 <> 0 Then NewObject.Edge.Add f1
    If F2 <> 0 Then NewObject.Edge.Add F2
    If F3 <> 0 Then NewObject.Edge.Add F3
    If F4 <> 0 Then NewObject.Edge.Add F4
    If F5 <> 0 Then NewObject.Edge.Add F5
    If F6 <> 0 Then NewObject.Edge.Add F6
    If F7 <> 0 Then NewObject.Edge.Add F7
    If F8 <> 0 Then NewObject.Edge.Add F8
    If F9 <> 0 Then NewObject.Edge.Add F9
    If F10 <> 0 Then NewObject.Edge.Add F10
    If F11 <> 0 Then NewObject.Edge.Add F11
    If F12 <> 0 Then NewObject.Edge.Add F12
    NewObject.EdgeCount = EdgeCount
    AllFace.Add NewObject, , Before
    Set NewObject = Nothing
End Function


Public Function Add(EdgeCount As Integer, Optional f1 As Integer = 0, Optional F2 As Integer = 0, Optional F3 As Integer = 0, Optional F4 As Integer = 0, Optional F5 As Integer = 0, Optional F6 As Integer = 0, Optional F7 As Integer = 0, Optional F8 As Integer = 0, Optional F9 As Integer = 0, Optional F10 As Integer = 0, Optional F11 As Integer = 0, Optional F12 As Integer = 0) As clsFace
    'This creates a new face at the end of the existing list of faces. You must specify the
    'number of edges the face will have, and you  can also define upto the first 12 vertecies
    'directly. If you want to create a face with more than 12 edges, you will have to use the
    'edge.add function
    Dim NewObject As clsFace
    Set NewObject = New clsFace
    If f1 <> 0 Then NewObject.Edge.Add f1
    If F2 <> 0 Then NewObject.Edge.Add F2
    If F3 <> 0 Then NewObject.Edge.Add F3
    If F4 <> 0 Then NewObject.Edge.Add F4
    If F5 <> 0 Then NewObject.Edge.Add F5
    If F6 <> 0 Then NewObject.Edge.Add F6
    If F7 <> 0 Then NewObject.Edge.Add F7
    If F8 <> 0 Then NewObject.Edge.Add F8
    If F9 <> 0 Then NewObject.Edge.Add F9
    If F10 <> 0 Then NewObject.Edge.Add F10
    If F11 <> 0 Then NewObject.Edge.Add F11
    If F12 <> 0 Then NewObject.Edge.Add F12
    NewObject.EdgeCount = EdgeCount
    AllFace.Add NewObject
    Set NewObject = Nothing
End Function


Public Property Get Object(Index As Integer) As clsFace
Attribute Object.VB_UserMemId = 0
    'Wierd collection type stuff
    Set Object = AllFace(Index)
End Property


Public Sub Remove(Index As Integer)
    'This removes the face that you specify from the list
    AllFace.Remove (Index)
End Sub


Public Function Count() As Integer
    'This counts the number of faces in the object
    Count = AllFace.Count
End Function


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    'Used for the For Each ..In syntax. Remember to set the property to -4 in the procedure propery window
    Set NewEnum = AllFace.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'Sets the AllFaces varible to be a collection
    Set AllFace = New Collection
End Sub


Private Sub Class_Terminate()
    'Sets the Collection to nothing
    Set AllFace = Nothing
End Sub


Public Function SelectAll()
    'Selects all the faces in the object
    Dim Am As clsFace
    For Each Am In Me: Am.Selected = True: Next Am
End Function


Public Function DeselectAll()
    'Deselects all the faces in the object
    Dim Am As clsFace
    For Each Am In Me: Am.Selected = False: Next Am
End Function

