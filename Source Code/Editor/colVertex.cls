VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colVertex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private AllVertex As Collection

' ############################################################################
' #                                                                          #
' #  This is the colelction of vertecies. Each object has a collection of    #
' #  vertecies. Each vertex in the collection is a 3D point in space. When   #
' #  you have many of these point, you can link their positions together     #
' #  with a face. A face outlines a serise of vertecies, and that outline is #
' #                        what makes up your model                          #
' #                                                                          #
' ############################################################################


Public Function Insert(Before As Integer, x As Integer, y As Integer, z As Integer) As clsVertex
    'This inserts a new vertex in the middle of the existing vertex list. You specify
    'where in the list to place the new vertex, and the X, Y and Z position of the veretx
    Dim NewObject As clsVertex, sKey As Variant
    Set NewObject = New clsVertex
    NewObject.x = x
    NewObject.y = y
    NewObject.z = z
    AllVertex.Add NewObject, , Before
    Set NewObject = Nothing
End Function


Public Function Add(x As Integer, y As Integer, z As Integer) As clsVertex
    'This creates a new vertex at the end of the existing list of vertecies. You must
    'specify the X, Y and Z position of the vertex
    Dim NewObject As clsVertex, sKey As Variant
    Set NewObject = New clsVertex
    NewObject.x = x
    NewObject.y = y
    NewObject.z = z
    AllVertex.Add NewObject
    Set NewObject = Nothing
End Function


Public Property Get Object(Index As Integer) As clsVertex
Attribute Object.VB_UserMemId = 0
    'This is a collection related property
    Set Object = AllVertex(Index)
End Property


Public Sub Remove(Index As Integer)
    'This removes a vertex from the list
    AllVertex.Remove (Index)
End Sub


Public Function Count() As Integer
    'Returns the number of vertecies in this object
    Count = AllVertex.Count
End Function


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    'Used for the For Each ..In syntax. Remember to set the property to -4 in the
    'procedure propery window
    Set NewEnum = AllVertex.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'Sets the AllVertex varible to a vertex collection
    Set AllVertex = New Collection
End Sub


Private Sub Class_Terminate()
    'Removes the AllVertex collection
    Set AllVertex = Nothing
End Sub


Public Function SelectAll()
    'Allows you to select every vertex in this object
    Dim Am As clsVertex
    For Each Am In Me: Am.Selected = True: Next Am
End Function


Public Function DeselectAll()
    'Deselects all of the vertecies in this object
    Dim Am As clsVertex
    For Each Am In Me: Am.Selected = False: Next Am
End Function
