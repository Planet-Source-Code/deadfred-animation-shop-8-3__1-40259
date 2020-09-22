VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colEdge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' ############################################################################
' #                                                                          #
' #  This is the collection of edges. Each face has one collection of edges, #
' # and each collection of edges can hold as many edges as required. Each    #
' # edge contains the index of a single. Three or more 'edges' can be linked #
' #  together to outline a face. When the verteceis in the face move, the    #
' #   face continues to outline from one vertex to the next. New edges can   #
' # be added to the face, or inserted within the existing face, increasing   #
' #     the number of edges that the face has. Edges can also be removed.    #
' #                                                                          #
' ############################################################################

Private AllEdge As Collection


Public Function Insert(Before As Integer, Vertex As Integer) As clsEdge
    'Insert a new edge within the existing face. You supply the position of the
    'new edge, and the vertex that the edge will use. You can't use this on an
    'empty face, or it will crash.
    Dim NewObject As clsEdge
    Set NewObject = New clsEdge
    NewObject.Vertex = Vertex
    AllEdge.Add NewObject, , Before
    Set NewObject = Nothing
End Function


Public Function Add(Vertex As Integer, Optional TexEdge As Integer = 0) As clsEdge
    'Add a new vertex onto the end on the face. You specify which vertex the edge will use
    Dim NewObject As clsEdge
    Set NewObject = New clsEdge
    NewObject.Vertex = Vertex
    NewObject.TexVertex = TexEdge
    AllEdge.Add NewObject
    Set NewObject = Nothing
End Function


Public Property Get Object(Index As Integer) As clsEdge
Attribute Object.VB_UserMemId = 0
    'Get property
    Set Object = AllEdge(Index)
End Property


Public Sub Remove(Index As Integer)
    'This removes the specified edge
    AllEdge.Remove (Index)
End Sub


Public Function Count() As Integer
    'Count the number of edges in this face
    Count = AllEdge.Count
End Function


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    'Used for the For Each ..In syntax. Remember to set the property to -4 in the
    'procedure propery window
    Set NewEnum = AllEdge.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'Set the AllEdge variable to a new collection
    Set AllEdge = New Collection
End Sub


Private Sub Class_Terminate()
    'Destroy the collection when the face is removed
    Set AllEdge = Nothing
End Sub



