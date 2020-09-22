VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' ############################################################################
' #                                                                          #
' #  This is the face class, which holds contains a number of Edge classes   #
' #  The edges are linked together to outline the face, but this class in    #
' #                      itself dosn't contain much code                     #
' #                                                                          #
' ############################################################################

Public Edge As colEdge, EdgeCount As Integer, Selected As Boolean


Public Function AddTextureMap(Optional f1 As Integer = 0, Optional F2 As Integer = 0, Optional F3 As Integer = 0, Optional F4 As Integer = 0, Optional F5 As Integer = 0) As clsFace
    'This creates a new face at the end of the existing list of faces. You must specify the
    'number of edges the face will have, and you  can also define upto the first 5 vertecies
    'directly. If you want to create a face with more than 5 edges, you will have to use the
    'edge.add function
    If f1 <> 0 Then Edge(1).TexVertex = f1
    If F2 <> 0 Then Edge(2).TexVertex = F2
    If F3 <> 0 Then Edge(3).TexVertex = F3
    If F4 <> 0 Then Edge(4).TexVertex = F4
    If F5 <> 0 Then Edge(5).TexVertex = F5
End Function


Public Sub RotateTextureFace()
    'This rotates the texture vertecies round by one vertex
    Dim n As Integer, Temp As Integer
    Temp = Edge(1).TexVertex
    For n = 1 To EdgeCount - 1: Edge(n).TexVertex = Edge(n + 1).TexVertex: Next n
    Edge(EdgeCount).TexVertex = Temp
End Sub


Private Sub Class_Initialize()
    'Sets the Edge varible to become a collection of Edge classes
    Set Edge = New colEdge
End Sub


Public Function ReverseFace() As Boolean
    'This code reverses the order that vertecies appear in a face, thus reversing the face
    Dim Am As clsEdge, n As Integer
    ReDim StoreFace(Edge.Count)
    For Each Am In Edge: n = n + 1: StoreFace(n) = Am.Vertex: Next Am
    For Each Am In Edge: Am.Vertex = StoreFace(n): n = n - 1: Next Am
End Function
