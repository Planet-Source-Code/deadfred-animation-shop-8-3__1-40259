VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colLayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ##############################################################################
' #                                                                            #
' #  This collection class holds all the layers for a file. No much code here  #
' #                                                                            #
' ##############################################################################

Dim AllLayers As Collection

Public Sub AddLayer(NewName As String, LayerKey As String)
    'Creates a new class and puts it into the collection
    On Error GoTo LayerExists
        Dim NewLayer As clsLayer
        Set NewLayer = New clsLayer
        NewLayer.LayerName = NewName
        NewLayer.Selected = True
        NewLayer.LayerKey = LayerKey
        AllLayers.Add NewLayer, LayerKey
    Exit Sub
LayerExists:
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    'This is Collection stuff, used in the For each .. in syntax
    Set NewEnum = AllLayers.[_NewEnum]
End Property

Public Function CountLayers() As Integer
    'Counts the number of layers in the file
    CountLayers = AllLayers.Count
End Function

Public Property Get Layer(Index As Variant) As clsLayer
Attribute Layer.VB_UserMemId = 0
    'Returns the layer object specified
    Set Layer = AllLayers(Index)
End Property

Public Function Default() As String
    'Returns the Default layer
    Dim Am As clsLayer
    For Each Am In AllLayers: If Am.Default = True Then Default = Am.LayerKey
    Next Am
End Function

Public Sub ClearLayers()
    'Removes every layer from the collection
    Dim Am As clsLayer
    For Each Am In AllLayers: AllLayers.Remove Am.LayerKey: Next Am
End Sub

Public Sub RemoveLayers(Index)
    'Removes a layer from the file
    AllLayers.Remove Index
End Sub

Private Sub Class_Initialize()
    'Sets up the collection
    Set AllLayers = New Collection
End Sub

Private Sub Class_Terminate()
    'Removes the collection
    Set AllLayers = Nothing
End Sub
