VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ############################################################################
' #                                                                          #
' #  This is the collection of groups. You can have groups within groups by  #
' #  adding two group objects to the group collection. It only looks at the  #
' #  top group object, but when you ungroup, it removes the top level, and   #
' #  if there are further groups underneath, these are exposed and used to   #
' #                             define the grouping                          #
' #                                                                          #
' ############################################################################

Private AllGroup As Collection


Public Function AddGroup(GroupKey As String) As clsGroup
    'Add a new level of grouping to this collection
    Dim NewObject As clsGroup
    Set NewObject = New clsGroup
    NewObject.GroupID = GroupKey
    If AllGroup.Count = 0 Then AllGroup.Add NewObject Else AllGroup.Add NewObject, , 1
    Set NewObject = Nothing
End Function


Public Property Get Object(Index As Integer) As clsGroup
Attribute Object.VB_UserMemId = 0
    'Get property
    Set Object = AllGroup(Index)
End Property


Public Sub RemoveGroup()
    'This removes the top level of grouping
    AllGroup.Remove 1
End Sub


Public Function Count() As Integer
    'Count the number of grouping layers this object is in
    Count = AllGroup.Count
End Function


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    'Used for the For Each ..In syntax. Remember to set the property to -4 in the
    'procedure propery window
    Set NewEnum = AllGroup.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'Set the AllEdge variable to a new collection
    Set AllGroup = New Collection
End Sub


Private Sub Class_Terminate()
    'Destroy the collection when the face is removed
    Set AllGroup = Nothing
End Sub
