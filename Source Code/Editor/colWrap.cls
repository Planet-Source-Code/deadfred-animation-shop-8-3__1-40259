VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colWrap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ############################################################################
' #                                                                          #
' #  This is the collection of wraps. You can have groups within groups by  #
' #                                                                          #
' ############################################################################

Private AllWrap As Collection




Public Sub Add(x As Integer, y As Integer)
    'Add a new level of grouping to this collection
    Dim NewWrap As clsWrap
    Set NewWrap = New clsWrap
    NewWrap.x = x: NewWrap.y = y
    AllWrap.Add NewWrap
    Set NewWrap = Nothing
End Sub


Public Property Get Object(Index As Integer) As clsWrap
Attribute Object.VB_UserMemId = 0
    'Get property
    Set Object = AllWrap(Index)
End Property


Public Sub RemoveGroup()
    'This removes the top level of grouping
    AllWrap.Remove 1
End Sub


Public Function Count() As Integer
    'Count the number of grouping layers this object is in
    If AllWrap Is Nothing Then Set AllWrap = New Collection
    Count = AllWrap.Count
End Function


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    'Used for the For Each ..In syntax. Remember to set the property to -4 in the
    'procedure propery window
    Set NewEnum = AllWrap.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'Set the AllEdge variable to a new collection
    Set AllWrap = New Collection
End Sub


Private Sub Class_Terminate()
    'Destroy the collection when the face is removed
    Set AllWrap = Nothing
End Sub

