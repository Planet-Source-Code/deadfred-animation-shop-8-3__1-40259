VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colLightStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim AllStyle As Collection
Public Sub AddStyle(NewName As String, StylePattern As String)
    Dim NewStyle As clsLightStyle
    Set NewStyle = New clsLightStyle
    NewStyle.Name = NewName
    NewStyle.Pattern = StylePattern
    AllStyle.Add NewStyle
End Sub
Public Property Get NewEnum() As IUnknown
    Set NewEnum = AllStyle.[_NewEnum]
End Property
Public Function CountStyles() As Integer
    CountStyles = AllStyle.Count
End Function
Public Property Get Style(Index As Variant) As clsLightStyle
Attribute Style.VB_UserMemId = 0
    Set Style = AllStyle(Index)
End Property
Public Sub ClearStyles()
    Dim n As Integer
    For n = 1 To AllStyle.Count
        AllStyle.Remove 1
    Next n
End Sub
Public Sub RemoveStyles(Index)
    AllStyle.Remove Index
End Sub
Private Sub Class_Initialize()
    Set AllStyle = New Collection
End Sub
Private Sub Class_Terminate()
    Set AllStyle = Nothing
End Sub
Public Sub UpdateLightList(List As ComboBox)
    Dim n As Integer, LastItem As String, LastIndex As Integer
    LastIndex = List.ListIndex: LastItem = List.Text: List.Clear: List.AddItem "<None>"
    For n = 1 To AllStyle.Count: List.AddItem AllStyle(n).Name: Next n
    If LastItem = "" Then List.ListIndex = 0
    For n = 0 To List.ListCount - 1
        If List.List(n) = LastItem Then List.ListIndex = n
    Next n
    If List.ListIndex = -1 And LastIndex < List.ListCount Then List.ListIndex = LastIndex
End Sub














