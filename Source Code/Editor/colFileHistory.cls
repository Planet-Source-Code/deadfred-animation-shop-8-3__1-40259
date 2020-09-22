VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colFileHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim AllHistory As Collection
Public Lenght As Integer

Public Sub AddHistory(FileName As String)
    'Creates a new class and puts it into the collection
    Dim NewHistory As clsFileHistory, Hm As clsFileHistory
    Set NewHistory = New clsFileHistory
    NewHistory.FilePath = FileName
    NewHistory.FileName = Mid(FileName, InStrRev(FileName, "\") + 1)
    For n = AllHistory.Count To 1 Step -1
        If AllHistory(n).FilePath = FileName Then AllHistory.Remove n
    Next n
    If AllHistory.Count = 0 Then
        AllHistory.Add NewHistory
    Else
        AllHistory.Add NewHistory, , 1
    End If
    Set NewHistory = Nothing
End Sub
Public Property Get NewEnum() As IUnknown
    'This is Collection stuff, used in the For each .. in syntax
    Set NewEnum = AllHistory.[_NewEnum]
End Property
Public Function CountHistory() As Integer
    CountHistory = AllHistory.Count
End Function
Public Property Get History(Index As Variant) As clsFileHistory
Attribute History.VB_UserMemId = 0
    'Returns the layer object specified
    Set History = AllHistory(Index)
End Property
Public Sub ClearHistory()
    Dim n As Integer
    For n = 1 To AllHistory.Count
        AllHistory.Remove 1
    Next n
End Sub
Public Sub RemoveStyles(Index)
    AllHistory.Remove Index
End Sub
Private Sub Class_Initialize()
    Set AllHistory = New Collection
End Sub
Private Sub Class_Terminate()
    Set AllHistory = Nothing
End Sub
