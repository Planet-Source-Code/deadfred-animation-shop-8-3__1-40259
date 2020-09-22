VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScene"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Key As String
Public Name As String
Private AllFrames As Collection


Public Function FrameIndex(FrameKey As String) As Integer
    'This finds the index of a frame, given the frames key, and returns it.
    Dim Am As clsFrame, Counter As Integer
    For Each Am In AllFrames
        Counter = Counter + 1
        If Am.Key = FrameKey Then FrameIndex = Counter: Exit Function
    Next Am
End Function


Public Sub CreateFrame(sKey As String, Optional Position As Integer = 0)
    'This creates a new frame in the scene
    Dim NewFrame As clsFrame
    Set NewFrame = New clsFrame
    NewFrame.Key = sKey
    NewFrame.Name = sKey
    NewFrame.Smooth = 10
    If Position = 0 Or AllFrames.Count = 0 Then
        AllFrames.Add NewFrame, sKey
    Else
        AllFrames.Add NewFrame, sKey, Position
    End If
    Set NewFrame = Nothing
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    'This is Collection stuff, used in the For each .. in syntax
    Set NewEnum = AllFrames.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'This sets up the AllFrames variable at startup
    Set AllFrames = New Collection
End Sub


Private Sub Class_Terminate()
    'This removes the collection, saving momeory, and being all nice
    Set AllFrames = Nothing
End Sub


Public Function FrameCount() As Integer
    'This returns the number of frames in the scene
    FrameCount = AllFrames.Count
End Function


Public Sub RemoveFrame(vKey As Variant)
    'This removes a rfame from the scene
    AllFrames.Remove vKey
End Sub


Public Property Get Frame(vntIndexKey As Variant) As clsFrame
Attribute Frame.VB_UserMemId = 0
    'This returns the contents of the specified frame
    Set Frame = AllFrames(vntIndexKey)
End Property

