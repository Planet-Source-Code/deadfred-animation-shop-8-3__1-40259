VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colScene"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsScene"
Attribute VB_Ext_KEY = "Member0" ,"clsScene"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private AllScenes As Collection


Public Function CopyToAnimate(SceneName As String, FrameName As String, Optional Increment As Integer = -1)
    Dim Sm As clsJointRow, n As Integer, TargetFrame As String
    For Each Sm In Me(SceneName)(FrameName)
        If Increment = -1 Then
            With Me("BaseFrame")("Animate")
                n = n + 1
                .Joint(n).AngleX = Sm.AngleX: .Joint(n).AngleY = Sm.AngleY
                .Joint(n).AngleZ = Sm.AngleZ: .Joint(n).OriginX = Sm.OriginX
                .Joint(n).OriginY = Sm.OriginY: .Joint(n).OriginZ = Sm.OriginZ
                .Joint(n).ScaleX = Sm.ScaleX: .Joint(n).ScaleY = Sm.ScaleY
                .Joint(n).ScaleZ = Sm.ScaleZ
            End With
        Else
            If Increment = 0 Then Increment = 1
            With Me("BaseFrame")("Inc")
                n = n + 1
                .Joint(n).AngleX = (Sm.AngleX - Me("BaseFrame")("Animate")(n).AngleX) / Increment
                .Joint(n).AngleY = (Sm.AngleY - Me("BaseFrame")("Animate")(n).AngleY) / Increment
                .Joint(n).AngleZ = (Sm.AngleZ - Me("BaseFrame")("Animate")(n).AngleZ) / Increment
                .Joint(n).OriginX = (Sm.OriginX - Me("BaseFrame")("Animate")(n).OriginX) / Increment
                .Joint(n).OriginY = (Sm.OriginY - Me("BaseFrame")("Animate")(n).OriginY) / Increment
                .Joint(n).OriginZ = (Sm.OriginZ - Me("BaseFrame")("Animate")(n).OriginZ) / Increment
                .Joint(n).ScaleX = (Sm.ScaleX - Me("BaseFrame")("Animate")(n).ScaleX) / Increment
                .Joint(n).ScaleY = (Sm.ScaleY - Me("BaseFrame")("Animate")(n).ScaleY) / Increment
                .Joint(n).ScaleZ = (Sm.ScaleZ - Me("BaseFrame")("Animate")(n).ScaleZ) / Increment
            End With
        End If
    Next Sm
End Function


Public Function MoveAnimation()
    Dim Sm As clsJointRow, n As Integer
    For Each Sm In Me("BaseFrame")("Animate")
        With Me("BaseFrame")("Inc")
            n = n + 1
            Sm.AngleX = Sm.AngleX + .Joint(n).AngleX
            Sm.AngleY = Sm.AngleY + .Joint(n).AngleY
            Sm.AngleZ = Sm.AngleZ + .Joint(n).AngleZ
            Sm.OriginX = Sm.OriginX + .Joint(n).OriginX
            Sm.OriginY = Sm.OriginY + .Joint(n).OriginY
            Sm.OriginZ = Sm.OriginZ + .Joint(n).OriginZ
            Sm.ScaleX = Sm.ScaleX + .Joint(n).ScaleX
            Sm.ScaleY = Sm.ScaleY + .Joint(n).ScaleY
            Sm.ScaleZ = Sm.ScaleZ + .Joint(n).ScaleZ
        End With
    Next Sm
End Function

Public Function AddSceneToWindow(Window As TreeView, Optional SceneName As String = "", Optional ShowBaseframe As Boolean = True, Optional ShowAll As Boolean = True)
    'This takes the details of the existing scenes and displays them in a Tree Viewobject. You can choose to
    'display a single scene, by specifying the scene name in the parameters, or display all the scenes bu
    'leaving the scene name empty. You can choose to display the baseframe, or to expand the tree automaticly
    Dim Am As clsScene, Pm As clsFrame, NodeOn As Integer, n As Integer
    Window.Visible = False: For n = 1 To Window.Nodes.Count: If Window.Nodes(n).Selected = True Then NodeOn = n
    Next n: Window.Nodes.Clear
    If ShowBaseframe = True Then Window.Nodes.Add , , "BaseFrame@BaseFrame", "Baseframe", 8: If ShowAll = True Then Window.Nodes("BaseFrame@BaseFrame").Expanded = True
    If SceneName = "" Then
        For Each Am In AllScenes
            If LCase(Am.Key) <> "baseframe" Then
                If ShowBaseframe = True Then Window.Nodes.Add "BaseFrame@BaseFrame", 4, Am.Key, Am.Name, 7 Else Window.Nodes.Add , , Am.Key, Am.Name, 7
                If ShowAll = True Then Window.Nodes(Am.Key).Expanded = True
                For Each Pm In AllScenes(Am.Key)
                    Window.Nodes.Add Am.Key, 4, Pm.Key & "@" & Am.Key, Pm.Name, 8
                Next Pm
            End If
        Next Am
    Else
        If ShowBaseframe = True Then Window.Nodes.Add "Baseframe@BaseFrame", 4, AllScenes(SceneName).Key, AllScenes(SceneName).Name, 7 Else Window.Nodes.Add , , AllScenes(SceneName).Key, AllScenes(SceneName).Name, 7
        If ShowAll = True Then Window.Nodes(AllScenes(SceneName).Key).Expanded = True
        For Each Pm In AllScenes(SceneName)
            Window.Nodes.Add SceneName, 4, Pm.Key & "@" & SceneName, Pm.Name, 8
        Next Pm
    End If
    Window.Visible = True: If NodeOn <> 0 And NodeOn <= Window.Nodes.Count Then Window.Nodes(NodeOn).Selected = True Else Window.Nodes(1).Selected = True
End Function


Public Function CreateScene(SceneKey As String, SceneName As String) As clsScene
    'This creates a new scene, with no frames in it.
    Dim NewScene As clsScene
    Set NewScene = New clsScene
    NewScene.Key = SceneKey
    NewScene.Name = SceneName
    AllScenes.Add NewScene, SceneKey
    Set NewScene = Nothing
End Function

Public Function GetScene(TreeKey As String) As String
    'This picks the Scene key out of the node key. The key of each node contains the frame and scene
    'together, to create a unique key
     GetScene = Mid(TreeKey, InStr(1, TreeKey, "@") + 1)
End Function


Public Function GetFrame(TreeKey As String) As String
    'This picks the Frame key out of the node key. The key of each node contains the frame and scene
    'together, to create a unique key
    If InStr(1, TreeKey, "@") = 0 Then Exit Function
    GetFrame = Mid(TreeKey, 1, InStr(1, TreeKey, "@") - 1)
End Function


Public Function UpdateAllScenes()
    'This updates every frame in every scene to take into acount changes within the Skelital structure
    Dim Am As clsScene, Pm As clsFrame
    If ActiveFile <> "" Then
        For Each Am In Me
            For Each Pm In Am: Pm.UpdateFrameContents: Next Pm
        Next Am
    End If
End Function


Public Function ListScenesInWindow(Window As ComboBox)
    'This displays all of the scenes in a dropdown box
    Dim Am As clsScene: Window.Clear
    For Each Am In AllScenes: Window.AddItem Am.Name: Next Am
End Function


Public Sub EraseScenes()
    Dim Scene As clsScene
    For Each Scene In AllScenes: RemoveScene Scene.Key: Next Scene
End Sub


Public Sub RemoveScene(vntIndexKey As Variant)
    'Removes the specified scene
    AllScenes.Remove vntIndexKey
End Sub


Public Property Get SceneCount() As Long
    'Counts the number of scenes
    SceneCount = AllScenes.Count
End Property


Public Property Get Frame(vntIndexKey As Variant) As clsScene
Attribute Frame.VB_UserMemId = 0
    'Collection stuff, returns the specified scene
    Set Frame = AllScenes(vntIndexKey)
End Property


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    'Used for the For each .. In syntax
    Set NewEnum = AllScenes.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'Set the AllScenes as a collection
    Set AllScenes = New Collection
End Sub


Private Sub Class_Terminate()
    'Remove the Allscenes
    Set AllScenes = Nothing
End Sub
