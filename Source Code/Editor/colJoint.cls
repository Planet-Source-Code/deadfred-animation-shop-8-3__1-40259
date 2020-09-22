VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colJoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsVertex"
Attribute VB_Ext_KEY = "Member0" ,"clsVertex"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' ############################################################################
' #                                                                          #
' #   This is the collection of joint objects. Each file has one of these    #
' #    classes in it. It contains the code to display the collection in a    #
' #     tree window, and to manipulate and add joints to the collection      #
' #                                                                          #
' ############################################################################

Private AllJoints As Collection


Public Sub DisplayTreeInWindow(Window As TreeView, Optional JointKey As String = "", Optional SelectedOnly As Boolean = False, Optional EnsureVisible As Boolean = True, Optional ShowSelected As Boolean = False, Optional ShowColour As Boolean = False)
    'This is the public function that calls the pivate function that adds the joints to a tree view.
    'This is much nicer than in previous versions of the program. You can recreate the tree from
    'scratch at any time, verty easily, and put the tree in any tree view you want.
    Dim Am As clsJoint, SelectedJoint As Integer
    Window.Visible = False: Window.Nodes.Clear
    Window.Nodes.Add , , "BaseJoint", Am8(ActiveFile).ModelName, 2
    If JointKey = "" Then
        For Each Am In AllJoints
            If Am.Target = "" Then AddJointToTree Am.Key, Window, SelectedOnly, EnsureVisible, ShowSelected, ShowColour
        Next Am
    Else
        AddJointToTree JointKey, Window, SelectedOnly, EnsureVisible, ShowSelected, ShowColour
    End If
    If SelectedJoint > 0 And Window.Nodes.Count > SelectedJoint Then Window.Nodes(SelectedJoint).Selected = True
    Window.Visible = True
End Sub


Private Sub AddJointToTree(sJoint As String, Window As TreeView, Optional SelectedOnly As Boolean, Optional EnsureVisible As Boolean, Optional ShowSelected As Boolean, Optional ShowColour As Boolean)
    'This is the actual function that adds the joint to the tree view. It can call itself repeatedly to
    'add the joints children and their children to the tree. There can be as many levels of children
    'as you need.
    Dim Am As clsJoint, Pm As clsJoint, setorder As Integer, ICon As Integer
    For Each Am In AllJoints
        If sJoint = Am.Key Then
            If (SelectedOnly = True And Am.Selected = True) Or (SelectedOnly = False) Then
                If Am.Target = "" Then
                    Window.Nodes.Add "BaseJoint", 4, Am.Key, Am.Name
                ElseIf FindNode(Am.Target, Window) = False Then
                    Window.Nodes.Add , , Am.Key, Am.Name
                Else
                    Window.Nodes.Add Am.Target, 4, Am.Key, Am.Name
                End If
                ICon = 1
                If Am.Locked = True Then ICon = ICon + 2
                If Am.Hidden = True Then ICon = ICon + 1
                Window.Nodes(Am.Key).Image = ICon
                If EnsureVisible = True Then Window.Nodes(Am.Key).EnsureVisible
            End If
            For Each Pm In AllJoints
                If Pm.Target = Am.Key Then
                    AddJointToTree Pm.Key, Window, SelectedOnly, EnsureVisible, ShowSelected, ShowColour
                End If
            Next Pm
        End If
    Next Am
End Sub


Private Function FindNode(NodeKey As String, Window As TreeView) As Boolean
    'This is a private function to see the a node exists in a tree window.
    Dim XX As String
    On Error GoTo FailedToFindNode
        XX = Window.Nodes(NodeKey).Key
    FindNode = True
FailedToFindNode:
End Function


Public Function AddJoint(sKey As String, Optional sTarget As String = "") As clsJoint
    'This adds a new joint to the model. You must give a unique key. You can also give
    'a name of a target joint, and an order to diplay the joints in the tree. The order
    'dosn't really matter that much. Its purly cosmetic.
    Dim NewJoint As clsJoint
    Set NewJoint = New clsJoint
    NewJoint.Key = sKey
    NewJoint.Target = sTarget
    AllJoints.Add NewJoint, sKey
    Set AddJoint = NewJoint
    Set NewJoint = Nothing
End Function


Public Function FirstSelectedJoint() As Integer
    'This returns the index of the first selected joint
    Dim Jm As clsJoint, Counta As Integer
    For Each Jm In AllJoints
        Counta = Counta + 1
        If Jm.Selected = True Then FirstSelectedJoint = Counta: Exit Function
    Next Jm
End Function


Public Function CountChildren(Optional iJointName As String = "") As Integer
    'This selects all the children of a given joint. Where no joint is given,
    'every joint is counted
    Dim Am As clsJoint, iChildren As Integer
    For Each Am In AllJoints
        If Am.Target = iJointName Then iChildren = iChildren + CountChildren(Am.Key) + 1
    Next Am
    CountChildren = iChildren
End Function


Public Function RemoveChildren(Optional iJointName As String = "") As String
    'This removes all of a joints children. If you don't give a joint name, then
    'every single joint is removed.
    Dim Am As clsJoint
    For Each Am In AllJoints
        If Am.Target = iJointName Then RemoveChildren Am.Key: AllJoints.Remove Am.Key
    Next Am
End Function


Public Sub Remove(JointKey As String)
    'This removes a single joint, but as you can't have child joints without a parent
    'joint, this checks for the existance of children and removes them if they exist.
    If CountChildren(JointKey) <> 0 Then RemoveChildren JointKey
    AllJoints.Remove JointKey
End Sub


Public Sub MoveSelected(x As Single, y As Single, ViewMode As Integer)
    'This moves the selected object by the distance specified with X and Y. It checks
    'to see which view you are in (side view, top view, front view), and moves you
    'reletive to that view
    Dim NowON As Integer, Ox As Integer, Oy As Integer, oZ As Integer, n As Integer
    Dim Am As clsJoint
    If ViewMode = 1 Then Ox = x: Oy = 0: oZ = y
    If ViewMode = 2 Then Ox = x: Oy = y: oZ = 0
    If ViewMode = 3 Then Ox = 0: Oy = y: oZ = x
    For Each Am In Me
        If Am.Selected = True Then Am.Move Ox, Oy, oZ
    Next Am
    Am8(ActiveFile).FindModelOutline
End Sub


Public Function CountSelected() As Integer
    'This counts the number of selected joints
    Dim Jm As clsJoint
    For Each Jm In Me
        If Jm.Selected = True Then CountSelected = CountSelected + 1
    Next Jm
End Function


Private Sub Class_Initialize()
    'This sets the mCol variable to a new collection
    Set AllJoints = New Collection
End Sub


Private Sub Class_Terminate()
    'This removes the collection, saving momeory, and being all nice
    Set AllJoints = Nothing
End Sub


Public Property Get Item(vntIndexKey As Variant) As clsJoint
Attribute Item.VB_UserMemId = 0
    'This returns the contents of the specified joint
    Set Item = AllJoints(vntIndexKey)
End Property


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    'This is Collection stuff, used in the For each .. in syntax
    Set NewEnum = AllJoints.[_NewEnum]
End Property

