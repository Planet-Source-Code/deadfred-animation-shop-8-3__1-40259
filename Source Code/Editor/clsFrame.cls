VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' ############################################################################
' #                                                                          #
' #  This the the frame class. If contains details for the position of one   #
' #   frame in a sequence of animation. If contains a list of all the joints #
' # presant in the model, and the positions that each joint is in for this   #
' #                                   frame                                  #
' #                                                                          #
' ############################################################################

Public Key As String, Name As String, Speed As Single, Smooth As Integer
Dim AllJoints As Collection


Public Property Get Joint(vntIndexKey As Variant) As clsJointRow
Attribute Joint.VB_UserMemId = 0
    'This returns the required joint row
    Set Joint = AllJoints(vntIndexKey)
End Property


Private Function NewJointRow(sKey As String) As clsJointRow
    'Creates a new joint row. This function is private, because you shouldn't use
    'it directly. Instead, the Update frame function is used, which adds and
    'removes joint rows to match the existing Joint structure.
    Dim objNewMember As clsJointRow
    Set objNewMember = New clsJointRow
    objNewMember.Key = sKey
    AllJoints.Add objNewMember, sKey
    Set objNewMember = Nothing
End Function


Private Sub Class_Initialize()
    'Set AllJoints to be a new collection. Pretty much self explanitary really..
    Set AllJoints = New Collection
End Sub


Private Sub Class_Terminate()
    'Remove the collection. Save memory
    Set AllJoints = Nothing
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    'Used for the for each .. in syntax
    Set NewEnum = AllJoints.[_NewEnum]
End Property


Public Function UpdateFrameContents()
    'This updates the joint rows in the frame. Each joint row holds details for one joint, so if joints
    'are added or removed, then the frame needs to be updated to remove or add joint rows
    Dim Am As clsJoint, Pm As clsJointRow
    For Each Pm In Me
        If CheckForJoint(Pm.Key) = False Then AllJoints.Remove Pm.Key
    Next Pm
    For Each Am In Am8(ActiveFile).Joint
        If CheckForFrameRow(Am.Key) = False Then
            NewJointRow Am.Key
            Me(Am.Key).Name = Am.Name
            Me(Am.Key).AngleX = 0
            Me(Am.Key).AngleY = 0
            Me(Am.Key).AngleZ = 0
            Me(Am.Key).OriginY = 0
        End If
    Next Am
End Function


Private Function CheckForJoint(JointName As String) As Boolean
    'This checks for the existance of a joint. If the joint dosn't exist, the joint row can be removed
    Dim Am As clsJoint
    For Each Am In Am8(ActiveFile).Joint
        If Am.Key = JointName Then CheckForJoint = True: Exit Function
    Next Am
End Function


Private Function CheckForFrameRow(RowName As String) As Boolean
    'This checks for the existance of a jointrow. If it dosn't exist, then a new one needs to be added
    Dim x As Single
    On Error GoTo FrameRowDosntExist
        x = AllJoints(RowName).AngleX
        CheckForFrameRow = True
FrameRowDosntExist:
End Function


