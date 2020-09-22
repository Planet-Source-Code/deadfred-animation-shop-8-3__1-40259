VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsJoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

' #############################################################################
' #                                                                           #
' #  This class defines each joint in the model, and contains information to  #
' #   allow the storing and morphing of the skeliton. Each joint has a        #
' #           target veriable, which links it to another joint                #
' #                                                                           #
' #############################################################################

Public Key As String            'The unique name for the joint
Public Name As String           'The name displayed on screen
Public Target As String         'The key of the joint this joint is attached to
Public Selected As Boolean      'If the joint is currently selected

Public x As Single, y As Single, z As Single '  -  -  -  -  -  -  -  -  -  The aboslute  position of the joint
Public PositionX As Single, PositionY As Single, PositionZ As Single '  -  The position of the joint in the scene
Public AngleX As Single, AngleY As Single, AngleZ As Single '  -  -  -  -  The angle of the joint in the scene
Public ScaleX As Single, ScaleY As Single, ScaleZ As Single '  -  -  -  -  The scale of the joint in the scene
Public NewPositX As Single, NewPositY As Single, NewPositZ As Single '  -  The position of the joint after morphing

Public Locked As Boolean        'If the joint is locked
Public Grayed As Boolean        'If the joint is greyed
Public Hidden As Boolean        'If the joint is hidden
Public Colour As Long           'The colour of the joint

Public JointProfileIndex As Integer  'The preset joint profile used
Dim JointProfile(5) As Integer  'Used to store the joint profile information


Public Sub Move(abX As Integer, abY As Integer, abZ As Integer)
    'This moves the joint by the given amount
    x = x + abX: y = y + abY: z = z + abZ
End Sub


Public Function Grow(aX As Single, aY As Single, aZ As Single, Cx As Single, Cy As Single, Cz As Single) As Boolean
    'This is the scale function, but Scale is a reserved word. You suppply the
    'amount to enlarge or shrink, and the center point of the change
    x = x - Cx:    y = y - Cy:    z = z - Cz
    x = x * aX:    y = y * aY:    z = z * aZ
    x = x + Cx:    y = y + Cy:    z = z + Cz
End Function


Public Sub Rotate(Angle1 As Single, Angle2 As Single, Angle3 As Single, Cx As Single, Cy As Single, Cz As Single)
    'This does a proper rotates on the selected object, given the
    'angles around each axis, and the center of rotation
    Dim JointLocation As clsVertex, Rotated As clsVertex
    Set JointLocation = New clsVertex
    JointLocation.x = x: JointLocation.y = y: JointLocation.z = z
    Set Rotated = RotatePoint(JointLocation, Angle1, Angle2, Angle3, Int(Cx), Int(Cy), Int(Cz))
    x = Rotated.x: y = Rotated.y: z = Rotated.z
End Sub


Public Property Get JProf(ByVal Index As Integer) As Integer
    'This returns and sets the Joint Profile information
    JProf = JointProfile(Index)
End Property
Public Property Let JProf(ByVal Index As Integer, vNewValue As Integer)
    If Index < 6 Then JointProfile(Index) = vNewValue
End Property

