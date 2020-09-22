VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsJointRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' ############################################################################
' #                                                                          #
' #  The joint row is what is held in each frame class. There is one joint   #
' #  row class for every joint in the model, and each joint row holds the    #
' #   position of the joint in the frame. Flicking through the frames fast   #
' #     changes the joint rows and moves the position of the model           #
' #                                                                          #
' ############################################################################


Public Key As String
Public Name As String

Public OriginX As Single, OriginY As Single, OriginZ As Single
Public AngleX As Single, AngleY As Single, AngleZ As Single
Public ScaleX As Single, ScaleY As Single, ScaleZ As Single

