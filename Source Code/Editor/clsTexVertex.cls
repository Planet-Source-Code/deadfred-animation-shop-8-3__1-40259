VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTexVertex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public x As Integer, y As Integer

Public Selected As Boolean
Public TargetName As String

' ############################################################################
' #                                                                          #
' #   This is the tex vertex class, which holds contains the position of     #
' #  a 2D Co-ordinate Many vertecies are linked together by faces to create  #
' #                      the outline of an object                            #
' #                                                                          #
' ############################################################################

Public Sub Move(iX As Integer, iY As Integer, iZ As Integer)
    'This is the move command for the vertecis. You sppecify the positions to
    'move it by, along all three axies
    x = x + iX
    y = y + iY
End Sub

