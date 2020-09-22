VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ############################################################################
' #                                                                          #
' #  Another load of comments for one line of code. This is for the grouping #
' #   functions Ieach object has a collection of this class to hold the      #
' #              grouping levels that the object is in.                      #
' #                                                                          #
' ############################################################################


Public GroupID As String
