VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


' ############################################################################
' #                                                                          #
' #   The layer class holds details about each layer in the model, whether   #
' #                  it is locked, visiable, grayed etc..                    #
' #                                                                          #
' ############################################################################

Public LayerName As String
Public LayerKey As String
Public LayerColour As Long
Public LayerLocked As Boolean
Public LayerHidden As Boolean
Public LayerGrayed As Boolean
Public Selected As Boolean
Public Default As Boolean
