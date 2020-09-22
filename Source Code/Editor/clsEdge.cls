VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEdge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' ############################################################################
' #                                                                          #
' #    10 lines of comments for one line of code! This is the Edge class,    #
' #  which is contained within the face class. Each face will have several   #
' #    edges, which hold the sequesnce of verteceis that make up the face    #
' #                                                                          #
' ############################################################################

Public Vertex As Integer
Public TexVertex As Integer
