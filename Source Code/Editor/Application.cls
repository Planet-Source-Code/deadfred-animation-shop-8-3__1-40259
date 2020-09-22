VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Application"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"colFile"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' ############################################################################
' #                                                                          #
' #     This is the main class of the program. All files are within this     #
' #   class. This also contains system settings, and displays alot of the    #
' #                      general system screens.                             #
' #                                                                          #
' ############################################################################


Public Forms As Collection 'The collection that holds each edit form in the program, so that they can be referenced later

Public Properties As colEntity



Private lSnapSize As Integer, AllFiles As colFile
Public pShowNoDX As Boolean, OpenGallary As Integer
Public ShowLayers As Boolean, ShowStatusBar As Boolean
Public ShowSidebar As Boolean, LeftSidebar As Boolean
Public ErrorLog As Boolean, AlwaysCenter As Boolean
Public FullPath As Boolean, ShowNewWindow As Boolean
Public ConfirmCloseNoSave As Boolean, HighLightSection As Boolean
Public SnapSize As Integer, ShowTips As Boolean
Public LightStyle As colLightStyle
Public FileHistory As colFileHistory

Public Function OpenAnimator() As Boolean
    'This opens the help program, and jumps to the topic specified. If you don't give
    'a topic name, it will open to the start screen
    On Error GoTo Bodged
    Dim x As String
    x = App.Path & "\Player.exe x"
    Shell x, vbNormalFocus
    OpenAnimator = True
    Exit Function
Bodged:
    MsgBox "There was an error starting the Animation Viewer" & vbNewLine & vbNewLine & Err.Description, vbCritical, "Error"
End Function


Public Function ShowHelp(Optional HelpTopic As String) As Boolean
    'This opens the help program, and jumps to the topic specified. If you don't give
    'a topic name, it will open to the start screen
    On Error GoTo Bodged
    Dim x As String
    If HelpTopic = "EditME!!" Then
        x = App.Path & "\Help.exe <edit=yes> "
    Else
        x = App.Path & "\Help.exe  <page=" & HelpTopic & ">  <confirm=no> <edit=no> <password=deathtodreamland>"
    End If
    Shell x, vbNormalFocus
    ShowHelp = True
    Exit Function
Bodged:
    MsgBox "There was an error starting the help program" & vbNewLine & vbNewLine & Err.Description, vbCritical, "Error"
End Function


Public Sub LoadQuickHelp(TargetForm As frmEdit)
    'This loads the Quick Help menu into the supplied edit file
    Dim HelpCount As Integer, Helpname As String, HelpContent As String
    If Dir(App.Path & "\data\quickhelp.dat", 1) <> "" Then
        Open App.Path & "\data\quickhelp.dat" For Input As #1
            Do While EOF(1) = False
                HelpCount = HelpCount + 1
                Input #1, Helpname, HelpContent
                Load TargetForm.meuQuickHelp(HelpCount)
                TargetForm.meuQuickHelp(HelpCount).Visible = True
                TargetForm.meuQuickHelp(HelpCount).Caption = Helpname
                TargetForm.meuQuickHelp(HelpCount).Tag = HelpContent
            Loop
            TargetForm.meuQuickHelp(0).Visible = False
        Close
    Else
        TargetForm.mnuHelp(5).Visible = False
    End If
End Sub


Private Sub Class_Initialize()
    Set AllFiles = New colFile
    Set LightStyle = New colLightStyle
    Set FileHistory = New colFileHistory
    Set Properties = New colEntity
    Set Forms = New Collection
    SnapSize = 10
End Sub


Public Sub ShowNew()
    'This displays the 'new file' window
    frmNew.RunAtStart
End Sub


Public Sub ShowAbout()
    'This displays the about box
    frmAbout.RunAtStart
End Sub


Private Sub Class_Terminate()
    'Empties the collection of files
    Set AllFiles = Nothing
End Sub


Public Sub ShowTipofDay()
    'This displays the naff Tip of Day window, but its got to be done!
    frmTip.RunAtStart
End Sub


Public Sub ShowSettings()
    'This displays the program settings window
    frmSettings.RunAtStart
End Sub


Public Property Get File() As colFile
Attribute File.VB_UserMemId = 0
    'Returns the All files collection class, containing all the files loaded in memory
    Set File = AllFiles
End Property
Public Property Set File(vData As colFile)
    Set AllFiles = vData
End Property

