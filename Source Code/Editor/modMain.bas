Attribute VB_Name = "modMain"
Option Explicit

'#####################################################################
'#                                                                   #
'#   This module contains the start-up code, namly sub Main(). This  #
'#     is where the program begins. It also includes a few other     #
'#               inportant function and variables                    #
'#                                                                   #
'#####################################################################


Public Am8 As Application 'The program application that contains all the files and system variables used in the program
Public ActiveFile As String
Public Const Pie As Double = 57.2084022218701
Public Sine(-3610 To 3610) As Single
Public Cosine(-3610 To 3610) As Single
Public DirectXNotAvaliable As Boolean
Public Const JointWid = 2


Public Sub Main()
    'This is the start up code, where everything begins
    Set Am8 = New Application
    Make_LookUp
    Load frmMain
    LoadProgramSettings
    Am8.LightStyle.UpdateLightList frmMain.lstLightStyle
    frmMain.UpdateHistoryMenu
    
    Am8.Properties.LoadEntities
    
    If Command = "" Then
        If Am8.ShowTips = True Then Am8.ShowTipofDay
        If Am8.ShowNewWindow = True Then Am8.ShowNew
    Else
        frmMain.LoadExistingFileWithWindow Command
    End If
    frmMain.sBar.Panels(2) = amWelcomeMessage
End Sub


Public Function SetNewShapeMenu(ShapeName As String)
    'This displays the different slide bar propeties for each type of shape. When you create
    'a new shape, and select a shape type from the list, this function is used to hide or
    'show the different options
    Select Case ShapeName
        Case "Cube":    SetPropertysOnorOff
        Case "Face":    SetPropertysOnorOff 1, 2
        Case "Sphere":  SetPropertysOnorOff 1, 5
        Case "Torous":  SetPropertysOnorOff 1, 2, 5
        Case "Cone":    SetPropertysOnorOff 1, 2
        Case "Dimond":  SetPropertysOnorOff 1, 2
        Case "Prism":   SetPropertysOnorOff 1, 2, 3, 4
        Case "Star":    SetPropertysOnorOff 1, 2, 3, 4
        Case "Grid":    SetPropertysOnorOff 3, 4
        Case "Wrap"
            SetPropertysOnorOff 1
            frmMain.EditLine(0).Visible = True
            frmMain.EditLine(1).Visible = True
            frmMain.EditLine(2).Visible = True
        Case "Rubix Cube": SetPropertysOnorOff 1
    End Select
End Function


Private Function SetPropertysOnorOff(Optional P1 As Byte = 0, Optional P2 As Byte = 0, Optional P3 As Byte = 0, Optional P4 As Byte = 0, Optional P5 As Byte = 0)
    'This is the part that actually hides or shows the slidebars
    Dim n As Byte
    For n = 1 To 5
        frmMain.ShpName(n).Visible = False
        frmMain.ShpProp(n).Visible = False
    Next n
    frmMain.EditLine(0).Visible = False
    frmMain.EditLine(1).Visible = False
    frmMain.EditLine(2).Visible = False
    If P1 <> 0 Then frmMain.ShpName(P1).Visible = True: frmMain.ShpProp(P1).Visible = True
    If P2 <> 0 Then frmMain.ShpName(P2).Visible = True: frmMain.ShpProp(P2).Visible = True
    If P3 <> 0 Then frmMain.ShpName(P3).Visible = True: frmMain.ShpProp(P3).Visible = True
    If P4 <> 0 Then frmMain.ShpName(P4).Visible = True: frmMain.ShpProp(P4).Visible = True
    If P5 <> 0 Then frmMain.ShpName(P5).Visible = True: frmMain.ShpProp(P5).Visible = True
End Function


Public Function SaveProgramSettings()
    'Save the program settings to the disk, so they can be reloaded next time
    Dim n As Integer, NumLights As Integer
    Dim Hm As clsFileHistory, Lm As clsLightStyle
    On Error GoTo CantSaveSettings
    Open App.Path & "\Data\Settings.dat" For Output As #1
    Print #1, "GLPG: "; frmMain.cmbGallary.ListIndex
    For n = Am8.FileHistory.CountHistory To 1 Step -1
        Print #1, "AMFH: "; Am8.FileHistory(n).FilePath
    Next n
    For n = 1 To Am8.LightStyle.CountStyles
        Print #1, "AMLS: "; Am8.LightStyle(n).Name
        Print #1, Am8.LightStyle(n).Pattern
    Next n
    For n = 1 To 4
        If frmMain.cBar.Bands(n).Visible = True Then Print #1, "TBAR: "; n
    Next n
    Print #1, "FLHT: "; Am8.FileHistory.Lenght
    Print #1, "SWLS: "; IntBo(Am8.LeftSidebar)
    Print #1, "SWAC: "; IntBo(Am8.AlwaysCenter)
    Print #1, "SWHL: "; IntBo(Am8.HighLightSection)
    Print #1, "SWFP: "; IntBo(Am8.FullPath)
    Print #1, "SWST: "; IntBo(Am8.ShowTips)
    Print #1, "SWNW: "; IntBo(Am8.ShowNewWindow)
    Print #1, "SBAR: "; IntBo(Am8.ShowSidebar)
    Print #1, "STBR: "; IntBo(Am8.ShowStatusBar)
    Print #1, "LAYA: "; IntBo(Am8.ShowLayers)
    Print #1, "GALL: "; Am8.OpenGallary
    Print #1, "SWDX: "; IntBo(Am8.pShowNoDX)
CantSaveSettings:
    Close
End Function


Public Function LoadProgramSettings()
    'This loads the file containing the program settings. As with all
    'file formats in this program, each line contains a Command and a value
    Dim Comand As String, Funct As String
    Dim n As Integer, Value As String
    On Error GoTo CantLoadSettings
    Open App.Path & "\Data\Settings.dat" For Input As #1
    Do While EOF(1) = False
        Line Input #1, Comand
        Funct = Mid(Comand, 1, InStr(1, Comand, ":") - 1)
        Value = Mid(Comand, InStr(1, Comand, ":") + 2, Len(Comand))
        Select Case Funct
            Case "FLHT": Am8.FileHistory.Lenght = Val(Value)
            Case "AMLS": Input #1, Comand: Am8.LightStyle.AddStyle Value, Comand
            Case "TBAR": frmMain.cBar.Bands(Val(Value)).Visible = True
            Case "SWLS": Am8.LeftSidebar = Value: If Value = 1 Then frmMain.SideFrame.Align = 3
            Case "SWAC": Am8.AlwaysCenter = Value
            Case "SWHL": Am8.HighLightSection = Value
            Case "SWFP": Am8.FullPath = Value
            Case "SWST": Am8.ShowTips = Value
            Case "SWNW": Am8.ShowNewWindow = Value
            Case "LAYA": Am8.ShowLayers = Value
            Case "STBR": If Value = 1 Then frmMain.sBar.Visible = True: Am8.ShowStatusBar = True
            Case "SBAR": If Value = 1 Then Am8.ShowSidebar = True
            Case "AMFH": Am8.FileHistory.AddHistory Value
            Case "GALL": Am8.OpenGallary = Val(Value)
            Case "SWDX": If Value = 1 Then Am8.pShowNoDX = True
        End Select
    Loop
CantLoadSettings:
    Close
End Function
