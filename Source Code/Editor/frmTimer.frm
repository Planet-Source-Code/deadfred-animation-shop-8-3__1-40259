VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTimer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exporting model... Please wait"
   ClientHeight    =   1050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4275
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1050
   ScaleWidth      =   4275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar TimeLeft 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Click this bar to stop compiling"
      Top             =   600
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Max             =   1000
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar FilesLeft 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "Click this bar to stop compiling"
      Top             =   1080
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Max             =   1000
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Percentage Done"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Model As clsFile

'#####################################################################
'#                                                                   #
'#   This form originally had no code in it at all. However, after   #
'#     getting things to work with the VBModal statment, all the     #
'#     compile code is now here, and its alot nicer here as well.    #
'#    It contains the code to export the model, and another sub to   #
'#                       compile entities                            #
'#                                                                   #
'#####################################################################


Private Sub TimeLeft_Click()
    'This is an attempted abort command, where you click on the scroll bar
    'to stop. Dosn't work too well, could do with some work
    If TimeLeft.ToolTipText = "" Then Exit Sub
    If MsgBox("Are you sure you want to stop exporting your model?", 292) = 6 Then
        If frmMain.Visible = True Then
            Unload frmTimer
            Unload frmCompile
        Else
            frmTimer.Visible = False
            End
        End If
    End If
End Sub


Public Sub RunAtStart(AssignedFile As clsFile, FileName As String, FileOn As Integer, TotalFiles As Integer)
    'This starts the compile proccess for what ever is loaded. It also
    'handles hiding and showing the form, so all you have to do is call
    'this sub to compile whatever is loaded
    Set Model = AssignedFile
    Unload frmCompile
    If TotalFiles <> 0 Then
        Height = 1905
        FilesLeft.Max = TotalFiles + 1
        FilesLeft = FileOn + 1
        Refresh
    End If
    Visible = True
    Open FileName For Output As #1
        CompleteObject
    Close
    If TotalFiles = 0 Then Unload Me
End Sub


Public Sub CompleteObject()
    'This compiles and writes the actual file to disk.
    Dim n As Long, EdgeOn As Long, MoveBack As Byte, UniqueNumber As Long
    Dim Fm As clsFace, Em As clsEdge, VertexString As String
    Dim Am As clsObject, Vm As clsVertex, Jm As clsJoint
    If frmCompile.chkCompile(1) = 0 Then MoveBack = 1 Else MoveBack = 0
    TimeLeft.Max = Model.Geometery.CountFaces
    TimeLeft = 0
    Print #1, "ID,"
    Print #1, "Discription , 'dsdas'"
    Print #1, "Cost , 0"
    Print #1, "Weight , 0"


    If frmCompile.chkCompile(1) = 0 Then MoveBack = 1
    frmCompile.List1.Clear: Refresh




    For Each Am In Model.Geometery
        For Each Vm In Am.Vertex
            If frmCompile.chkCompile(3) = 1 Then
                VertexString = Vm.X & ", " & Vm.y & ", " & Vm.z & ", " & Vm.TargetName
            Else
                VertexString = Vm.X & ", " & Vm.y & ", " & Vm.z
            End If
            frmCompile.List1.AddItem VertexString
        Next Vm
    Next Am

RestartLoop:
    For n = 0 To frmCompile.List1.ListCount - 2
        If frmCompile.List1.List(n) = frmCompile.List1.List(n + 1) Then
            frmCompile.List1.RemoveItem n + 1
            GoTo RestartLoop
        End If
    Next n
    ReDim VertexList(frmCompile.List1.ListCount - 1) As String
    For n = 0 To frmCompile.List1.ListCount - 1
        VertexList(n) = frmCompile.List1.List(n)
    Next n




    DoEvents
    Print #1, frmCompile.List1.ListCount
    Print #1, Model.Geometery.CountFaces
    For n = 0 To frmCompile.List1.ListCount - 1
        Print #1, frmCompile.List1.List(n)
    Next n




    For Each Am In Model.Geometery
        For Each Fm In Am.Face
            Print #1, Fm.EdgeCount; ", ";
            For Each Em In Fm.Edge
                EdgeOn = EdgeOn + 1
                If frmCompile.chkCompile(3) = 1 Then
                    VertexString = Am.Vertex(Em.Vertex).X & ", " & Am.Vertex(Em.Vertex).y & ", " & Am.Vertex(Em.Vertex).z & ", " & Am.Vertex(Em.Vertex).TargetName
                Else
                    VertexString = Am.Vertex(Em.Vertex).X & ", " & Am.Vertex(Em.Vertex).y & ", " & Am.Vertex(Em.Vertex).z
                End If
                For n = 0 To frmCompile.List1.ListCount
                    If VertexList(n) = VertexString Then UniqueNumber = n: Exit For
                Next n
                Print #1, UniqueNumber - MoveBack + 1;
                If EdgeOn <> Fm.EdgeCount Then Print #1, ", ";
            Next Em
            EdgeOn = 0
            Print #1,
            TimeLeft = TimeLeft + 1
        Next Fm
    Next Am



    Print #1, Model.Joint.CountChildren
    For Each Jm In Model.Joint
        Print #1, Jm.Name; ", ";
        Print #1, Jm.PositionX; ", ";
        Print #1, Jm.PositionY; ", ";
        Print #1, Jm.PositionZ; ", ";
        Print #1, Jm.Target
    Next Jm




End Sub




