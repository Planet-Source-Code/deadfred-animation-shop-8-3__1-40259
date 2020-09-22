VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEntity 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entity Editor"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8160
   Icon            =   "frmProp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   8160
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAct 
      Caption         =   "Customize"
      Height          =   375
      Index           =   2
      Left            =   3840
      TabIndex        =   19
      ToolTipText     =   "Click to display the entity creator window"
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Frame Outer 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4575
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   7695
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   4215
         Left            =   0
         TabIndex        =   6
         Top             =   360
         Width           =   7695
         Begin VB.TextBox mluMulti 
            Height          =   315
            Index           =   0
            Left            =   3600
            MultiLine       =   -1  'True
            TabIndex        =   16
            Top             =   1200
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.ComboBox cmdEditcombo 
            Height          =   315
            Index           =   0
            Left            =   3600
            TabIndex        =   14
            Text            =   "Combo1"
            Top             =   720
            Visible         =   0   'False
            Width           =   1935
         End
         Begin MSComctlLib.Slider sldSlider 
            Height          =   275
            Index           =   0
            Left            =   3600
            TabIndex        =   7
            Top             =   240
            Visible         =   0   'False
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   476
            _Version        =   393216
         End
         Begin VB.TextBox filFile 
            Height          =   315
            Index           =   0
            Left            =   3600
            Locked          =   -1  'True
            MousePointer    =   1  'Arrow
            TabIndex        =   12
            Text            =   "Text1"
            Top             =   240
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.PictureBox colColour 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   0
            Left            =   3600
            ScaleHeight     =   285
            ScaleWidth      =   1905
            TabIndex        =   11
            Top             =   240
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.ComboBox cmbCombo 
            Height          =   315
            Index           =   0
            Left            =   3600
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   240
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.TextBox txtText 
            Height          =   315
            Index           =   0
            Left            =   3600
            TabIndex        =   8
            Text            =   "Text1"
            Top             =   240
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.CheckBox chkCheck 
            Height          =   315
            Index           =   0
            Left            =   3600
            TabIndex        =   9
            Top             =   240
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label lblHelp 
            Caption         =   "Label1"
            Height          =   255
            Index           =   0
            Left            =   3600
            TabIndex        =   18
            Top             =   2760
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Label1"
            Height          =   315
            Index           =   0
            Left            =   3720
            TabIndex        =   13
            Top             =   240
            Visible         =   0   'False
            Width           =   480
         End
      End
      Begin VB.Label lblShowMessage 
         Alignment       =   2  'Center
         Caption         =   "There are no properties on this page"
         Height          =   375
         Left            =   0
         TabIndex        =   15
         Top             =   1680
         Visible         =   0   'False
         Width           =   7695
      End
   End
   Begin VB.Timer Timer1 
      Left            =   4080
      Top             =   5760
   End
   Begin VB.CommandButton cmdAct 
      Caption         =   "Help"
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Click to get help on this window"
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton cmdAct 
      Caption         =   "Cancel"
      Height          =   375
      Index           =   1
      Left            =   6720
      TabIndex        =   3
      ToolTipText     =   "Click to cancel any changes made in this window"
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton cmdAct 
      Caption         =   "Okay"
      Height          =   375
      Index           =   0
      Left            =   5280
      TabIndex        =   2
      ToolTipText     =   "Click to confirm the changes made in this window"
      Top             =   6360
      Width           =   1335
   End
   Begin VB.ComboBox EntityList 
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Sets the type of entity that this object represents"
      Top             =   240
      Width           =   7935
   End
   Begin MSComDlg.CommonDialog GetWindow 
      Left            =   1200
      Top             =   -4360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.TabStrip TabMain 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   9128
      TabWidthStyle   =   1
      MultiRow        =   -1  'True
      HotTracking     =   -1  'True
      Separators      =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Properties"
            Object.ToolTipText     =   "Standard Proprties for most entities"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Extended Properties"
            Object.ToolTipText     =   "More specific properties to set how this entity behaves"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Advanced"
            Object.ToolTipText     =   "Advanced properties that may not be easy to use"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Flags"
            Object.ToolTipText     =   "Extra settings to customize this entity"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Extended Flags"
            Object.ToolTipText     =   "Lower level flag settings for this entity"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label lblShowDisc 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   720
      Width           =   7935
   End
End
Attribute VB_Name = "frmEntity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private PropertyCount As Integer
Private Model As clsFile

Private Sub UpdatePropertyScreen()
    'This routine displays all the controls to allow you to edit an entities properties. It has to unload
    'all the previous controls and reload the new ones in the righrt position
    Dim ControlObject As Control, ColumnOn As Integer
    Dim Property As clsProperty, NextControlHeight As Integer
    Dim LabelOn As Integer, ControlOn As Integer
    Dim StartP As Integer, EndP As Integer, n As Integer
    Dim ItemHeight As Integer, ItemLeft As Integer
    Dim WidestProperty As Integer
    'On Error GoTo FailedToLoadProperties
    Frame1.Visible = False: NextControlHeight = 100
    ItemLeft = 1100
    'This unloads all the loaded controls
    For n = txtText.Count - 1 To 1 Step -1:       Unload txtText(n):       Next n
    For n = chkCheck.Count - 1 To 1 Step -1:      Unload chkCheck(n):      Next n
    For n = lblLabel.Count - 1 To 1 Step -1:      Unload lblLabel(n):      Next n
    For n = filFile.Count - 1 To 1 Step -1:       Unload filFile(n):       Next n
    For n = cmbCombo.Count - 1 To 1 Step -1:      Unload cmbCombo(n):      Next n
    For n = cmdEditcombo.Count - 1 To 1 Step -1:  Unload cmdEditcombo(n):  Next n
    For n = colColour.Count - 1 To 1 Step -1:     Unload colColour(n):     Next n
    For n = sldSlider.Count - 1 To 1 Step -1:     Unload sldSlider(n):     Next n
    For n = mluMulti.Count - 1 To 1 Step -1:      Unload mluMulti(n):      Next n
    For n = lblHelp.Count - 1 To 1 Step -1:       Unload lblHelp(n):       Next n
    lblShowDisc.Caption = Am8.Properties(EntityList).Discription
    Frame1.Width = 10
    If Am8.Properties(EntityList.Text).Count <> 0 Then
        For Each Property In Am8.Properties(EntityList.Text)
            LabelOn = lblLabel.Count
            Load lblLabel(LabelOn)
            Select Case Property.PropertyType
            
                Case "help"
                    'A simple text box for entering words and such
                    ControlOn = lblHelp.Count: Load lblHelp(ControlOn)
                    Set ControlObject = lblHelp(ControlOn)
                    ControlObject.Caption = Property.Default
                    If Val(Property.Choice) > 1 Then ControlObject.Height = ControlObject.Height * Property.Choice
            
                Case "text"
                    If Val(Property.Choice) > 1 Then
                        'A simple text box for entering words and such
                        ControlOn = mluMulti.Count: Load mluMulti(ControlOn)
                        Set ControlObject = mluMulti(ControlOn)
                        ControlObject.Height = ControlObject.Height * Property.Choice
                        ControlObject = Property.Default
                    Else
                        'A simple text box for entering words and such
                        ControlOn = txtText.Count: Load txtText(ControlOn)
                        Set ControlObject = txtText(ControlOn)
                        ControlObject = Property.Default
                    End If
                
                Case "tick"
                    'A tick box for entering yes/no values
                    ControlOn = chkCheck.Count: Load chkCheck(ControlOn)
                    Set ControlObject = chkCheck(ControlOn)
                    ControlObject = Property.Default
    
                Case "editcombo"
                    'Combo lists. Drop down list boxes where you can pick from a selection of options
                    ControlOn = cmdEditcombo.Count: Load cmdEditcombo(ControlOn)
                    Set ControlObject = cmdEditcombo(ControlOn)
                    StartP = 1: EndP = 1
                    Do
                        EndP = InStr(StartP, Property.Choice, "|")
                        If EndP = 0 Then EndP = Len(Property.Choice) + 1
                        ControlObject.AddItem Mid(Property.Choice, StartP, EndP - StartP)
                        StartP = EndP + 1
                    Loop Until EndP = Len(Property.Choice) + 1
                    If Property.Default = Empty Then Property.Default = ""
                    ControlObject = Property.Default
                    
                Case "combo"
                    'Combo lists. Drop down list boxes where you can pick from a selection of options
                    ControlOn = cmbCombo.Count: Load cmbCombo(ControlOn)
                    Set ControlObject = cmbCombo(ControlOn)
                    StartP = 1: EndP = 1
                    Do
                        EndP = InStr(StartP, Property.Choice, "|")
                        If EndP = 0 Then EndP = Len(Property.Choice) + 1
                        ControlObject.AddItem Mid(Property.Choice, StartP, EndP - StartP)
                        StartP = EndP + 1
                    Loop Until EndP = Len(Property.Choice) + 1
                    If Property.Default = Empty Then Property.Default = 0
                    ControlObject.ListIndex = Val(Property.Default) - 1
    
                Case "colour"
                    'A box where when you click on it, you are allowed to choose a colour
                    ControlOn = colColour.Count: Load colColour(ControlOn)
                    Set ControlObject = colColour(ControlOn)
                    ControlObject.BackColor = Property.Default
    
                Case "file"
                    'A file name picker, where when you click on it, an Open File window apperas
                    ControlOn = filFile.Count: Load filFile(ControlOn)
                    Set ControlObject = filFile(ControlOn)
                    ControlObject.Tag = Property.Choice
                    ControlObject = Property.Default
    
                Case "slider"
                    'A slider control allowing you to pick a number from a range
                    ControlOn = sldSlider.Count: Load sldSlider(ControlOn)
                    Set ControlObject = sldSlider(ControlOn)
                    If LCase(Property.Choice) = "angle" Then
                        ControlObject.Max = 359
                        ControlObject.Min = 0
                        ControlObject.Tag = "Angle"
                    End If
                    If InStr(1, Property.Choice, "|") <> 0 Then
                        ControlObject.Max = Val(Mid(Property.Choice, InStr(1, Property.Choice, "|") + 1))
                        ControlObject.Min = Val(Mid(Property.Choice, 1, InStr(1, Property.Choice, "|") - 1))
                    End If
                    ControlObject = Property.Default
                    Select Case ControlObject.Max - ControlObject.Min
                        Case 1 To 10: ControlObject.TickFrequency = 1
                        Case 11 To 30: ControlObject.TickFrequency = 2
                        Case 31 To 80: ControlObject.TickFrequency = 5
                        Case 81 To 180: ControlObject.TickFrequency = 10
                        Case 181 To 360: ControlObject.TickFrequency = 30
                        Case Is > 360: ControlObject.TickFrequency = 50
                    End Select
            End Select
            If Property.NewPage = True Then
                If ColumnOn / 2 = Int(ColumnOn / 2) Then
                    ColumnOn = ColumnOn + 2
                    NextControlHeight = 100
                    ItemLeft = ItemLeft + 8000
                Else
                    ColumnOn = ColumnOn + 1
                    NextControlHeight = 100
                    ItemLeft = ItemLeft + 4000
                End If
            End If
            If NextControlHeight > 4000 Or Property.NewColumn = True Then
                ColumnOn = ColumnOn + 1
                WidestProperty = 0
                NextControlHeight = 100
                ItemLeft = ItemLeft + 4000
            End If
            If Property.Width > WidestProperty Then WidestProperty = Property.Width
            If Property.Locked = True Then lblLabel(LabelOn).ForeColor = vbBlue
            ControlObject.Visible = True
            ControlObject.Width = Property.Width
            ControlObject.Top = NextControlHeight
            ControlObject.Left = ItemLeft
            ControlObject.Enabled = Not Property.Locked
            ControlObject.ToolTipText = Property.ToolTip
            lblLabel(LabelOn).Tag = LabelOn
            lblLabel(LabelOn).Top = ControlObject.Top + 50
            lblLabel(LabelOn).Left = ControlObject.Left - lblLabel(LabelOn).Width - 100
            lblLabel(LabelOn).Visible = True
            If Property.PropertyType = "help" Then lblLabel(LabelOn).Visible = False
            NextControlHeight = NextControlHeight + 185 + ControlObject.Height
            lblLabel(LabelOn) = Property.PropertyName
        Next Property
        PropertyCount = LabelOn
        Frame1.Width = ControlObject.Left + WidestProperty + 600
        Frame1.Visible = True
    End If
    If Frame1.Visible = False Or (Frame1.Left + Frame1.Width < 0) Then lblShowMessage.Visible = True Else lblShowMessage.Visible = False
    ShowValues
Exit Sub
FailedToLoadProperties:
    MsgBox "One or more of the properties was invalid, and could not be displayed" & vbNewLine & vbNewLine & "Please check the entity creator window to find the problem", vbExclamation
End Sub


Private Property Get PropertyValue(PropIndex As Integer) As Variant
    Dim Prop As Control, NextControl As Boolean
    For Each Prop In Me.Controls
        If NextControl = True Then
            Select Case TypeName(Prop)
                Case "TextBox", "Slider": PropertyValue = Prop
                Case "PictureBox": PropertyValue = Prop.BackColor
                Case "ComboBox": If Prop.Style = 0 Then PropertyValue = Prop.Text Else PropertyValue = Prop.ListIndex
            End Select
            Exit Property
        End If
        If Prop.Tag = PropIndex Then NextControl = True
    Next Prop
End Property

Private Property Let PropertyValue(PropIndex As Integer, ByVal vNewValue As Variant)
    Dim Prop As Control, NextControl As Boolean
    For Each Prop In Me.Controls
        If NextControl = True Then
            Select Case TypeName(Prop)
                Case "TextBox", "Slider":  Prop = vNewValue
                Case "PictureBox": Prop.BackColor = vNewValue
                Case "ComboBox": If Prop.Style = 0 Then Prop.Text = vNewValue Else Prop.ListIndex = vNewValue
            End Select
            Exit Property
        End If
        If Prop.Tag = PropIndex Then NextControl = True
    Next Prop
End Property





Private Sub cmdACT_Click(Index As Integer)
    Dim n As Integer, EntityString  As String, Am As clsObject
    Select Case Index
        Case 0
            For n = 1 To PropertyCount
                EntityString = EntityString & "[" & n & "=" & PropertyValue(n) & "]"
            Next n
            For Each Am In Model.Geometery
                If Am.Selected = True Then
                    Am.EntityDetails = EntityString
                    Am.EntityName = EntityList
                End If
            Next Am
            Unload Me
        Case 1: Unload Me
        Case 2: frmEntEdit.RunAtStart EntityList
        Case 3: Am8.ShowHelp "Creating New Entities"
    End Select
End Sub



Private Sub colColour_Click(Index As Integer)
    'When you click on a select colour box, this displays the 'Select Colour' common diolog box
    GetWindow.ShowColor
    colColour(Index).BackColor = GetWindow.Color
End Sub



Private Sub EntityList_Click()
    'When you click on an entity in the drop down list, this updates the propertys window
    Timer1.Interval = 1
End Sub



Private Sub filFile_Click(Index As Integer)
    'When you click on a set file name box, this displays the 'Open' diolog box
    If filFile(Index).ToolTipText <> "" Then GetWindow.DialogTitle = filFile(Index).ToolTipText
    If filFile(Index).Tag <> "" Then
        GetWindow.Filter = filFile(Index).Tag
        GetWindow.FilterIndex = 0
    End If
    GetWindow.ShowOpen
    If GetWindow.FileName <> "" Then filFile(Index) = GetWindow.FileName
End Sub



Private Sub UpdateDropdownList()
    'This displays the name of each entity in the drop down list at the top of the screen
    Dim Am As clsEntity, EntityOn As String, n As Integer
    EntityOn = EntityList: EntityList.Clear
    For Each Am In Am8.Properties: EntityList.AddItem Am.EntityName: Next Am
    For n = 1 To EntityList.ListCount
        If EntityList.List(n - 1) = EntityOn Then EntityList.ListIndex = n - 1
    Next n
End Sub



Private Sub Form_Activate()
    Am8.Properties.LoadEntities
    UpdateDropdownList
    If EntityList.ListCount = 0 Then Frame1.Visible = False
    If Frame1.Visible = False Or (Frame1.Left + Frame1.Width < 0) Then lblShowMessage.Visible = True Else lblShowMessage.Visible = False
End Sub


Private Sub sldSlider_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = 2 And sldSlider(Index).Tag = "Angle" Then
        Select Case sldSlider(Index).Value
            Case 0 To 89: sldSlider(Index).Value = 90
            Case 90 To 179: sldSlider(Index).Value = 180
            Case 179 To 269: sldSlider(Index).Value = 270
            Case 270 To 359: sldSlider(Index).Value = 0
        End Select
    End If
End Sub



Private Sub TabMain_Click()
    'When you press the different pages on the tab bar, the different properties are shown
    Frame1.Left = -(8000 * (TabMain.SelectedItem.Index - 1))
    If Frame1.Visible = False Or (Frame1.Left + Frame1.Width < 0) Then lblShowMessage.Visible = True Else lblShowMessage.Visible = False
End Sub



Private Sub TabMain_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    'When you press the different pages on the tab bar, the different properties are shown
    Frame1.Left = -(8000 * (TabMain.SelectedItem.Index - 1))
    If Frame1.Visible = False Or (Frame1.Left + Frame1.Width < 0) Then lblShowMessage.Visible = True Else lblShowMessage.Visible = False
End Sub



Private Sub Timer1_Timer()
    'For some bizzard reason, you cant unload controls from within a combo_click event. To get around this,
    'you use a timer to call the unload events after .001 second, and then set the timer interval to nothing.
    'Then, when the combo_click event is called, it sets the timer to .0001, and .0001 seconds later, the
    'properties are unloaded and reloaded, and you cant tell the difference!
    UpdatePropertyScreen
    Timer1.Interval = 0
End Sub



Private Function GetDetail(StringLine As String, DetailIndex As Integer) As Variant
    Dim iStart As Integer, n As Integer
    iStart = 0
    For n = 1 To DetailIndex
        iStart = InStr(iStart + 1, StringLine, "[")
    Next n
    iStart = InStr(iStart, StringLine, "=") + 1
    GetDetail = Mid(StringLine, iStart, InStr(iStart, StringLine, "]") - iStart)
End Function



Public Sub RunAtStart(AssignedFile As clsFile)
    Set Model = AssignedFile
    UpdateDropdownList
    If EntityList.ListCount > 0 Then EntityList.ListIndex = 0 Else Frame1.Visible = False
    ShowValues
    Show vbModal
End Sub


Private Sub ShowValues()
    Dim n As Integer
    With Model.Geometery(Model.Geometery.FirstSelectedObject)
        If .EntityName <> "" Then
            For n = 1 To EntityList.ListCount
                If EntityList.List(n - 1) = .EntityName Then EntityList.ListIndex = n - 1
            Next n
            For n = 1 To PropertyCount
                PropertyValue(n) = GetDetail(.EntityDetails, n)
            Next n
        End If
    End With
End Sub




