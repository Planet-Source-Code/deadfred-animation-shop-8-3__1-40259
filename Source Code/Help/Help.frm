VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Animation Shop 8 Help"
   ClientHeight    =   7680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12660
   Icon            =   "Help.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7680
   ScaleWidth      =   12660
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Scroller 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   2520
      ScaleHeight     =   855
      ScaleWidth      =   135
      TabIndex        =   16
      Top             =   0
      Width           =   135
   End
   Begin MSComctlLib.ImageList Icons 
      Left            =   480
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Help.frx":0742
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Help.frx":0B96
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView Pages 
      Height          =   3855
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   6800
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   450
      Style           =   7
      ImageList       =   "Icons"
      Appearance      =   1
   End
   Begin VB.Frame Tools 
      Caption         =   "Edit Tools"
      Height          =   2175
      Left            =   2760
      TabIndex        =   1
      Top             =   3960
      Width           =   7095
      Begin VB.CheckBox ckStyle 
         Caption         =   "O__"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Bulleted"
         Top             =   960
         Width           =   495
      End
      Begin VB.ListBox lstColour 
         Height          =   1425
         Left            =   6000
         TabIndex        =   13
         Top             =   480
         Width           =   855
      End
      Begin VB.ListBox lstFontSize 
         Height          =   1425
         Left            =   5280
         TabIndex        =   12
         Top             =   480
         Width           =   615
      End
      Begin VB.OptionButton opTextPos 
         Caption         =   "Right"
         Height          =   255
         Index           =   1
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1680
         Width           =   975
      End
      Begin VB.OptionButton opTextPos 
         Caption         =   "Centre"
         Height          =   255
         Index           =   2
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1680
         Width           =   975
      End
      Begin VB.OptionButton opTextPos 
         Caption         =   "Left"
         Height          =   255
         Index           =   0
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1680
         Width           =   735
      End
      Begin VB.CheckBox ckStyle 
         Caption         =   "I"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Italic"
         Top             =   960
         Width           =   495
      End
      Begin VB.CheckBox ckStyle 
         Caption         =   "U"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Underline"
         Top             =   960
         Width           =   495
      End
      Begin VB.CheckBox ckStyle 
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Bold"
         Top             =   960
         Width           =   495
      End
      Begin VB.ComboBox cmdFonts 
         Height          =   315
         Left            =   2280
         Sorted          =   -1  'True
         TabIndex        =   5
         Text            =   "Combo1"
         Top             =   480
         Width           =   2655
      End
      Begin VB.CommandButton cmdAddPage 
         Caption         =   "Add Page"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove Page"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton cmdRename 
         Caption         =   "Rename Page"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   1440
         Width           =   1575
      End
   End
   Begin RichTextLib.RichTextBox PageDetail 
      Height          =   3255
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   5741
      _Version        =   393217
      HideSelection   =   0   'False
      ScrollBars      =   2
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Help.frx":0FEA
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DB As New txtDataBase
Dim PageOn As String
Dim EditMode As Boolean
Dim ConfirmExit As Boolean
Dim LocalPassword As String
Dim Dragging As Boolean
Const vbPurple = 16711935
Const vbSilver = 8224125
Const vbOrange = 39935
Const DefaultEncription As String = "Crack This!"

Private Function GetParameter2(ByVal Content, ByVal ParameterName) As String
    'This is another function to get values from a string. Its used on the command line to get the details
    'on what settings are used at the start
    Dim StartPoint As Integer, EndPoint As Integer
    Content = LCase(Content)
    If InStr(1, Content, "<" & ParameterName & "=") <> 0 And ParameterName <> "" Then
        StartPoint = InStr(1, Content, "<" & ParameterName & "=") + Len(ParameterName) + 2
        EndPoint = InStr(StartPoint, Content, ">")
        GetParameter2 = Mid(Content, StartPoint, EndPoint - StartPoint)
    End If
End Function

Private Sub ConfirmFileExistance()
    On Error GoTo CreateFile
    'Check to see if the file is there
    Open DB.FileName For Input As #99
    Close #99
Exit Sub
CreateFile:
    Open DB.FileName For Output As #99
    Close #99
    
    DB.EncriptionKey = DefaultEncription
    DB.CreatePage "READ-ME"
    DB.CreatePage "TREE"
    DB.DataStore("TREE") = "TITLE=<Enter Title Here>" & vbNewLine & "READ_PASSWORD=<>" & vbNewLine & "EDIT_PASSWORD=<>" & vbNewLine & "<TREE>" & vbNewLine & "<TREE>"
    
    DB.DataStore("READ-ME") = _
        "{\rtf1\ansi\ansicpg1252\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\fscript\fprq2\fcharset186 Comic Sans MS;}{\f3\froman Times New Roman;}}" & vbNewLine & _
        "{\colortbl\red0\green0\blue0;\red255\green0\blue0;\red0\green0\blue255;}" & vbNewLine & _
        "\deflang2057\horzdoc{\*\fchars }{\*\lchars }\pard\qc\plain\f3\fs24\b\ul Read Me\plain\f3\fs20" & vbNewLine & _
        "\par \pard\plain\f3\fs20" & vbNewLine & _
        "\par \plain\f2\fs20 This datastore has been opened for the first time, and these two pages have been made automaticly. The \plain\f2\fs20\cf2 TREE\plain\f2\fs20  page contains system details for this database, such as \plain\f2\fs20\cf1 layout\plain\f2\fs20 , \plain\f2\fs20\cf1 passwords\plain\f2\fs20  and \plain\f2\fs20\cf1 title\plain\f2\fs20 . You can remove this page of you want, but you MUST keep the \plain\f2\fs20\cf2 TREE\plain\f2\fs20  page as it is used by the system. The TREE page is not visible in read only mode\plain\f3\fs20" & vbNewLine & _
        "\par" & vbNewLine & _
        "\par \plain\f2\fs20 Use the controls at the bottom of the page to create new pages, and edit the text on each page. When you change pages, or close the progrma, your changes will automaticly be saved." & vbNewLine & _
        "\par }" & vbNewLine

End Sub

Private Sub Form_Load()
    Dim n As Integer, MainData As String, StartPage As String, TypeOfOpen As String
    
    If App.PrevInstance = True Then End
    
    'This sets whether or not you can edit the contents of the help file or not
    If GetParameter2(Command, "edit") = "yes" Then EditMode = True
    
    'If the EditMode is set to false then hide the tool bar and set the page window to locked
    If EditMode = False Then
        PageDetail.Locked = True
        Pages.LabelEdit = tvwManual
        Tools.Visible = False
    End If
    
    'This simply adds a list of common font sizes into the list box
    lstFontSize.AddItem "8":               lstFontSize.AddItem "10"
    lstFontSize.AddItem "11":              lstFontSize.AddItem "12"
    lstFontSize.AddItem "14":              lstFontSize.AddItem "16"
    lstFontSize.AddItem "18":              lstFontSize.AddItem "20"
    lstFontSize.AddItem "22":              lstFontSize.AddItem "26"
    lstFontSize.AddItem "28":              lstFontSize.AddItem "32"
    
    'This puts the list of colour names into the list box
    lstColour.AddItem "Black":             lstColour.AddItem "Silver"
    lstColour.AddItem "Red":               lstColour.AddItem "Green"
    lstColour.AddItem "Blue":              lstColour.AddItem "Orange"
    lstColour.AddItem "Purple"
    
    'This gets all the avalible fonts from the screen oject and puts them into the dropdown box
    For n = 1 To Screen.FontCount: cmdFonts.AddItem Screen.Fonts(n)
    Next n: cmdFonts.RemoveItem 0: cmdFonts.Text = "Times New Roman"
    
    'The first thing you do is set the name of a text file to use as a data store
    If GetParameter2(Command, "file") = "" Then DB.FileName = App.Path & "\data\am8.dat" Else DB.FileName = GetParameter2(Command, "file")
    
    'This checks whether or not the file exists, and if not, creates it
    ConfirmFileExistance
    
    'Next, set up an encription key. The key used to encript a page MUST be used to decript it
    If GetParameter2(Command, "encription") = "" Then DB.EncriptionKey = DefaultEncription Else DB.EncriptionKey = GetParameter2(Command, "encription")
    
    'The sets whether a Yes/No box is shown when you go to exit the program
    If GetParameter2(Command, "confirm") = "no" Then ConfirmExit = False Else ConfirmExit = True
    
    'This calls a procedure that displays all the existing pages in a list box
    UpdateListBox
    
    If EditMode = False Then TypeOfOpen = "To read the document, please enter the correct password now"
    If EditMode = True Then TypeOfOpen = "To edit the document, please enter the correct password now"
    
    'Check wether the file has a password. If it does, then it checks to see wether a password is given in the command line. If there
    'is, then it checks this password against the file password. If not, if asks the user for a password. In either case, a wrong
    'password results in an error message and the program closing.
    If LocalPassword <> "" Then
        If GetParameter2(Command, "password") = "" Then
            If LCase(InputBox("This file is protected." & vbNewLine & vbNewLine & TypeOfOpen, "Password")) <> LocalPassword Then
                MsgBox "You have entered an incorrect password", vbCritical, "Password": End
            End If
        Else
            If GetParameter2(Command, "password") <> LocalPassword Then
                MsgBox "You have entered an incorrect password", vbCritical, "Password": End
            End If
        End If
    End If

    'So long as there are pages in the list, the first page is displayed, unless an alternative page name is given, in which case,
    'it searches for that page in the list, and then selects that page to display.
    If Pages.Nodes.Count > 0 Then
        If GetParameter2(Command, "page") = "" Then
            Pages.Nodes(1).Selected = True
            PageOn = Pages.SelectedItem.Text
            PageDetail = DB(PageOn)
        Else
            StartPage = GetParameter2(Command, "page")
            For n = 1 To Pages.Nodes.Count
                If LCase(Pages.Nodes(n).Text) = StartPage Then
                    Pages.Nodes(n).Selected = True
                    PageOn = Pages.SelectedItem.Text
                    PageDetail = DB(PageOn)
                End If
            Next n
        End If
    End If
    
    Pages.Width = Scroller.Left - Pages.Left
    PageDetail.Left = Scroller.Left + Scroller.Width
    PageDetail.Width = ScaleWidth - PageDetail.Left - 50
    Tools.Left = Scroller.Left + Scroller.Width

End Sub

Private Sub ckStyle_Click(Index As Integer)
    'When you click on the Bold, Underlined or Italic buttons, this sets the selected font to
    'have that particular style
    Select Case Index
        Case 0: PageDetail.SelBold = ckStyle(Index)
        Case 1: PageDetail.SelUnderline = ckStyle(Index)
        Case 2: PageDetail.SelItalic = ckStyle(Index)
        Case 3: PageDetail.SelBullet = ckStyle(Index)
    End Select
End Sub

Private Sub cmdFonts_Click()
    'When you select a font from the dropdown list, this sets the selected text to use that font
    PageDetail.SelFontName = cmdFonts.Text
End Sub

Private Function GetParameter(Index As Integer, MainData As String) As String
    'This gets retives a parameter from a given string. A parameter is marked as being the text between two < ??? >
    Dim n As Integer, OldPosition As Integer, NewPosition As Integer
    For n = 1 To Index
        OldPosition = InStr(OldPosition + 1, MainData, "<"): NewPosition = InStr(OldPosition + 1, MainData, ">")
        If OldPosition = 0 Then Exit Function
        GetParameter = Mid(MainData, OldPosition + 1, NewPosition - OldPosition - 1)
    Next n
End Function

Private Sub UpdateListBox()
    'Declare this variable as an integer
    Dim TreeList(99) As String
    Dim PageNumber As Integer, MainData As String, n As Integer, LastPageName As String
    Dim OldPosition As Integer, NewPosition As Integer, NextPageName As String, Indent As Integer
    'Empty the list box
    Pages.Visible = False
    Pages.Nodes.Clear
    
    'Read in the main data page
    MainData = DB("TREE")
    'Skip past the first two data entries
    
    Caption = GetParameter(1, MainData)
    If EditMode = False Then LocalPassword = LCase(GetParameter(2, MainData)) Else LocalPassword = LCase(GetParameter(3, MainData))
    
    OldPosition = InStr(1, MainData, "<TREE>") + 6
    
    If EditMode = False Then
        'If your not in edit mode then read the tree properly from the data page
        Do
            OldPosition = InStr(OldPosition + 1, MainData, "<")
            NewPosition = InStr(OldPosition + 1, MainData, ">")
            If NewPosition = 0 Then Exit Sub
            NextPageName = Mid(MainData, OldPosition + 1, NewPosition - OldPosition - 1)
            If NextPageName <> "TREE" Then
                Indent = Len(NextPageName) - Len(LTrim(NextPageName))
                NextPageName = Trim(NextPageName)
                'This positions the item in the tree at the right level
                If Indent = 1 Or Indent = 0 Then
                    LastPageName = "Key" & Pages.Nodes.Count + 1
                    Pages.Nodes.Add , , LastPageName, NextPageName, 1
                    TreeList(Indent) = LastPageName
                Else
                    LastPageName = "Key" & Pages.Nodes.Count + 1
                    Pages.Nodes.Add TreeList(Indent - 1), 4, LastPageName, NextPageName, 1
                    Pages.Nodes(LastPageName).Image = 2
                    TreeList(Indent) = LastPageName
                End If
            End If
        Loop Until NextPageName = "TREE"
    Else
        'If your ARE in edit mode then read every single page
        'Set up a ForNext loop to run through each page
        For PageNumber = 1 To DB.PageCount
            'Add the name of the page to the list box
            NextPageName = DB.PageName(PageNumber)
            Pages.Nodes.Add , , , NextPageName
            'Complete the ForNext Loop
        Next PageNumber
    End If
    Pages.Visible = True
End Sub

Private Sub cmdAddPage_Click()
    'Declare the varable to input the new name to as a string
    Dim NewName As String
    'Get the new name using an input box
    NewName = InputBox("What is the name of this new page?")
    'If you didn't enter a name, exit this prodecure
    If NewName = "" Then Exit Sub
    'Use the CreatePage function to create the page
    DB.CreatePage NewName, "New Page"
    'Update the list of names in the list box
    UpdateListBox
End Sub

Private Sub cmdRemove_Click()
    If MsgBox("Are you sure you want to remove this page?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    'Remove the page that is currently highlighed in the list box
    DB.RemovePage Pages.SelectedItem.Text
    'Updaye the list of names in the list box
    UpdateListBox
End Sub

Private Sub cmdRename_Click()
    'Declare the varable to input the new name to as a string
    Dim NewName As String
    'Get the new name using an input box
    NewName = InputBox("What is the name of this new page?", , Pages.SelectedItem.Text)
    'If you didn't enter a name, exit this prodecure
    If NewName = "" Then Exit Sub
    'Set the page name of the selected page to the new string inputted by the user
    DB.PageName(Pages.SelectedItem.Index) = NewName
    'Updaye the list of names in the list box
    UpdateListBox
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X > Scroller.Left And X < Scroller.Left + Scroller.Width Then
        Dragging = True
        Scroller.BackColor = &H80000010
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Dragging = True Then Scroller.Left = X - (Scroller.Width / 2)
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Dragging = True Then
        Dragging = False
        PageDetail.Left = Scroller.Left + Scroller.Width
        Tools.Left = Scroller.Left + Scroller.Width
        Pages.Width = Scroller.Left - Pages.Left
        PageDetail.Width = ScaleWidth - PageDetail.Left - 50
        Scroller.BackColor = BackColor
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'This asks you if you really want to close the program or not
    If ConfirmExit = True Then If MsgBox("Are you sure you want to quit?", vbYesNo + vbQuestion, Caption) = vbNo Then Cancel = 1
End Sub

Private Sub Form_Resize()
    'This simply arranges all the screen objects to fit on the the form no matter what
    'size or shape the form is. You've seen it before...
    On Error Resume Next
    Pages.Height = ScaleHeight - (Pages.Top * 2)
    PageDetail.Width = ScaleWidth - Pages.Width - (Pages.Left * 3)
    PageDetail.Left = Pages.Width + (Pages.Left * 2)
    If EditMode = True Then
        PageDetail.Height = Pages.Height - Tools.Height
    Else
        PageDetail.Height = Pages.Height
    End If
    Tools.Top = ScaleHeight - Pages.Top - Tools.Height
    Tools.Left = PageDetail.Left
    Scroller.Height = ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'This saves the current page before you close the program
    If EditMode = True Then DB(PageOn) = PageDetail
    Set DB = Nothing
End Sub

Private Sub lstColour_Click()
    'When you click on an colour in the list, this sets the selected text to be that colour
    Select Case LCase(lstColour.Text)
        Case "black": PageDetail.SelColor = vbBlack
        Case "silver": PageDetail.SelColor = vbSilver
        Case "red": PageDetail.SelColor = vbRed
        Case "green": PageDetail.SelColor = vbGreen
        Case "blue": PageDetail.SelColor = vbBlue
        Case "orange": PageDetail.SelColor = vbOrange
        Case "purple": PageDetail.SelColor = vbPurple
    End Select
End Sub

Private Sub lstFontSize_Click()
    'When you select a number in the list box, you set the selected text to display at the selected size
    PageDetail.SelFontSize = lstFontSize
End Sub

Private Sub opTextPos_Click(Index As Integer)
    'This sets the text alignment when you click one of the three buttons
    PageDetail.SelAlignment = Index
End Sub

Private Sub PageDetail_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'When you click on somehwhere on the text window, the controls are updated with the
    'text settings
    Dim n As Integer
    On Error Resume Next
    cmdFonts.Text = ""
    cmdFonts.Text = PageDetail.SelFontName
    If PageDetail.SelBold = True Then ckStyle(0) = 1 Else ckStyle(0) = 0
    If PageDetail.SelUnderline = True Then ckStyle(1) = 1 Else ckStyle(1) = 0
    If PageDetail.SelItalic = True Then ckStyle(2) = 1 Else ckStyle(2) = 0
    If PageDetail.SelBullet = True Then ckStyle(3) = 1 Else ckStyle(3) = 0
    opTextPos(PageDetail.SelAlignment) = True
    Select Case PageDetail.SelColor
        Case vbBlack: lstColour.ListIndex = 0
        Case vbSilver: lstColour.ListIndex = 1
        Case vbRed: lstColour.ListIndex = 2
        Case vbGreen: lstColour.ListIndex = 3
        Case vbBlue: lstColour.ListIndex = 4
        Case vbOrange: lstColour.ListIndex = 5
        Case vbPurple: lstColour.ListIndex = 6
    End Select
    For n = 0 To lstFontSize.ListCount - 1
        If lstFontSize.List(n) = PageDetail.SelFontSize Then lstFontSize.ListIndex = n
    Next n
End Sub

Private Sub Pages_AfterLabelEdit(Cancel As Integer, NewString As String)
    'Declare the varable to input the new name to as a string
    Dim NewName As String
    'Get the new name using an input box
    NewName = NewString
    'If you didn't enter a name, exit this prodecure
    If NewName = "" Then Exit Sub
    'Set the page name of the selected page to the new string inputted by the user
    DB.PageName(Pages.SelectedItem.Index) = NewName
    'Updaye the list of names in the list box
    UpdateListBox
End Sub

Private Sub Pages_Click()
    'This saves whatever is in the text window to the database, and loads the next window
    If Pages.Nodes.Count > 0 Then
        If EditMode = True Then DB(PageOn) = PageDetail
        PageOn = Pages.SelectedItem.Text
        PageDetail = DB(PageOn)
    End If
End Sub

