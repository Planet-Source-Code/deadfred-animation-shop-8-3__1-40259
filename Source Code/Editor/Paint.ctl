VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl PaintWindow 
   BackColor       =   &H8000000C&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8175
   FillStyle       =   7  'Diagonal Cross
   ScaleHeight     =   433
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   545
   Begin VB.CommandButton cmdBlock 
      Enabled         =   0   'False
      Height          =   255
      Left            =   7920
      TabIndex        =   2
      Top             =   6240
      Width           =   255
   End
   Begin MSComCtl2.FlatScrollBar fBar 
      Height          =   6255
      Index           =   1
      Left            =   7920
      TabIndex        =   4
      Top             =   0
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   11033
      _Version        =   393216
      MousePointer    =   1
      Appearance      =   0
      LargeChange     =   100
      Min             =   -100
      Max             =   1
      Orientation     =   8323072
      SmallChange     =   100
   End
   Begin MSComCtl2.FlatScrollBar fBar 
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   6240
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   450
      _Version        =   393216
      MousePointer    =   1
      Appearance      =   0
      Arrows          =   65536
      LargeChange     =   100
      Min             =   -100
      Max             =   1
      Orientation     =   8323073
      SmallChange     =   100
   End
   Begin VB.PictureBox TexMap 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H008080FF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   5535
      Left            =   840
      MousePointer    =   2  'Cross
      ScaleHeight     =   369
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   449
      TabIndex        =   1
      Top             =   480
      Width           =   6735
      Begin VB.Line gGuide 
         Visible         =   0   'False
         X1              =   40
         X2              =   41
         Y1              =   32
         Y2              =   33
      End
   End
   Begin VB.PictureBox PicStore 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   1440
      ScaleHeight     =   1
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1
      TabIndex        =   0
      Top             =   2040
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.PictureBox Shadow 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      FillColor       =   &H008080FF&
      FillStyle       =   5  'Downward Diagonal
      ForeColor       =   &H80000008&
      Height          =   5535
      Left            =   720
      ScaleHeight     =   369
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   449
      TabIndex        =   5
      Top             =   360
      Width           =   6735
      Begin VB.Line Line1 
         Visible         =   0   'False
         X1              =   40
         X2              =   41
         Y1              =   32
         Y2              =   33
      End
   End
End
Attribute VB_Name = "PaintWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'##################################################################################################
'#                                                                                                #
'#                                   Paint Brush Control Thing                                    #
'#                                                                                                #
'#   This control allows you to create simple bitmaps. Its used so that you can create and edit   #
'#   the texture maps for your models. You can load images from other software, which you         #
'#   would probably like better, but this is here for quick edits. It is pretty good for what     #
'#   it is, with shape previews as you create them, paint tools, insert tools and shape tools.    #
'#                                                                                                #
'##################################################################################################

Private Model As clsFile

Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function FloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Public pDrawShapes As Boolean, DrawMode As Integer
Private lForeColour As Long, lBackColour As Long
Private lPictureFile As String, lDrawWidth As Integer
Private Canceler As Boolean, API As POINTAPI
Const BorderSize = 50



Public Property Get DrawShapes() As Boolean
    DrawShapes = pDrawShapes
End Property

Public Property Let DrawShapes(ByVal vNewValue As Boolean)
    pDrawShapes = vNewValue
    If Model Is Nothing Then Else DrawObjects True
End Property



Private Sub DrawObjects(Optional AlwaysDraw As Boolean = False)
    'Draws the objects texture map outline
    Dim Am As clsObject, fm As clsFace, n As Integer
    Dim Coord(25, 2) As Integer
    If pDrawShapes = True Or AlwaysDraw = True Then
        TexMap.ForeColor = vbBlack
        TexMap.DrawWidth = 1
        TexMap.DrawMode = 5
        For Each Am In Model.Geometery
            If Am.TexVert.Count > 0 Then
                For Each fm In Am.Face
                    For n = 1 To fm.EdgeCount
                        Coord(n, 1) = Am.TexVert(fm.Edge(n).TexVertex).Y
                        Coord(n, 2) = Am.TexVert(fm.Edge(n).TexVertex).X
                    Next n
                    MoveTo Coord(fm.EdgeCount, 1), Coord(fm.EdgeCount, 2)
                    For n = 1 To fm.EdgeCount: DrawTo Coord(n, 1), Coord(n, 2): Next n
                Next fm
            End If
        Next Am
        TexMap.ForeColor = lForeColour
        TexMap.DrawWidth = lDrawWidth
        TexMap.Refresh
        TexMap.DrawMode = 13
    End If
End Sub

Private Sub TexMap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'When you click down on the image, this sets the image so that you can draw the shape previews onto it without
    'going over the original image. It also sets the start of the guide line to when you clicked the mouse
    DrawObjects
    If Button = 1 Then
        Select Case DrawMode
            Case 1, 2, 3, 7: gGuide.x1 = X: gGuide.y1 = Y: TexMap.AutoRedraw = False: TexMap.DrawMode = 10
            Case 5: TexMap.AutoRedraw = True: FloodFill TexMap.hdc, X, Y, lForeColour: TexMap.Refresh
            Case 6: If PicStore.Picture <> 0 Then TexMap.PaintPicture PicStore.Picture, X - (PicStore.ScaleWidth / 2), Y - (PicStore.ScaleHeight / 2)
        End Select
    End If
End Sub

Private Sub TexMap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'This runs when you are moving the mouse around. Most of these draw shape previews, which dont actually get
    'added to the final image, but the spray paint, and freehand line is drawn here
    Static Ox As Single, Oy As Single, n As Integer
    If Button = 1 Then
        If TexMap.AutoRedraw = True And DrawMode <> 4 And DrawMode <> 6 And DrawMode <> 8 Then TexMap.AutoRedraw = False
        Select Case DrawMode
            Case 1: gGuide.x2 = X: gGuide.y2 = Y: TexMap.Cls: TexMap.Line (gGuide.x1, gGuide.y1)-(X, Y), , B
            Case 2: gGuide.x2 = X: gGuide.y2 = Y: TexMap.Cls: TexMap.Circle (gGuide.x1, gGuide.y1), Dist(gGuide.x1 - X, gGuide.y1 - Y), ForeColour
            Case 3: gGuide.x2 = X: gGuide.y2 = Y: TexMap.Cls: Ellipse TexMap.hdc, gGuide.x1, gGuide.y1, X, Y
            Case 4: TexMap.Line (X, Y)-(Ox, Oy), ForeColour
            Case 6: If PicStore.Picture <> 0 Then TexMap.PaintPicture PicStore.Picture, X - (PicStore.ScaleWidth / 2), Y - (PicStore.ScaleHeight / 2)
            Case 7: TexMap.Cls: TexMap.Line (gGuide.x1, gGuide.y1)-(X, Y), ForeColour
            Case 8: For n = 1 To 8: TexMap.PSet (X + ((Rnd * 10) - 5), Y + ((Rnd * 10) - 5)), ForeColour: Next n
        End Select
        Canceler = False
    End If
    If Button = 2 Then
        Select Case DrawMode
            Case 4: TexMap.Line (X, Y)-(Ox, Oy), BackColour
            Case 8: For n = 1 To 8: TexMap.PSet (X + ((Rnd * 10) - 5), Y + ((Rnd * 10) - 5)), BackColour: Next n
        End Select
    End If
    Ox = X: Oy = Y
End Sub

Private Sub TexMap_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'When you release the mouse, the shapes are drawn for real here
    Static Ox As Single, Oy As Single
    If Button = 1 Then
        If Canceler = False Then
            Select Case DrawMode
                Case 1: TexMap.AutoRedraw = True: TexMap.DrawMode = 13: TexMap.Line (gGuide.x1, gGuide.y1)-(X, Y), ForeColour, B
                Case 2: TexMap.AutoRedraw = True: TexMap.DrawMode = 13: TexMap.Circle (gGuide.x1, gGuide.y1), Dist(gGuide.x1 - X, gGuide.y1 - Y), ForeColour
                Case 3: TexMap.AutoRedraw = True: TexMap.DrawMode = 13: Ellipse TexMap.hdc, gGuide.x1, gGuide.y1, X, Y: TexMap.Refresh
                Case 4, 8: TexMap.AutoRedraw = True: TexMap.DrawMode = 13
                Case 7: TexMap.AutoRedraw = True: TexMap.DrawMode = 13: TexMap.Line (gGuide.x1, gGuide.y1)-(X, Y), ForeColour
            End Select
        End If
        TexMap.DrawMode = 13: TexMap.DrawStyle = 0: Canceler = False
    End If
    If Button = 2 Then
        Cls
        TexMap.AutoRedraw = True: TexMap.Refresh: Canceler = True
    End If
    Ox = X: Oy = Y
    If Model Is Nothing Then Else DrawObjects
End Sub

Private Sub ResizeForm()
    'When you resize the form, this moves the scrollbars to fit around the edit of the screen. It then adjusts the values
    'of the scroll bars so that they allow you to scroll the entire texture map across the screen. It also hides the scroll
    'bars if the entire image fits on the screen without needing to scroll the image.
    Dim n As Integer, PosX As Integer, PosY As Integer
    fBar(1).Left = ScaleWidth - fBar(1).Width
    fBar(0).Top = ScaleHeight - fBar(0).Height
    fBar(1).Height = ScaleHeight - fBar(0).Height
    fBar(0).Width = ScaleWidth - fBar(1).Width
    cmdBlock.Left = fBar(1).Left
    cmdBlock.Top = fBar(0).Top
    fBar(0).Min = -BorderSize
    fBar(1).Min = -BorderSize
    fBar(0).Max = (TexMap.Width - ScaleWidth) + BorderSize + fBar(1).Width
    fBar(1).Max = (TexMap.Height - ScaleHeight) + BorderSize + fBar(1).Width
    For n = 0 To 1
        If fBar(n).Max < 0 Then
            fBar(n).Visible = False
            fBar(n) = -BorderSize
            PosX = (ScaleWidth / 2) - (TexMap.ScaleWidth / 2): If PosX < BorderSize Then PosX = BorderSize
            PosY = (ScaleHeight / 2) - (TexMap.ScaleHeight / 2): If PosY < BorderSize Then PosY = BorderSize
            If n = 0 Then TexMap.Left = PosX
            If n = 1 Then TexMap.Top = PosY
        Else
            fBar(n).Visible = True
            fBar(n).SmallChange = (fBar(n).Max - fBar(n).Min) / 10
            fBar(n).LargeChange = (fBar(n).Max - fBar(n).Min) / 4
            If n = 1 Then TexMap.Top = -fBar(1) Else TexMap.Left = -fBar(0)
        End If
    Next n
    If fBar(0).Visible = False And fBar(1).Visible = False Then cmdBlock.Visible = False Else cmdBlock.Visible = True
    Shadow.Move TexMap.Left + 5, TexMap.Top + 5, TexMap.Width, TexMap.Height

End Sub

Public Sub TileImage()
    'This routine takes an image and places it repeatedly over the image, filling it up. It works best with wallpaper
    'type images, which blend together so you cant see the edhes
    Dim n As Integer, m As Integer
    For n = 0 To TexMap.ScaleWidth Step PicStore.ScaleWidth
        For m = 0 To TexMap.ScaleHeight Step PicStore.ScaleHeight
            TexMap.PaintPicture PicStore.Picture, n, m
        Next m
    Next n
End Sub

Public Function GetCurrentHDC() As Long
    'This returns the hDc of the control
    GetCurrentHDC = TexMap.hdc
End Function

Private Sub TexMap_DblClick()
    'This ensures that even if you double click, the mouse_up event fires, which avoids errors
    TexMap_MouseUp 1, 0, 0, 0
End Sub

Private Sub UserControl_Initialize()
    'This sets up default values for the control
    ForeColour = vbBlue
    BackColour = vbWhite
    DrawMode = 4
    lDrawWidth = 1
    fBar_Change 1
End Sub

Private Sub UserControl_Resize()
    'This runs the resize code
    ResizeForm
End Sub

Public Sub Refresh()
    'Draws the texture faces
    If pDrawShapes = True Then If Model Is Nothing Then Else DrawObjects
End Sub

Public Function MoveTo(x1 As Integer, y1 As Integer)
    'This function uses the MoveTo command on its own, to set the position of the cursor on the screen.
    MoveToEx TexMap.hdc, x1, y1, API
End Function

Public Function DrawTo(x1 As Integer, y1 As Integer)
    'This uses the LineTo command to draw from where ever the cursor was before, to
    'a new position, which is faster than having to set the cursor every time.
    LineTo TexMap.hdc, x1, y1
End Function

Private Function Dist(X As Single, Y As Single) As Single
    'This little function calculates the hypotonuse of a right angled triangle
    Dist = Sqr((X ^ 2) + (Y ^ 2))
End Function

Public Sub AssignTexmapTo(AssignedModel As clsFile)
    'Sets the control to point to an existing file
    Set Model = AssignedModel
End Sub

Public Sub ClearImage()
    'This sub allows you to clear the form
    TexMap.Cls
    TexMap.Picture = Nothing
End Sub

Public Sub LoadImage(FileName As String)
    'This sub allows you to load a bitmap or jpg file as the image
    TexMap.AutoRedraw = True
    TexMap.Picture = LoadPicture(FileName)
    ResizeForm
End Sub

Public Function SaveImage(FileName As String) As Boolean
    'This allows you to save the current image as the given file
    TexMap.AutoRedraw = True
    SavePicture TexMap.Image, FileName
End Function

Private Sub fBar_Change(Index As Integer)
    'When you change the value of the window scroll bars by clicking on them, this alters the position of the image
    If Index = 1 Then TexMap.Top = -fBar(1) Else TexMap.Left = -fBar(0)
    Shadow.Move TexMap.Left + 5, TexMap.Top + 5, TexMap.Width, TexMap.Height
    Refresh
End Sub

Private Sub fBar_Scroll(Index As Integer)
    'When you change the value of the window scroll bars by dragging them around, this alters the position of the image
    If Index = 1 Then TexMap.Top = -fBar(1) Else TexMap.Left = -fBar(0)
    Shadow.Move TexMap.Left + 5, TexMap.Top + 5, TexMap.Width, TexMap.Height
    Refresh
End Sub

Public Property Get FillPattern() As Long
    'This property sets or returns the fill pattern, which is used to fill in objects. You can
    'have stripy or solid and other patterns
    FillPattern = TexMap.FillStyle
End Property
Public Property Let FillPattern(ByVal NewMode As Long)
    TexMap.FillStyle = NewMode
End Property

Public Property Get ForeColour() As Long
    'This property sets or returns the fore colour.
    ForeColour = lForeColour
End Property
Public Property Let ForeColour(ByVal Colour As Long)
    lForeColour = Colour
    TexMap.ForeColor = Colour
End Property

Public Property Get BackColour() As Long
    'This property sets or returns the back colour
    BackColour = lBackColour
End Property
Public Property Let BackColour(ByVal Colour As Long)
    lBackColour = Colour
    TexMap.FillColor = Colour
End Property

Public Property Get PictureFileName() As Variant
    'This property allows you to set or reteive the current image file name that has been loaded to
    'paste into the texture, or tile over the texture
    PictureFileName = lPictureFile
End Property
Public Property Let PictureFileName(ByVal vNewValue As Variant)
    lPictureFile = vNewValue
    PicStore.Picture = LoadPicture(lPictureFile)
End Property

Public Property Get LineWidth() As Integer
    'This property allows you to set and retrive the current line width that is being drawn
    LineWidth = TexMap.DrawWidth
End Property
Public Property Let LineWidth(ByVal NewLineWidth As Integer)
    TexMap.DrawWidth = NewLineWidth
    lDrawWidth = NewLineWidth
End Property

Public Property Get TexHeight() As Integer
    'This property allows you to get or set the height of the texture map
    TexHeight = TexMap.ScaleHeight
End Property
Public Property Let TexHeight(ByVal vNewValue As Integer)
    TexMap.Height = vNewValue
    ResizeForm
End Property

Public Property Get TexWidth() As Integer
    'This property allows you to get or set the width of the texture map
    TexWidth = TexMap.ScaleWidth
End Property
Public Property Let TexWidth(ByVal vNewValue As Integer)
    TexMap.Width = vNewValue
    ResizeForm
End Property



