VERSION 5.00
Begin VB.UserControl Gallary 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   4410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   EditAtDesignTime=   -1  'True
   KeyPreview      =   -1  'True
   ScaleHeight     =   294
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   304
   Begin Project1.Tablet Item 
      Height          =   1575
      Index           =   0
      Left            =   360
      TabIndex        =   3
      Top             =   -3720
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   2143
   End
   Begin VB.FileListBox LoadFile 
      Height          =   1260
      Left            =   720
      OLEDragMode     =   1  'Automatic
      Pattern         =   "*.cpy"
      TabIndex        =   1
      Top             =   2880
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.VScrollBar UpdownBAR 
      Height          =   1215
      Left            =   2520
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "This gallary is currently empty"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   2400
      Width           =   3855
   End
End
Attribute VB_Name = "Gallary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'#########################################################################################
'#                                                                                       #
'#                                Gallary Display Control                                #
'#                                                                                       #
'# This control displays a list of Cpy files that are in a folder as a serise of images. #
'#  The contents of each file is drawn in its own window, then the wndows are arranged   #
'# on the control so that you can browse though them and pick the one you want. You can  #
'#  then drag the window with the mouse onto another control with OLEDrop mode set to    #
'#  automatic. This allows you to drag objects into your files by draging the pictures   #
'#                                                                                       #
'#########################################################################################

Public DraggedItemName As String    'This holds the name of the item dragged
Public SelectedItem As Integer     'This holds the index of the selected item
Private lImageCount As Integer      'This holds the total number of windows displayed
Private lImagePerRow As Integer     'This holds the number of images per row
Const cstItemSpace = 5              'This is the gap between each window
Const cstShadowOfset = 2            'This is the distance that the window shadow is moved from the window itself
Event DropObject(Item As Integer, X As Integer, Y As Integer)   'Called before a drag opperation
Event DragObject(Item As Integer, X As Integer, Y As Integer)   'Called after a drag opperation

Public Property Let GallaryViewMode(Mode As Integer)
    'This allows you to Set the view mode that the gallary is in. When you change modes,
    'each window is redrawn with the new view mode.
    Dim n As Integer
    For n = 1 To lImageCount
        Item(n).ViewMode = Mode
        Am8(Item(n).FileKey).SelectAll
        Am8(Item(n).FileKey).FindModelOutline
        Item(n).ZoomLevel = Item(n).ZoomToSelected * 0.75
        Item(n).CenterView
        Am8(Item(n).FileKey).DeselectAll
        Item(n).Refresh
    Next n
End Property

Public Property Let FolderLocation(ByVal FolderName As String)
    'This sets the name and path of the folder that the control is looking in
    LoadFile.Path = FolderName
    LoadFile.Refresh
    lImageCount = LoadFile.ListCount
    If lImageCount = 0 Then Label1.Visible = True Else Label1.Visible = False
    RefreshItemGrid , 1
End Property

Private Sub item_MouseDown(Index As Integer, X As Single, Y As Single, Button As Integer, Shift As Integer)
    'This sets the item clicked to become selected, redraws the border around the selected
    'item, and calls the DragObject event. This may be useful, it may not.
    SelectedItem = Index
    RefreshItemGrid
    LoadFile.ListIndex = Index - 1
    DraggedItemName = LoadFile.List(LoadFile.ListIndex)
    LoadFile.OLEDrag
    If Button = 2 Then
        PopupMenu frmMain.menuPopupGallary
    Else
        If DirectXNotAvaliable = False Then
            frmMain.ShowGallary.AssignDXEngineTo Am8("GallaryItem" & Index)
            frmMain.ShowGallary.RefreshModel
            frmMain.ShowGallary.SetMode 4
        End If
    End If
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    'This allows you to move through the visible windows using the arrow keys
    If KeyCode = vbKeyUp Then SelectedItem = SelectedItem - ImagePerRow
    If KeyCode = vbKeyLeft Then SelectedItem = SelectedItem - 1
    If KeyCode = vbKeyRight Then SelectedItem = SelectedItem + 1
    If KeyCode = vbKeyDown Then SelectedItem = SelectedItem + ImagePerRow
    If SelectedItem < 1 Then SelectedItem = 1
    If SelectedItem > lImageCount Then SelectedItem = lImageCount
    RefreshItemGrid
End Sub

Public Sub RefreshItemGrid(Optional ForceScrollBar As Integer = 0, Optional ForceRedraw As Integer = 0)
    'This is the sub that does the actual work. It checks to see if the number
    'of windows has altered, loading and removing them as required.
    'It then aligns the new windows into rows and columns. It also moves the
    'objects up and down to take into acount moving the scroll bar, and can
    'set the scrollbar min and max values to display all of the windows
    Dim ReLoadObjects As Boolean, iHeight As Integer, iWidth As Integer
    Dim ItemWidth As Integer, ItemHeight As Integer, n As Integer
    Dim NewKey As String, OldActive As String
    OldActive = ActiveFile
    If ForceRedraw = 1 Then ReLoadObjects = True: LoadFile.Refresh
    lImageCount = LoadFile.ListCount
    If lImageCount <> Item.Count - 1 Then ReLoadObjects = True
    ItemHeight = Item(0).Height
    If (ItemHeight + cstItemSpace) * ((lImageCount - 1) / lImagePerRow) > ScaleHeight - (ItemHeight / 5) Then
        UpdownBAR.Visible = True
        ItemWidth = ((ScaleWidth - UpdownBAR.Width) / ImagePerRow) - cstItemSpace
    Else
        UpdownBAR.Visible = False
        ItemWidth = ((ScaleWidth) / ImagePerRow) - cstItemSpace
    End If
    If ReLoadObjects = True Then
        For n = Item.Count - 1 To 1 Step -1
            Am8.File.Remove "GallaryItem" & n
            Unload Item(n)
        Next n
    End If
    Cls
    For n = 1 To lImageCount
        If ReLoadObjects = True Then
            Load Item(n)
            NewKey = "GallaryItem" & n
            ActiveFile = NewKey
            Am8.File.Add NewKey
            Item(n).AssignTabletTo Am8(NewKey)
            Item(n).ViewMode = 2
            Item(n).SetBorderStyle 0
            Item(n).SetScrollBarStyle 0
            Item(n).Visible = True
            Item(n).ToolTipText = RightClip(LoadFile.List(n - 1), 4)
        End If
        Item(n).Width = ItemWidth
        Item(n).Height = ItemHeight
        Item(n).Left = (Item(n).Width + cstItemSpace) * iWidth
        Item(n).Top = ((Item(n).Height + cstItemSpace) * iHeight) - UpdownBAR
        iWidth = iWidth + 1
        If iWidth = ImagePerRow Then iWidth = 0: iHeight = iHeight + 1
        If ReLoadObjects = True Then
            If Am8(NewKey).LoadFromFile(LoadFile.Path & "\" & LoadFile.List(n - 1), 2) = False Then
                Item(n).ToolTipText = Item(n).ToolTipText & " Failed to Load - Please Remove or fix"
            End If
            Am8(NewKey).SelectAll
            Am8(NewKey).FindModelOutline
            Am8(NewKey).Saved = True
            Am8(NewKey).NonEditableFile = True
            Am8(NewKey).Scene.UpdateAllScenes
            Am8(NewKey).MorphSkeliton "BaseFrame", "BaseFrame"
            Item(n).ZoomLevel = Item(n).ZoomToSelected * 0.75
            Item(n).CenterView
            Am8(NewKey).DeselectAll
            Item(n).Refresh
        End If
        If n = SelectedItem Then Line (Item(n).Left + cstShadowOfset, Item(n).Top + cstShadowOfset)-(Item(n).Left + Item(n).Width + cstShadowOfset, Item(n).Top + Item(n).Height + cstShadowOfset), RGB(100, 100, 100), BF
    Next n
    If ReLoadObjects = True Or ForceScrollBar = 1 Then
        If lImageCount > 0 And DirectXNotAvaliable = False Then
            frmMain.ShowGallary.AssignDXEngineTo Am8("GallaryItem1")
            frmMain.ShowGallary.RefreshModel
            frmMain.ShowGallary.SetMode 4
        End If
        UpdownBAR.Max = Item(lImageCount).Top - cstItemSpace
        UpdownBAR.LargeChange = Item(lImageCount).Height + cstItemSpace
        UpdownBAR.SmallChange = Item(lImageCount).Height + cstItemSpace
    End If
    ActiveFile = OldActive
    Refresh
End Sub

Private Sub UserControl_Initialize()
    'This sets up the standard values etc
    ImagePerRow = 2
    SelectedItem = 1
    RefreshItemGrid 1
End Sub

Private Sub UserControl_Resize()
    'This resizes the scrollbar to fit down the side of the control, and redisplays the objects
    UpdownBAR.Height = ScaleHeight
    UpdownBAR.Left = ScaleWidth - UpdownBAR.Width
    Label1.Width = ScaleWidth
    RefreshItemGrid 1
End Sub

Private Sub Item_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'This event is called when you release the mouse button. Checking whether the
    'object was dropped on the right control cannot be decided from within the control
    RaiseEvent DropObject(Index, (X + Item(Index).Left), (Y + Item(Index).Top + 20))
End Sub

Public Property Get FolderLocation() As String
    'This returns the name of the folder the control is looking at
    FolderLocation = LoadFile.Path
End Property

Public Function GetItemName(Index As Integer) As String
    'This returns the name of a window
    GetItemName = Item(Index).ToolTipText
End Function

Public Function ImageCount() As Integer
    'This returns the local image count value
    ImageCount = lImageCount
End Function

Public Property Get ImagePerRow() As Integer
    'This returns the value of lImagePerRow
    ImagePerRow = lImagePerRow
End Property

Public Property Let ImagePerRow(ByVal vNewValue As Integer)
    'This sets the value of lImagePerRow
    lImagePerRow = vNewValue
    RefreshItemGrid
End Property

Private Sub UpdownBAR_Change()
    'When you CLICK the scroll bar, this repositions each of the windows
    RefreshItemGrid
End Sub

Private Sub UpdownBAR_Scroll()
    'When you SLIDE the scroll bar, this repositions each of the windows
    RefreshItemGrid
End Sub
