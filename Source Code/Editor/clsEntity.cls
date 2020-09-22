VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEntity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsProperty"
Attribute VB_Ext_KEY = "Member0" ,"clsProperty"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public EntityName As String
Public Discription As String

Private AllProps As Collection


Public Function Add(PropertyName As String, Optional PropertyType As String = "", Optional ToolTip As String = "", Optional Choice As String = "", Optional Default As Variant = "", Optional ObjectWidth As Integer = 2400, Optional Locked As Boolean = False, Optional NewColumn As Boolean = False, Optional NewPage As Boolean = False) As clsProperty
    'This is the Add Property function, which allows you to set all the attributes in the function call
    'line, so the rest of the program is much tidyer
    On Error GoTo FailedToLoadProperty
    Dim NewProperty As clsProperty
    Set NewProperty = New clsProperty
    If PropertyType <> "" Then NewProperty.PropertyType = PropertyType
    If ToolTip <> "" Then NewProperty.ToolTip = ToolTip
    If Choice <> "" Then NewProperty.Choice = Choice
    If Default <> "" Then NewProperty.Default = Default
    NewProperty.NewColumn = NewColumn
    NewProperty.NewPage = NewPage
    NewProperty.Width = ObjectWidth
    NewProperty.Locked = Locked
    NewProperty.PropertyName = PropertyName
    AllProps.Add NewProperty, PropertyName
    Set Add = NewProperty
FailedToLoadProperty:
Set NewProperty = Nothing
End Function


Public Property Get Item(vntIndexKey As Variant) As clsProperty
Attribute Item.VB_UserMemId = 0
  Set Item = AllProps(vntIndexKey)
End Property


Public Property Get Count() As Long
    Count = AllProps.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    AllProps.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = AllProps.[_NewEnum]
End Property


Private Sub Class_Initialize()
    Set AllProps = New Collection
End Sub


Private Sub Class_Terminate()
    Set AllProps = Nothing
End Sub

