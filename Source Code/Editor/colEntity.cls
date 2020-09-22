VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colEntity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsProperty"
Attribute VB_Ext_KEY = "Member0" ,"clsProperty"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable to hold collection
Private mCol As Collection



Public Function Add(EntityName As String) As clsEntity
    'create a new object
    Dim objNewMember As clsEntity
    Set objNewMember = New clsEntity
    objNewMember.EntityName = EntityName
    mCol.Add objNewMember, EntityName
    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function



Public Property Get Item(vntIndexKey As Variant) As clsEntity
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
    Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub



Sub SaveEntities()
    'This writes the details of all the entities and properties to disk
    Dim Am As clsEntity, Pm As clsProperty
    Open App.Path & "\data\entity.dat" For Output As #1
        For Each Am In Am8.Properties
            Print #1, Am.EntityName
            Print #1, Am.Discription
            Print #1, Am.Count
            For Each Pm In Am
                Print #1, Pm.Choice
                Print #1, Pm.Default
                Print #1, Pm.Locked
                Print #1, Pm.NewColumn
                Print #1, Pm.NewPage
                Print #1, Pm.PropertyName
                Print #1, Pm.PropertyType
                Print #1, Pm.ToolTip
                Print #1, Pm.Width
            Next Pm
        Next Am
    Close
End Sub




Sub LoadEntities()
    'This reads all the entities and properties into memory from a text file
    Dim Am As clsEntity, Pm As clsProperty, EntityName As String, n As Integer
    Dim PMDefault As String, PMWidth As Integer, PMLocked As Boolean, PMNewColumn As Boolean, PMNewPage As Boolean
    Dim PMPropName As String, PMPropType As String, PMToolTip As String, PMChoice As String, PropertyCount As Integer
    Dim SToBool As String, StrDiscription As String
    For n = Am8.Properties.Count To 1 Step -1: Am8.Properties.Remove n: Next n
    'On Error GoTo FailedToLoadFile
    Open App.Path & "\data\entity.dat" For Input As #1
    Do Until EOF(1) = True
        Input #1, EntityName
        Line Input #1, StrDiscription
        Input #1, PropertyCount
        Am8.Properties.Add EntityName
        For n = 1 To PropertyCount
            Line Input #1, PMChoice
            Line Input #1, PMDefault
            Line Input #1, SToBool: If SToBool = "True" Then PMLocked = True Else PMLocked = False
            Line Input #1, SToBool: If SToBool = "True" Then PMNewColumn = True Else PMNewColumn = False
            Line Input #1, SToBool: If SToBool = "True" Then PMNewPage = True Else PMNewPage = False
            Line Input #1, PMPropName
            Line Input #1, PMPropType
            Line Input #1, PMToolTip
            Input #1, PMWidth
            Am8.Properties(EntityName).Add PMPropName, PMPropType, PMToolTip, PMChoice, PMDefault, PMWidth, PMLocked, PMNewColumn, PMNewPage
            Am8.Properties(EntityName).Discription = StrDiscription
        Next n
    Loop
    Close
FailedToLoadFile:
End Sub
