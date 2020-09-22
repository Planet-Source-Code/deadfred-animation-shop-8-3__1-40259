VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "txtDataBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public EncriptionKey As String
Public FileName As String




Public Function CreatePage(ByVal PageName As String, Optional ByVal PageContent As String = "") As Boolean
    'This simply creates a new marker point in a file by inserting the given name on
    'to the bottom of the text file
    'PageName = CipherText(PageName)
    Open FileName For Append As #1
        Print #1, "<#" & PageName & "#>"
        If PageContent <> "" Then
            Print #1, CipherText(PageContent)
        End If
    Close
    CreatePage = True
End Function




Public Function RemovePage(ByVal PageName As String) As Boolean
    'This removes a marker and its content from a storage file
    'PageName = CipherText(PageName)
    Dim TmpLine As String
    DataStore(PageName) = ""
    Open FileName For Input As #1
    Open FileName & "tmp" For Output As #2
    Do While EOF(1) = False
        Line Input #1, TmpLine
        If TmpLine <> "<#" & PageName & "#>" Then Print #2, TmpLine
    Loop
    Close
    Kill FileName
    Name FileName & "tmp" As FileName
End Function




Public Property Get PageName(ByVal PageNumber As Integer) As String
    'This returns the name of a page when you give a page number
    Dim Counter As Integer, TextLine As String
    Open FileName For Input As #1
        Do While EOF(1) = False
            Line Input #1, TextLine
            If Left(TextLine, 2) = "<#" Then
                Counter = Counter + 1
                If Counter = PageNumber Then
                    PageName = Mid(TextLine, 3, Len(TextLine) - 4)
                    Close
                    Exit Property
                End If
            End If
        Loop
    Close
    PageName = "@not_found"
End Property




Public Property Let PageName(ByVal PageNumber As Integer, ByVal vNewValue As String)
    'This allows you to sets the name of the page with the number you supply
    Dim Counter As Integer, TextLine As String
    'vNewValue = CipherText(vNewValue)
    Open FileName For Input As #1
    Open FileName & "tmp" For Output As #2
        Do While EOF(1) = False
            Line Input #1, TextLine
            If Left(TextLine, 2) = "<#" Then
                Counter = Counter + 1
                If Counter = PageNumber Then
                    Print #2, "<#" & vNewValue & "#>"
                Else
                    Print #2, TextLine
                End If
            Else
                Print #2, TextLine
            End If
        Loop
    Close
    Kill FileName
    Name FileName & "tmp" As FileName
End Property




Public Property Get DataStore(ByVal PageName As String) As String
Attribute DataStore.VB_UserMemId = 0
    'This function reads in a section of text from the storage file. The Marker name variable
    'decides which area of the text file to return.
    Dim TempLine As String, TotalContent As String, TextLine As String, ContentString As String
    
    'PageName = CipherText(PageName)
    
    Open FileName For Input As #1
        Do While EOF(1) = False
            Line Input #1, TextLine
            If Left(TextLine, 2) = "<#" Then
                If LCase(Mid(TextLine, 3, Len(TextLine) - 4)) = LCase(PageName) And EOF(1) = False Then
                    Line Input #1, ContentString
                    TotalContent = ContentString & vbNewLine
                    Do While Left(ContentString, 2) <> "<#" And EOF(1) = False
                        Line Input #1, ContentString
                        If Left(ContentString, 2) <> "<#" Then
                            TotalContent = TotalContent & ContentString & vbNewLine
                        End If
                    Loop
                    If EncriptionKey = "" Then
                        DataStore = TotalContent
                    Else
                        DataStore = PlainText(TotalContent)
                    End If
                    Close
                    Exit Property
                End If
            End If
        Loop
    Close
    DataStore = "@not_found - '" & PlainText(PageName) & "'"
End Property




Public Property Let DataStore(ByVal PageName As String, ByVal vNewValue As String)
    'This writes a section of text to the storage file. Again, the MarkerName defines where
    'the given peice of text is saved.
    Dim TextLine As String, ContentString As String, TmpLine As String
    'PageName = CipherText(PageName)
    Open FileName & "tmp" For Output As #1
    Open FileName For Input As #2
    Dim TempLine As String
        Do While EOF(2) = False
            Line Input #2, TextLine
            Print #1, TextLine
            If Left(TextLine, 2) = "<#" Then
            
            
                If LCase(Mid(TextLine, 3, Len(TextLine) - 4)) = LCase(PageName) Then
                    If vNewValue <> "" Then
                        Print #1, CipherText(vNewValue)
                    End If
                    Do While Left(ContentString, 2) <> "<#" And EOF(2) = False
                        Line Input #2, ContentString
                        If Left(ContentString, 2) = "<#" Then Print #1, ContentString
                    Loop
                End If
            End If
        Loop
    Close
    Kill FileName
    Name FileName & "tmp" As FileName
End Property




Private Function CipherText(ByVal TextString As String) As String
    'This takes a text string, and a key, and returns an enciphered string
    Dim n As Integer, KeyPos As Integer, newchar As String, NewText As String
    If EncriptionKey = "" Then CipherText = TextString: Exit Function
    If TextString = "" Then Exit Function
    For n = 1 To Len(TextString)
        KeyPos = KeyPos + 1: If KeyPos = Len(EncriptionKey) Then KeyPos = 0
        newchar = Asc(Mid(TextString, n, 1)) + Asc(Mid(EncriptionKey, KeyPos + 1, 1))
        If newchar > 250 Then newchar = newchar - 200
        NewText = NewText & Chr(newchar)
    Next n
    CipherText = NewText
End Function




Private Function PlainText(ByVal TextString As String) As String
    'This takes a enciphered text string, and a key, and returns an plaintext string
    Dim n As Integer, KeyPos As Integer, KeyValue As String, newchar As String, NewText As String
    If EncriptionKey = "" Then PlainText = TextString: Exit Function
    If TextString = "" Then Exit Function
    For n = 1 To Len(TextString)
        KeyPos = KeyPos + 1: If KeyPos = Len(EncriptionKey) Then KeyPos = 0
        KeyValue = Asc(Mid(EncriptionKey, KeyPos + 1, 1))
        newchar = Asc(Mid(TextString, n, 1)) - KeyValue
        If newchar < 1 Then newchar = newchar + 200
        NewText = NewText & Chr(newchar)
    Next n
    PlainText = Mid(NewText, 1, Len(NewText) - 1)
End Function




Public Property Get PageCount() As Integer
Attribute PageCount.VB_UserMemId = -4
    'This counts the number of pages that are in the current text file
    Dim Counter As Integer, TextLine As String
    On Error GoTo FileNotFound
    Open FileName For Input As #2
    Do While EOF(2) = False
        Line Input #2, TextLine
        If Left(TextLine, 2) = "<#" Then Counter = Counter + 1
    Loop
    PageCount = Counter
    Close
    Exit Property
FileNotFound:
    Open FileName For Output As #1
    Close
    Resume
End Property





