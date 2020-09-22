Attribute VB_Name = "modImport"
Option Explicit

'#####################################################################
'#                                                                   #
'# This function allows files to be imported to memory. Its basicly  #
'#    just openning files, but of a different format. This should    #
'#     allow you to inport compiled files made with this program,    #
'#  MAP files, or ASC files, which are a more common file type that  #
'# some other programs support, so there is some way of moving files #
'#                 from other programs to this one                   #
'#                                                                   #
'#####################################################################


Function ImportModel(FileName As String, ActiveFile As clsFile) As Boolean
    
    Dim NewKey As String, Num As Integer, EdgeCount As Integer, Poss(12) As Integer
    Dim x As Single, y As Single, z As Single, FaceON As Integer, Edge As Integer
    Dim n As Integer, m As Integer, LineON As String, VertexCount As Integer
    Dim FileType As String, Temp As String, FaceCount As Integer
    Dim Xx1 As Integer, Yy1 As Integer, Zz1 As Integer
    Dim Xx2 As Integer, Yy2 As Integer, Zz2 As Integer
    Dim Xx3 As Integer, Yy3 As Integer, Zz3 As Integer
    
    On Error GoTo FailedToImport
    
    FileType = LCase(Right(FileName, 3))
    If FileType = "cpy" Or FileType = "am8" Then
        ActiveFile.LoadFromFile FileName, 1
        ImportModel = True
        Exit Function
    End If
    
    Open FileName For Input As #1:
    NewKey = "Import" & Timer & Rnd
    ActiveFile.Geometery.CreateObject NewKey
    
    With ActiveFile.Geometery(NewKey)
        Select Case FileType
    
            '============================================================================================
            'Inport compiled files from previous versions of AM
            Case "dat"
                Line Input #1, Temp
                Line Input #1, Temp
                Line Input #1, Temp
                Line Input #1, Temp
                Input #1, VertexCount
                Input #1, FaceCount
                For n = 1 To VertexCount
                    Input #1, Xx1, Yy1, Zz1, Num
                    .Vertex.Add Xx1, Yy1, Zz1
                Next n
                For n = 1 To FaceCount
                    Input #1, EdgeCount
                    .Face.Add EdgeCount
                    Input #1, Edge
                    .Face(n).Edge.Add Edge + 1
                    For m = 2 To EdgeCount
                        Input #1, Edge
                        .Face(n).Edge.Insert 1, Edge + 1
                    Next m
                Next n
            '============================================================================================


            '============================================================================================
            'Load raw text tri-mesh models
            Case "txt"
                Input #1, Temp
                Do
                    Input #1, Xx1:    Input #1, Yy1:     Input #1, Zz1
                    Input #1, Xx2:    Input #1, Yy2:     Input #1, Zz2
                    Input #1, Xx3:    Input #1, Yy3:     Input #1, Zz3
                    .Vertex.Add Xx1, -Zz1, Yy1
                    .Vertex.Add Xx2, -Zz2, Yy2
                    .Vertex.Add Xx3, -Zz3, Yy3
                    .Face.Add 3, .Vertex.Count - 2, .Vertex.Count - 1, .Vertex.Count
                Loop While EOF(1) = False
            '============================================================================================
 

            '============================================================================================
            'Load a hight matrix text file. Check help for specifications of this
            Case "map"
                Input #1, Xx1, Yy1
                For n = 1 To Xx1
                    For m = 1 To Yy1
                        Input #1, Num
                        .Vertex.Add (n * 10) - (Xx1 * 5) - 5, -Num, (m * 10) - (Yy1 * 5) - 5
                    Next m
                Next n
                For n = 1 To Xx1 - 1
                    For m = 1 To Yy1 - 1
                        .Face.Add 3, m + ((n - 1) * Xx1), m + ((n - 1) * Xx1) + 1, m + ((n - 1) * Xx1) + Xx1
                        .Face.Add 3, m + ((n - 1) * Xx1) + Xx1, m + ((n - 1) * Xx1) + 1, m + ((n - 1) * Xx1) + Xx1 + 1
                    Next m
                Next n
            '============================================================================================


            '============================================================================================
             'Loads a raw ASC text file. Check out the examples
             Case "asc"
                Do: Input #1, Temp: Loop While Temp <> "Tri-mesh"
                Input #1, LineON
                Num = 1
                For n = 1 To Len(LineON)
                    If Mid(LineON, n, 1) = ":" Then
                        If Num = 1 Then VertexCount = Mid(LineON, n + 1, 4)
                        If Num = 2 Then FaceCount = Mid(LineON, n + 1, 4)
                        Num = Num + 1
                    End If
                Next n
                Input #1, LineON
                For m = 1 To VertexCount
                    Input #1, LineON
                    Num = 1
                    For n = 1 To Len(LineON)
                        If Mid(LineON, n, 1) = ":" Then Poss(Num) = n: Num = Num + 1
                    Next n
                    Poss(5) = Len(LineON)
                    x = Val(Mid(LineON, Poss(2) + 1, Poss(3) - Poss(2) - 2))
                    y = Val(Mid(LineON, Poss(3) + 1, Poss(4) - Poss(3) - 2))
                    z = Val(Mid(LineON, Poss(4) + 1, Poss(5) - Poss(4) - 2))
                    .Vertex.Add x * 2, -z * 2, y * 2
                Next m
                Input #1, LineON
                FaceON = 1
                For m = 1 To FaceCount
                    Input #1, LineON
                    Num = 1
                    For n = 1 To Len(LineON)
                        If Mid(LineON, n, 1) = ":" Then Poss(Num) = n: Num = Num + 1
                    Next n
                    Xx1 = Val(Mid(LineON, Poss(2) + 1, Poss(3) - Poss(2) - 2)) + 1
                    Yy1 = Val(Mid(LineON, Poss(3) + 1, Poss(4) - Poss(3) - 2)) + 1
                    Zz1 = Val(Mid(LineON, Poss(4) + 1, Poss(5) - Poss(4) - 3)) + 1
                    .Face.Add 3, Zz1, Yy1, Xx1
                Next m
            '============================================================================================


            '============================================================================================
            'Return a message for unknown file formats
            Case Else
                MsgBox amUnKnownInport, vbExclamation
            '============================================================================================


        End Select
        .FindObjectOutline
        .Selected = True
        .Layer = ActiveFile.Layers.Default
    End With
    Close
    ImportModel = True
    ActiveFile.Geometery(NewKey).FindObjectOutline
    ActiveFile.FindModelOutline
Exit Function


FailedToImport:
    'This bit only gets run when an error occurs, which coloud be quit often on unknown files
    MsgBox amFailedInport, vbExclamation
End Function
