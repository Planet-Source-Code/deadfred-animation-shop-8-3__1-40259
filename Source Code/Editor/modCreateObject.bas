Attribute VB_Name = "CreateObject"
Option Explicit

' ############################################################################
' #                                                                          #
' #   This module contains the code to create any of the default shapes.     #
' #   It takes a object class and builds the shape in that object using the  #
' #   parameters given. It can create the following objects                  #
' #                                                                          #
' #   1.  Sphere   [Vertical][Horizontal]                                    #
' #   2.  Tourus   [Vertical][Horizontal][Angle]                             #
' #   3.  Cube     None                                                      #
' #   4.  Prism    [Faces][Angle][Topface][BottomFace]                       #
' #   5.  Face     [Edges][Angle]                                            #
' #   6.  Grid     [Vertical][Horizontal]                                    #
' #   7.  Dimond   [Faces][Angle]                                            #
' #   8.  Cone     [Faces][Angle]                                            #
' #   9.  Star     [Faces][Angle][InnerRadius][OuterRadius]                  #
' #   10. Wrap     None (Uses a seperate data array)                         #
' #                                                                          #
' #   The first six parameters, o1 to o6 are used to set the position of     #
' #   new object. The rest of the parameters 07 onwards are used to define   #
' #   the actual object, and are used in different ways for different        #
' #   objects, as shown in the list above.                                   #
' #                                                                          #
' ############################################################################


Public Function Create3DObject(Object As clsObject, ViewMode As Integer, Class As String, Optional O1 As Integer = 0, Optional O2 As Integer = 0, Optional O3 As Integer = 0, Optional O4 As Integer = 0, Optional O5 As Integer = 0, Optional O6 As Integer = 0, Optional O7 As Integer = 0, Optional O8 As Integer = 0, Optional O9 As Integer = 0, Optional O10 As Integer = 0)

    Dim x1 As Integer, x2 As Integer, y1 As Integer, y2 As Integer, z1 As Integer, z2 As Integer
    Dim n As Single, TopSize As Single, BottomSize As Single, Srt As Integer, m As Single
    Dim sHeight As Integer, sWidth As Integer, sHorizontal As Single, sVertical As Single
    Dim Rotation As Integer, lHeight As Single, FaceON As Integer, XX As Single, Ct As Integer
    Dim HFace As Integer, VFace As Integer, X As Single, z As Single, MDown As Single
    Dim Ega As Single, Xx1 As Single, Yy1 As Single, Zz1 As Single, VAxis As Integer
    Dim Xx2 As Single, Yy2 As Single, Zz2 As Single, Mup As Integer
    
    x1 = O1:        x2 = O2
    y1 = O3:        y2 = O4
    z1 = O5:        z2 = O6
    
    If x1 > x2 Then Swap x1, x2
    If y1 > y2 Then Swap y1, y2
    If z1 > z2 Then Swap z1, z2

    With Object 'The string 'CLASS' holds the name of what type of object is to be created.
        Select Case Class


            '=========================================================================
            Case "Sphere"
                For n = 0 To 180 Step 180 / O7
                    If n = 0 Then
                        X = (Sin(m / Pie) * ((x2 - x1)) * Sin(n / Pie))
                        X = (X * 0.5) + ((x1 + x2) * 0.5)
                        z = Cos(m / Pie) * ((z2 - z1)) * Sin(n / Pie)
                        z = (z * 0.5)
                        Select Case ViewMode
                            Case 1: .Vertex.Add Int(X), Int(z), Int(-Cos(n / Pie) * ((z2 - z1) * 0.5) + ((z1 + z2) * 0.5))
                            Case 2: .Vertex.Add Int(X), Int(-Cos(n / Pie) * ((z2 - z1) * 0.5) + ((z1 + z2) * 0.5)), Int(z)
                            Case 3: .Vertex.Add Int(z), Int(-Cos(n / Pie) * ((z2 - z1) * 0.5) + ((z1 + z2) * 0.5)), Int(X)
                        End Select
                    ElseIf n = 180 Then
                        X = (Sin(m / Pie) * ((x2 - x1)) * Sin(n / Pie))
                        X = (X * 0.5) + ((x1 + x2) * 0.5)
                        z = (Cos(m / Pie) * ((z2 - z1)) * Sin(n / Pie)) * 0.5
                        z = (z * 0.5)
                        Select Case ViewMode
                            Case 1: .Vertex.Add Int(X), Int(z), Int(-Cos(n / Pie) * ((z2 - z1) * 0.5) + ((z1 + z2) * 0.5))
                            Case 2: .Vertex.Add Int(X), Int(-Cos(n / Pie) * ((z2 - z1) * 0.5) + ((z1 + z2) * 0.5)), Int(z)
                            Case 3: .Vertex.Add Int(z), Int(-Cos(n / Pie) * ((z2 - z1) * 0.5) + ((z1 + z2) * 0.5)), Int(X)
                        End Select
                    Else
                        For m = 0 To 359 Step 360 / O8
                            X = (Sin(m / Pie) * ((x2 - x1)) * Sin(n / Pie))
                            X = (X * 0.5) + ((x1 + x2) * 0.5)
                            z = Cos(m / Pie) * ((z2 - z1)) * Sin(n / Pie)
                            z = (z * 0.5)
                            Select Case ViewMode
                                Case 1: .Vertex.Add Int(X), Int(z), Int(-Cos(n / Pie) * ((z2 - z1) * 0.5) + ((z1 + z2) * 0.5))
                                Case 2: .Vertex.Add Int(X), Int(-Cos(n / Pie) * ((z2 - z1) * 0.5) + ((z1 + z2) * 0.5)), Int(z)
                                Case 3: .Vertex.Add Int(z), Int(-Cos(n / Pie) * ((z2 - z1) * 0.5) + ((z1 + z2) * 0.5)), Int(X)
                            End Select
                        Next m
                    End If
                Next n
                MDown = (O8 * 2) - O8 - O8 + 1
                For n = 1 To O8 - 1: .Face.Add 3, 1, MDown + n + 1, MDown + n: Next n
                .Face.Add 3, 1, MDown + n + 1 - O8, MDown + n: MDown = (O8 * (O7 - 2)) + 1
                For n = 1 To O8 - 1: .Face.Add 3, MDown + n, MDown + n + 1, .Vertex.Count: Next n
                .Face.Add 3, MDown + n, MDown + 1, .Vertex.Count
                For m = 2 To O7 - 1: MDown = (O8 * m) - O8 - O8 + 1
                    For n = 1 To O8 - 1: .Face.Add 4, MDown + n, MDown + n + 1, MDown + n + O8 + 1, MDown + n + O8
                    Next n: .Face.Add 4, MDown + n, MDown + n + 1 - O8, MDown + n + 1, MDown + n + O8
                Next m
            '=========================================================================


            '=========================================================================
            Case "Tourus"
                Rotation = (180 / O7) + (O10 * 5)
                For Ega = Rotation To 359 + Rotation Step 360 / O7
                    Xx1 = Sin(Ega / Pie): Xx1 = Xx1 * (x1 - x2) * 0.5
                    Xx1 = Xx1 + (x1 + x2) / 2: Zz1 = Cos(Ega / Pie)
                    Zz1 = Zz1 * (z1 - z2) * 0.5: Zz1 = Zz1 + (z1 + z2) / 2
                    Xx1 = Xx1 - VAxis
                    For n = 0 To 359 Step 360 / O9
                        Xx2 = (Cos(n / Pie) * Xx1)
                        Yy2 = (Sin(n / Pie) * Xx1)
                        Zz2 = Zz1
                        Select Case ViewMode
                            Case 1: .Vertex.Add Int(Xx2), Int(Yy2), Int(Zz2)
                            Case 2: .Vertex.Add Int(Xx2), Int(Zz2), Int(Yy2)
                            Case 3: .Vertex.Add Int(Yy2), Int(Zz2), Int(Xx2)
                        End Select
                    Next n
                Next Ega
                For m = 0 To O9 - 2
                    For n = 1 To O7 - 1
                        Mup = (n * O9) - O9
                        .Face.Add 4, 1 + Mup + m, 2 + Mup + m, (2 + O9 + Mup + m), 1 + O9 + Mup + m
                    Next n
                    Mup = (n * O9) - O9
                    .Face.Add 4, 1 + Mup + m, 2 + Mup + m, 2 + m, 1 + m
                Next m
                For n = 1 To O7 - 1
                    Mup = (n * O9) - O9
                    .Face.Add 4, 1 + Mup, O9 + 1 + Mup, (O9 * 2) + Mup, O9 + Mup
                Next n
                .Face.Add 4, 1, O9, .Vertex.Count, .Vertex.Count + 1 - O9
                .ReverseFace
            '=========================================================================


            '=========================================================================
            Case "Cube"
                .TexVert.Add 0, 20
                .TexVert.Add 0, 30
                .TexVert.Add 10, 0
                .TexVert.Add 10, 10
                .TexVert.Add 10, 20
                .TexVert.Add 10, 30
                .TexVert.Add 10, 40
                .TexVert.Add 20, 0
                .TexVert.Add 20, 10
                .TexVert.Add 20, 20
                .TexVert.Add 20, 30
                .TexVert.Add 20, 40
                .TexVert.Add 30, 20
                .TexVert.Add 30, 30
                For Ct = 1 To .TexVert.Count
                    .TexVert(Ct).X = .TexVert(Ct).X * 10
                    .TexVert(Ct).y = .TexVert(Ct).y * 10
                Next Ct
                Select Case ViewMode
                    Case 1
                        .Vertex.Add x1, y1, z2
                        .Vertex.Add x2, y1, z2
                        .Vertex.Add x2, y2, z2
                        .Vertex.Add x1, y2, z2
                        .Vertex.Add x1, y1, z1
                        .Vertex.Add x2, y1, z1
                        .Vertex.Add x2, y2, z1
                        .Vertex.Add x1, y2, z1
                    Case 2
                        .Vertex.Add x1, z1, y2
                        .Vertex.Add x2, z1, y2
                        .Vertex.Add x2, z2, y2
                        .Vertex.Add x1, z2, y2
                        .Vertex.Add x1, z1, y1
                        .Vertex.Add x2, z1, y1
                        .Vertex.Add x2, z2, y1
                        .Vertex.Add x1, z2, y1
                    Case 3
                        .Vertex.Add y1, z1, x2
                        .Vertex.Add y2, z1, x2
                        .Vertex.Add y2, z2, x2
                        .Vertex.Add y1, z2, x2
                        .Vertex.Add y1, z1, x1
                        .Vertex.Add y2, z1, x1
                        .Vertex.Add y2, z2, x1
                        .Vertex.Add y1, z2, x1
                End Select
                .Face.Add 4, 4, 3, 2, 1
                .Face.Add 4, 5, 6, 7, 8
                .Face.Add 4, 3, 4, 8, 7
                .Face.Add 4, 1, 2, 6, 5
                .Face.Add 4, 2, 3, 7, 6
                .Face.Add 4, 5, 8, 4, 1
                

                .Face(1).AddTextureMap 1, 2, 6, 5
                .Face(2).AddTextureMap 3, 4, 9, 8
                .Face(3).AddTextureMap 4, 5, 10, 9
                .Face(4).AddTextureMap 5, 6, 11, 10
                .Face(5).AddTextureMap 6, 7, 12, 11
                .Face(6).AddTextureMap 10, 11, 14, 13
            '=========================================================================


            '=========================================================================
            Case "Prism"
                TopSize = O9 / 20
                BottomSize = O8 / 20
                sVertical = (x1 + x2) / 2
                sHorizontal = (z1 + z2) / 2
                sHeight = Abs(x2 - x1) / 2
                sWidth = Abs(z2 - z1) / 2
                Rotation = (180 / O7) + (O10 * 5)
                .TexVert.Add 0, 0
                .TexVert.Add 10, 0
                .TexVert.Add 10, 10
                .TexVert.Add 0, 10
                For n = 1 To (O7 - 1) * 2 Step 2
                    .Face.Add 4, Int(n), Int(n + 2), Int(n + 3), Int(n + 1)
                    .Face(.Face.Count).AddTextureMap 1, 2, 3, 4
                Next n
                .Face.Add 4, Int(n), 1, 2, Int(n + 1)
                .Face(.Face.Count).AddTextureMap 1, 2, 3, 4
                Select Case ViewMode
                    Case 1
                        For n = 0 To 359 Step 360 / O7
                            .Vertex.Add -Sin((n + Rotation) / Pie) * (sHeight * TopSize) + sVertical, y2, -Cos((n + Rotation) / Pie) * (sWidth * TopSize) + sHorizontal
                            .Vertex.Add -Sin((n + Rotation) / Pie) * (sHeight * BottomSize) + sVertical, y1, -Cos((n + Rotation) / Pie) * (sWidth * BottomSize) + sHorizontal
                            .TexVert.Add 10 - Sin((n + Rotation) / Pie) * 10, 10 - Cos((n + Rotation) / Pie) * 10
                        Next n
                    Case 2
                        For n = 0 To 359 Step 360 / O7
                            .Vertex.Add -Sin((n + Rotation) / Pie) * (sHeight * TopSize) + sVertical, -Cos((n + Rotation) / Pie) * (sWidth * TopSize) + sHorizontal, z1
                            .Vertex.Add -Sin((n + Rotation) / Pie) * (sHeight * BottomSize) + sVertical, -Cos((n + Rotation) / Pie) * (sWidth * BottomSize) + sHorizontal, z2
                            .TexVert.Add 10 - Sin((n + Rotation) / Pie) * 10, 10 - Cos((n + Rotation) / Pie) * 10
                        Next n
                    Case 3
                        For n = 0 To 359 Step 360 / O7
                            .Vertex.Add z2, -Cos((n + Rotation) / Pie) * (sWidth * TopSize) + sHorizontal, -Sin((n + Rotation) / Pie) * (sHeight * TopSize) + sVertical
                            .Vertex.Add z1, -Cos((n + Rotation) / Pie) * (sWidth * BottomSize) + sHorizontal, -Sin((n + Rotation) / Pie) * (sHeight * BottomSize) + sVertical
                            .TexVert.Add 10 - Sin((n + Rotation) / Pie) * 10, 10 - Cos((n + Rotation) / Pie) * 10
                        Next n
                End Select
                .Colour = vbRed
                .Face.Add O7
                For n = 1 To O7
                    .Face(O7 + 1).Edge.Add n * 2
                    .Face(O7 + 1).Edge(.Face(O7 + 1).Edge.Count).TexVertex = n + 4
                Next n
                .Face.Add O7
                For n = O7 To 1 Step -1
                    .Face(O7 + 2).Edge.Add (n * 2) - 1
                    .Face(O7 + 2).Edge(.Face(O7 + 2).Edge.Count).TexVertex = n + 4
                Next n
            '=========================================================================


            '=========================================================================
            Case "Face"
                sVertical = (x1 + x2) / 2
                sHorizontal = (z1 + z2) / 2
                sHeight = Abs(x2 - x1) / 2
                sWidth = Abs(z2 - z1) / 2
                Rotation = (180 / O7) + (O10 * 5)
                Select Case ViewMode
                    Case 1
                        For n = 0 To 359 Step 360 / O7
                            .Vertex.Add -Sin((n + Rotation) / Pie) * (sHeight) + sVertical, y2, -Cos((n + Rotation) / Pie) * (sWidth) + sHorizontal
                        Next n
                    Case 2
                        For n = 0 To 359 Step 360 / O7
                            .Vertex.Add -Sin((n + Rotation) / Pie) * (sHeight) + sVertical, -Cos((n + Rotation) / Pie) * (sWidth) + sHorizontal, z1
                        Next n
                    Case 3
                        For n = 0 To 359 Step 360 / O7
                            .Vertex.Add z2, -Cos((n + Rotation) / Pie) * (sWidth) + sHorizontal, -Sin((n + Rotation) / Pie) * (sHeight) + sVertical
                        Next n
                End Select
                .Face.Add O7
                For n = 1 To O7: .Face(1).Edge.Add Int(n): Next n
            '=========================================================================


            '=========================================================================
            Case "Grid"
                For n = x1 To x2 Step (x2 - x1) / O7
                    For m = z1 To z2 Step (z2 - z1) / O8
                        If ViewMode = 1 Then .Vertex.Add CInt(n), y2, CInt(m)
                        If ViewMode = 2 Then .Vertex.Add CInt(n), CInt(m), y2
                        If ViewMode = 3 Then .Vertex.Add y2, CInt(m), CInt(n)
                    Next m
                Next n
                For n = 1 To O7
                    For m = 1 To O8
                        Srt = ((m - 1) * (O8 + 1)) + n
                       .Face.Add 4, Srt, Srt + 1, (Srt + 2 + O7), (Srt + 1 + O7)
                    Next m
                Next n
            '=========================================================================


            '=========================================================================
            Case "Dimond"
                sVertical = (x1 + x2) / 2
                sHorizontal = (z1 + z2) / 2
                sHeight = Abs(x2 - x1) / 2
                sWidth = Abs(z2 - z1) / 2
                Rotation = (180 / O7) + (O10 * 5)
                Select Case ViewMode
                    Case 1
                        For n = 0 To 359 Step 360 / O7
                            .Vertex.Add -Sin((n + Rotation) / Pie) * (sHeight) + sVertical, (y1 + y2) / 2, -Cos((n + Rotation) / Pie) * (sWidth) + sHorizontal
                        Next n
                        .Vertex.Add Int(sVertical), y1, Int(sHorizontal)
                        .Vertex.Add Int(sVertical), y2, Int(sHorizontal)
                    Case 2
                        For n = 0 To 359 Step 360 / O7
                            .Vertex.Add -Sin((n + Rotation) / Pie) * (sHeight) + sVertical, -Cos((n + Rotation) / Pie) * (sWidth) + sHorizontal, (z1 + z2) / 2
                        Next n
                        .Vertex.Add Int(sVertical), Int(sHorizontal), z2
                        .Vertex.Add Int(sVertical), Int(sHorizontal), z1
                    Case 3
                        For n = 0 To 359 Step 360 / O7
                            .Vertex.Add (z1 + z2) / 2, -Cos((n + Rotation) / Pie) * (sWidth) + sHorizontal, -Sin((n + Rotation) / Pie) * (sHeight) + sVertical
                        Next n
                        .Vertex.Add z1, Int(sHorizontal), Int(sVertical)
                        .Vertex.Add z2, Int(sHorizontal), Int(sVertical)
                End Select
                For n = 1 To O7 - 1: .Face.Add 3, Int(n), n + 1, O7 + 1: Next n
                For n = 1 To O7 - 1: .Face.Add 3, Int(n), O7 + 2, n + 1: Next n
                .Face.Add 3, Int(n), 1, O7 + 1
                .Face.Add 3, Int(n), O7 + 2, 1
            '=========================================================================


            '=========================================================================
            Case "Cone"
                sVertical = (x1 + x2) / 2
                sHorizontal = (z1 + z2) / 2
                sHeight = Abs(x2 - x1) / 2
                sWidth = Abs(z2 - z1) / 2
                Rotation = (180 / O7) + (O10 * 5)
                Select Case ViewMode
                    Case 1
                        For n = 0 To 359 Step 360 / O7
                            .Vertex.Add -Sin((n + Rotation) / Pie) * (sHeight) + sVertical, y2, -Cos((n + Rotation) / Pie) * (sWidth) + sHorizontal
                        Next n
                        .Vertex.Add Int(sVertical), y1, Int(sHorizontal)
                    Case 2
                        For n = 0 To 359 Step 360 / O7
                            .Vertex.Add -Sin((n + Rotation) / Pie) * (sHeight) + sVertical, -Cos((n + Rotation) / Pie) * (sWidth) + sHorizontal, z1
                        Next n
                        .Vertex.Add Int(sVertical), Int(sHorizontal), z2
                    Case 3
                        For n = 0 To 359 Step 360 / O7
                            .Vertex.Add z2, -Cos((n + Rotation) / Pie) * (sWidth) + sHorizontal, -Sin((n + Rotation) / Pie) * (sHeight) + sVertical
                        Next n
                        .Vertex.Add z1, Int(sHorizontal), Int(sVertical)
                End Select
                .Face.Add O7
                For n = 1 To O7: .Face(1).Edge.Add O7 - Int(n) + 1: Next n
                For n = 1 To O7 - 1: .Face.Add 3, Int(n), n + 1, O7 + 1: Next n
                .Face.Add 3, Int(n), 1, O7 + 1
            '=========================================================================


            '=========================================================================
            Case "Star"
                TopSize = O9 / 20
                sVertical = (x1 + x2) / 2
                sHorizontal = (z1 + z2) / 2
                sHeight = Abs(x2 - x1) / 2
                sWidth = Abs(z2 - z1) / 2
                O7 = O7 * 2
                Rotation = (180 / O7) + (O10 * 5)
                .TexVert.Add 0, 0
                .TexVert.Add 10, 0
                .TexVert.Add 10, 10
                .TexVert.Add 0, 10
                For n = 1 To (O7 - 1) * 2 Step 2
                    .Face.Add 4, Int(n), Int(n + 2), Int(n + 3), Int(n + 1)
                    .Face(.Face.Count).AddTextureMap 1, 2, 3, 4
                Next n
                .Face.Add 4, Int(n), 1, 2, Int(n + 1)
                .Face(.Face.Count).AddTextureMap 1, 2, 3, 4
                Select Case ViewMode
                    Case 1
                        For n = 0 To 359 Step 360 / O7
                            If TopSize = (O9 / 20) Then TopSize = (O8 / 20) Else TopSize = (O9 / 20)
                            .Vertex.Add -Sin((n + Rotation) / Pie) * (sHeight * TopSize) + sVertical, y2, -Cos((n + Rotation) / Pie) * (sWidth * TopSize) + sHorizontal
                            .Vertex.Add -Sin((n + Rotation) / Pie) * (sHeight * TopSize) + sVertical, y1, -Cos((n + Rotation) / Pie) * (sWidth * TopSize) + sHorizontal
                            .TexVert.Add 10 - Sin((n + Rotation) / Pie) * 10, 10 - Cos((n + Rotation) / Pie) * 10
                        Next n
                    Case 2
                        For n = 0 To 359 Step 360 / O7
                            If TopSize = (O9 / 20) Then TopSize = (O8 / 20) Else TopSize = (O9 / 20)
                            .Vertex.Add -Sin((n + Rotation) / Pie) * (sHeight * TopSize) + sVertical, -Cos((n + Rotation) / Pie) * (sWidth * TopSize) + sHorizontal, z1
                            .Vertex.Add -Sin((n + Rotation) / Pie) * (sHeight * TopSize) + sVertical, -Cos((n + Rotation) / Pie) * (sWidth * TopSize) + sHorizontal, z2
                            .TexVert.Add 10 - Sin((n + Rotation) / Pie) * 10, 10 - Cos((n + Rotation) / Pie) * 10
                        Next n
                    Case 3
                        For n = 0 To 359 Step 360 / O7
                            If TopSize = (O9 / 20) Then TopSize = (O8 / 20) Else TopSize = (O9 / 20)
                            .Vertex.Add z2, -Cos((n + Rotation) / Pie) * (sWidth * TopSize) + sHorizontal, -Sin((n + Rotation) / Pie) * (sHeight * TopSize) + sVertical
                            .Vertex.Add z1, -Cos((n + Rotation) / Pie) * (sWidth * TopSize) + sHorizontal, -Sin((n + Rotation) / Pie) * (sHeight * TopSize) + sVertical
                            .TexVert.Add 10 - Sin((n + Rotation) / Pie) * 10, 10 - Cos((n + Rotation) / Pie) * 10
                        Next n
                End Select
                .Colour = vbRed
                .Face.Add O7
                For n = 1 To O7
                    .Face(O7 + 1).Edge.Add n * 2
                    .Face(O7 + 1).Edge(.Face(O7 + 1).Edge.Count).TexVertex = n + 4
                Next n
                .FragmentFace O7 + 1, 2
                .Face.Add O7
                For n = O7 To 1 Step -1
                    .Face(.Face.Count).Edge.Add (n * 2) - 1
                    .Face(.Face.Count).Edge(.Face(.Face.Count).Edge.Count).TexVertex = n + 4
                Next n
                .FragmentFace .Face.Count, 2
            '=========================================================================
        
            
            '=========================================================================
            Case "Wrap"
                HFace = frmMain.ShpProp(1)
                Select Case ViewMode
                    Case 1
                        For FaceON = 1 To Am8(ActiveFile).Wraper.Count: For n = 0 To 359 Step 360 / HFace
                            Xx1 = Sin(n / Pie) * Am8(ActiveFile).Wraper(FaceON).X
                            Yy1 = Cos(n / Pie) * Am8(ActiveFile).Wraper(FaceON).X
                           .Vertex.Add Int(Xx1), Int(Yy1), Am8(ActiveFile).Wraper(FaceON).y
                        Next n: Next FaceON
                    
                    Case 2
                        For FaceON = 1 To Am8(ActiveFile).Wraper.Count: For n = 0 To 359 Step 360 / HFace
                            Xx1 = Sin(n / Pie) * Am8(ActiveFile).Wraper(FaceON).X
                            Yy1 = Cos(n / Pie) * Am8(ActiveFile).Wraper(FaceON).X
                           .Vertex.Add Int(Xx1), Am8(ActiveFile).Wraper(FaceON).y, Int(Yy1)
                        Next n: Next FaceON
                    
                    Case 3
                        For FaceON = 1 To Am8(ActiveFile).Wraper.Count: For n = 0 To 359 Step 360 / HFace
                            Xx1 = Sin(n / Pie) * Am8(ActiveFile).Wraper(FaceON).X
                            Yy1 = Cos(n / Pie) * Am8(ActiveFile).Wraper(FaceON).X
                           .Vertex.Add Int(Yy1), Am8(ActiveFile).Wraper(FaceON).y, Int(Xx1)
                        Next n: Next FaceON
                End Select
                For TopSize = 1 To Am8(ActiveFile).Wraper.Count - 1: BottomSize = (TopSize - 1) * HFace
                    For n = 1 To HFace - 1: .Face.Add 4, n + HFace + BottomSize, n + 1 + HFace + BottomSize, n + 1 + BottomSize, n + BottomSize
                    Next n: .Face.Add 4, n + HFace + BottomSize, n + 1 + BottomSize, n + 1 - HFace + BottomSize, n + BottomSize
                Next TopSize
            '=========================================================================
        
        
        End Select
    End With
End Function

