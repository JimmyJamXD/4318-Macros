Sub poly()
    Dim p1_7 As Double
    Dim p1_6 As Double
    Dim p1_5 As Double
    Dim p1_4 As Double
    Dim p1_3 As Double
    Dim p1_2 As Double
    Dim p1_1 As Double
    Dim p1_0 As Double
    
    Dim p2_7 As Double
    Dim p2_6 As Double
    Dim p2_5 As Double
    Dim p2_4 As Double
    Dim p2_3 As Double
    Dim p2_2 As Double
    Dim p2_1 As Double
    Dim p2_0 As Double
    
    Dim t0 As Double
    Dim t1 As Double
    Dim t2 As Double
    Dim t3 As Double
    Dim t4 As Double
    Dim t5 As Double
    Dim t6 As Double
    Dim t7 As Double
    
    Dim i As Integer
    Dim at As Integer
    
    Dim rng As range
    Set rng = Selection
    
    Dim poly1 As Variant
    Dim poly2 As Variant
    Dim temp As Variant
    
    ReDim poly1(0 To 7)
    ReDim poly2(0 To 7)
    ReDim temp(0 To 7)
    
    p1_7 = rng.Cells(RowIndex:=1, ColumnIndex:=1)
    p1_6 = rng.Cells(RowIndex:=1, ColumnIndex:=2)
    p1_5 = rng.Cells(RowIndex:=1, ColumnIndex:=3)
    p1_4 = rng.Cells(RowIndex:=1, ColumnIndex:=4)
    p1_3 = rng.Cells(RowIndex:=1, ColumnIndex:=5)
    p1_2 = rng.Cells(RowIndex:=1, ColumnIndex:=6)
    p1_1 = rng.Cells(RowIndex:=1, ColumnIndex:=7)
    p1_0 = rng.Cells(RowIndex:=1, ColumnIndex:=8)
    poly1(0) = p1_7
    poly1(1) = p1_6
    poly1(2) = p1_5
    poly1(3) = p1_4
    poly1(4) = p1_3
    poly1(5) = p1_2
    poly1(6) = p1_1
    poly1(7) = p1_0
    
    'disp f(x)
    rng.Cells(RowIndex:=1, ColumnIndex:=1) = "f(x)"
    at = 2
    For i = 0 To 7
        If poly1(i) <> 0 Then
            If i < 6 Then
                rng.Cells(RowIndex:=1, ColumnIndex:=at) = CStr(poly1(i)) + "x^" + CStr(7 - i)
            ElseIf i = 6 Then
                rng.Cells(RowIndex:=1, ColumnIndex:=at) = CStr(poly1(i)) + "x"
            ElseIf i = 7 Then
                rng.Cells(RowIndex:=1, ColumnIndex:=at) = poly1(i)
                rng.Cells(RowIndex:=1, ColumnIndex:=at).VerticalAlignment = xlCenter
            End If
            at = at + 1
        End If
    Next i
    
    While at <= 9
        rng.Cells(RowIndex:=1, ColumnIndex:=at) = ""
        at = at + 1
    Wend
    
    p2_7 = rng.Cells(RowIndex:=2, ColumnIndex:=1)
    p2_6 = rng.Cells(RowIndex:=2, ColumnIndex:=2)
    p2_5 = rng.Cells(RowIndex:=2, ColumnIndex:=3)
    p2_4 = rng.Cells(RowIndex:=2, ColumnIndex:=4)
    p2_3 = rng.Cells(RowIndex:=2, ColumnIndex:=5)
    p2_2 = rng.Cells(RowIndex:=2, ColumnIndex:=6)
    p2_1 = rng.Cells(RowIndex:=2, ColumnIndex:=7)
    p2_0 = rng.Cells(RowIndex:=2, ColumnIndex:=8)
    
    poly2(0) = p2_7
    poly2(1) = p2_6
    poly2(2) = p2_5
    poly2(3) = p2_4
    poly2(4) = p2_3
    poly2(5) = p2_2
    poly2(6) = p2_1
    poly2(7) = p2_0
    
    'disp g(x)
    rng.Cells(RowIndex:=2, ColumnIndex:=1) = "g(x)"
    at = 2
    For i = 0 To 7
        If poly2(i) <> 0 Then
            If i < 6 Then
                rng.Cells(RowIndex:=2, ColumnIndex:=at) = CStr(poly2(i)) + "x^" + CStr(7 - i)
            ElseIf i = 6 Then
                rng.Cells(RowIndex:=2, ColumnIndex:=at) = CStr(poly2(i)) + "x"
            ElseIf i = 7 Then
                rng.Cells(RowIndex:=2, ColumnIndex:=at) = poly2(i)
                rng.Cells(RowIndex:=2, ColumnIndex:=at).VerticalAlignment = xlCenter
            End If
            at = at + 1
        End If
    Next i
    
    While at <= 9
        rng.Cells(RowIndex:=2, ColumnIndex:=at) = ""
        at = at + 1
    Wend
    
    
    'Math
    'Addition
    t0 = add(p1_0, p2_0)
    t1 = add(p1_1, p2_1)
    t2 = add(p1_2, p2_2)
    t3 = add(p1_3, p2_3)
    t4 = add(p1_4, p2_4)
    t5 = add(p1_5, p2_5)
    t6 = add(p1_6, p2_6)
    t7 = add(p1_7, p2_7)
    
    temp(0) = t7
    temp(1) = t6
    temp(2) = t5
    temp(3) = t4
    temp(4) = t3
    temp(5) = t2
    temp(6) = t1
    temp(7) = t0
    
    rng.Cells(RowIndex:=4, ColumnIndex:=1) = "Add"
    at = 2
    For i = 0 To 7
        If temp(i) <> 0 Then
            If i < 6 Then
                rng.Cells(RowIndex:=4, ColumnIndex:=at) = CStr(temp(i)) + "x^" + CStr(7 - i)
            ElseIf i = 6 Then
                rng.Cells(RowIndex:=4, ColumnIndex:=at) = CStr(temp(i)) + "x"
            ElseIf i = 7 Then
                rng.Cells(RowIndex:=4, ColumnIndex:=at) = temp(i)
                rng.Cells(RowIndex:=4, ColumnIndex:=at).VerticalAlignment = xlCenter
            End If
            at = at + 1
        End If
    Next i
    
    While at <= 9
        rng.Cells(RowIndex:=4, ColumnIndex:=at) = ""
        at = at + 1
    Wend
    
    'Subtraction
    t0 = subtract(p1_0, p2_0)
    t1 = subtract(p1_1, p2_1)
    t2 = subtract(p1_2, p2_2)
    t3 = subtract(p1_3, p2_3)
    t4 = subtract(p1_4, p2_4)
    t5 = subtract(p1_5, p2_5)
    t6 = subtract(p1_6, p2_6)
    t7 = subtract(p1_7, p2_7)
    
    temp(0) = t7
    temp(1) = t6
    temp(2) = t5
    temp(3) = t4
    temp(4) = t3
    temp(5) = t2
    temp(6) = t1
    temp(7) = t0
    
    rng.Cells(RowIndex:=5, ColumnIndex:=1) = "Sub"
    at = 2
    For i = 0 To 7
        If temp(i) <> 0 Then
            If i < 6 Then
                rng.Cells(RowIndex:=5, ColumnIndex:=at) = CStr(temp(i)) + "x^" + CStr(7 - i)
            ElseIf i = 6 Then
                rng.Cells(RowIndex:=5, ColumnIndex:=at) = CStr(temp(i)) + "x"
            ElseIf i = 7 Then
                rng.Cells(RowIndex:=5, ColumnIndex:=at) = temp(i)
                rng.Cells(RowIndex:=5, ColumnIndex:=at).VerticalAlignment = xlCenter
            End If
            at = at + 1
        End If
    Next i
    
    While at <= 9
        rng.Cells(RowIndex:=5, ColumnIndex:=at) = ""
        at = at + 1
    Wend
    
    'Multiplication
    rng.Cells(RowIndex:=6, ColumnIndex:=1) = "Mul"
    Dim polyMul As Variant
    ReDim polyMul(0 To 14)
    
    at = 2
    polyMul = mul(p1_7, p1_6, p1_5, p1_4, p1_3, p1_2, p1_1, p1_0, p2_7, p2_6, p2_5, p2_4, p2_3, p2_2, p2_1, p2_0)
    For i = 0 To 14
        If polyMul(i) <> 0 Then
            If i < 13 Then
                rng.Cells(RowIndex:=6, ColumnIndex:=at) = CStr(polyMul(i)) + "x^" + CStr(14 - i)
            ElseIf i = 13 Then
                rng.Cells(RowIndex:=6, ColumnIndex:=at) = CStr(polyMul(i)) + "x"
            ElseIf i = 14 Then
                rng.Cells(RowIndex:=6, ColumnIndex:=at) = polyMul(i)
                rng.Cells(RowIndex:=6, ColumnIndex:=at).VerticalAlignment = xlCenter
            End If
            at = at + 1
        End If
    Next i
    
    While at <= 15
        rng.Cells(RowIndex:=6, ColumnIndex:=at) = ""
        at = at + 1
    Wend
    
    'Division
    rng.Cells(RowIndex:=7, ColumnIndex:=1) = "Div"
    rng.Cells(RowIndex:=7, ColumnIndex:=2) = "Ans"
    Dim polyDiv As Variant
    ReDim polyDiv(0 To 15)
    
    at = 3
    polyDiv = div(p1_7, p1_6, p1_5, p1_4, p1_3, p1_2, p1_1, p1_0, p2_7, p2_6, p2_5, p2_4, p2_3, p2_2, p2_1, p2_0)
    For i = 0 To 7
        If polyDiv(i) <> 0 Then
            If i < 6 Then
                rng.Cells(RowIndex:=7, ColumnIndex:=at) = CStr(polyDiv(i)) + "x^" + CStr(7 - i)
            ElseIf i = 6 Then
                rng.Cells(RowIndex:=7, ColumnIndex:=at) = CStr(polyDiv(i)) + "x"
            ElseIf i = 7 Then
                rng.Cells(RowIndex:=7, ColumnIndex:=at) = polyDiv(i)
                rng.Cells(RowIndex:=7, ColumnIndex:=at).VerticalAlignment = xlCenter
            End If
            at = at + 1
        End If
    Next i
    
    While at <= 10
        rng.Cells(RowIndex:=7, ColumnIndex:=at) = ""
        at = at + 1
    Wend
    
    Dim remCount As Integer 'to check if there is a remainder, if so then print
    remCount = 0
    at = 3
    rng.Cells(RowIndex:=8, ColumnIndex:=2) = "Rem"
    For i = 0 To 7
        If polyDiv(i + 8) <> 0 Then
            If i < 6 Then
                rng.Cells(RowIndex:=8, ColumnIndex:=at) = CStr(polyDiv(i + 8)) + "x^" + CStr(7 - i)
            ElseIf i = 6 Then
                rng.Cells(RowIndex:=8, ColumnIndex:=at) = CStr(polyDiv(i + 8)) + "x"
            ElseIf i = 7 Then
                rng.Cells(RowIndex:=8, ColumnIndex:=at) = polyDiv(i + 8)
                rng.Cells(RowIndex:=8, ColumnIndex:=at).VerticalAlignment = xlCenter
            End If
            at = at + 1
            remCount = remCount + 1
        End If
    Next i
    
    If remCount > 0 Then
        rng.Cells(RowIndex:=8, ColumnIndex:=at) = "/"
        rng.Cells(RowIndex:=8, ColumnIndex:=at).HorizontalAlignment = xlCenter
        at = at + 1
    
        For i = 0 To 7
            If poly2(i) <> 0 Then
                If i < 6 Then
                    rng.Cells(RowIndex:=8, ColumnIndex:=at) = CStr(poly2(i)) + "x^" + CStr(7 - i)
                ElseIf i = 6 Then
                    rng.Cells(RowIndex:=8, ColumnIndex:=at) = CStr(poly2(i)) + "x"
                ElseIf i = 7 Then
                    rng.Cells(RowIndex:=8, ColumnIndex:=at) = poly2(i)
                    rng.Cells(RowIndex:=8, ColumnIndex:=at).VerticalAlignment = xlCenter
                End If
                at = at + 1
            End If
        Next i
    Else
        rng.Cells(RowIndex:=8, ColumnIndex:=at) = 0
        at = at + 1
    End If
    
    While at <= 10
        rng.Cells(RowIndex:=8, ColumnIndex:=at) = ""
        at = at + 1
    Wend
    
End Sub
