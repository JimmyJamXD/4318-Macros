Function funcAt(x As Double, p7 As Double, p6 As Double, p5 As Double, p4 As Double, p3 As Double, p2 As Double, p1 As Double, p0 As Double) As Double
    Dim t7 As Double
    Dim t6 As Double
    Dim t5 As Double
    Dim t4 As Double
    Dim t3 As Double
    Dim t2 As Double
    Dim t1 As Double
    Dim t0 As Double
    
    t7 = p7 * (x ^ 7)
    t6 = p6 * (x ^ 6)
    t5 = p5 * (x ^ 5)
    t4 = p4 * (x ^ 4)
    t3 = p3 * (x ^ 3)
    t2 = p2 * (x ^ 2)
    t1 = p1 * (x)
    t0 = p0
    
    funcAt = t7 + t6 + t5 + t4 + t3 + t2 + t1 + t0
End Function

Function areaUnder(left As Double, right As Double, p7 As Double, p6 As Double, p5 As Double, p4 As Double, p3 As Double, p2 As Double, p1 As Double, p0 As Double) As Double
    Dim count As Double
    Dim Area As Double
    Dim samples As Double
    Dim step As Double
    Dim i As Double
    count = 0
    Area = 0
    samples = (right - left) / 0.00001
    step = (right - left) / samples
    i = 0
    
    While i <= samples
        Area = Area + ((funcAt(left + i * step, p7, p6, p5, p4, p3, p2, p1, p0) + funcAt(left + (i + 1) * step, p7, p6, p5, p4, p3, p2, p1, p0)) * step) / 2
        i = i + 1
    Wend
    
    areaUnder = Area
End Function

Function largestDeg(p7 As Double, p6 As Double, p5 As Double, p4 As Double, p3 As Double, p2 As Double, p1 As Double, p0 As Double) As Integer
    Dim size As Double
    size = 0
    
    If p0 <> 0 Then size = 0
    If p1 <> 0 Then size = 1
    If p2 <> 0 Then size = 2
    If p3 <> 0 Then size = 3
    If p4 <> 0 Then size = 4
    If p5 <> 0 Then size = 5
    If p6 <> 0 Then size = 6
    If p7 <> 0 Then size = 7
    largestDeg = size
End Function

Function add(pA As Double, pB As Double) As Double
    add = pA + pB
End Function

Function subtract(pA As Double, pB As Double) As Double
    subtract = pA - pB
End Function

Function mul(p1_7 As Double, p1_6 As Double, p1_5 As Double, p1_4 As Double, p1_3 As Double, p1_2 As Double, p1_1 As Double, p1_0 As Double, p2_7 As Double, p2_6 As Double, p2_5 As Double, p2_4 As Double, p2_3 As Double, p2_2 As Double, p2_1 As Double, p2_0 As Double) As Variant
    Dim poly1 As Variant
    Dim poly2 As Variant
    Dim polyMul As Variant
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    ReDim poly1(0 To 7)
    ReDim poly2(0 To 7)
    ReDim polyMul(0 To 14)
    
    For k = 0 To 14
        polyMul(k) = 0
    Next k
    
    poly1(0) = p1_7
    poly1(1) = p1_6
    poly1(2) = p1_5
    poly1(3) = p1_4
    poly1(4) = p1_3
    poly1(5) = p1_2
    poly1(6) = p1_1
    poly1(7) = p1_0
    
    poly2(0) = p2_7
    poly2(1) = p2_6
    poly2(2) = p2_5
    poly2(3) = p2_4
    poly2(4) = p2_3
    poly2(5) = p2_2
    poly2(6) = p2_1
    poly2(7) = p2_0
    
    For i = 0 To 7
        If poly2(i) <> 0 Then
            For j = 0 To 7
                polyMul(i + j) = polyMul(i + j) + (poly2(i) * poly1(j))
            Next j
        End If
    Next i
    mul = polyMul
End Function

Function div(p1_7 As Double, p1_6 As Double, p1_5 As Double, p1_4 As Double, p1_3 As Double, p1_2 As Double, p1_1 As Double, p1_0 As Double, p2_7 As Double, p2_6 As Double, p2_5 As Double, p2_4 As Double, p2_3 As Double, p2_2 As Double, p2_1 As Double, p2_0 As Double) As Variant
    Dim pow1 As Integer
    Dim pow2 As Integer
    Dim poly1 As Variant
    Dim poly2 As Variant
    Dim polyDiv As Variant
    Dim i As Integer
    Dim t As Integer
    
    ReDim poly1(0 To 7)
    ReDim poly2(0 To 7)
    ReDim num(0 To 7)
    ReDim polyDiv(0 To 15)
    
    pow1 = largestDeg(p1_7, p1_6, p1_5, p1_4, p1_3, p1_2, p1_1, p1_0)
    pow2 = largestDeg(p2_7, p2_6, p2_5, p2_4, p2_3, p2_2, p2_1, p2_0)
    
    poly1(0) = p1_7
    poly1(1) = p1_6
    poly1(2) = p1_5
    poly1(3) = p1_4
    poly1(4) = p1_3
    poly1(5) = p1_2
    poly1(6) = p1_1
    poly1(7) = p1_0
    
    poly2(0) = p2_7
    poly2(1) = p2_6
    poly2(2) = p2_5
    poly2(3) = p2_4
    poly2(4) = p2_3
    poly2(5) = p2_2
    poly2(6) = p2_1
    poly2(7) = p2_0
    
    For i = 0 To 15
        polyDiv(i) = 0
    Next i
    
    'polyDiv(0) = pow1 'poly1(7 - pow1 - 1)
    'polyDiv(1) = pow2 'poly2(7 - pow2 - 1)
    'div = polyDiv
    
    Dim p1 As Variant
    Dim p2 As Variant
    
    Dim temp As Variant
    Dim j As Integer
    Dim k As Integer
    Dim tempVal As Integer
    
    ReDim p1(0 To pow1)
    ReDim p2(0 To pow2)
    
    t = pow1 - pow2
    k = 0

    If pow1 >= pow2 Then
        ReDim temp(0 To t)
        For i = 0 To pow1
            p1(i) = poly1(7 - pow1 + i)
        Next i
        For i = 0 To pow2
            p2(i) = poly2(7 - pow2 + i)
        Next i
        
        For i = 0 To t
            temp(i) = p1(i) / p2(0)
            p1(i) = 0
            If pow2 > 0 Then
                For j = 1 To pow2
                    p1(i + j) = p1(i + j) - (temp(i) * p2(j))
                Next j
            End If
            k = k + 1
        Next i
        
        For i = 0 To t
            polyDiv(7 - t + i) = temp(i)
        Next i
        
        For i = 0 To pow2 - 1
            polyDiv(15 - pow2 + 1 + i) = p1(k) 'remainder
            k = k + 1
        Next i
    Else
        For i = 8 To 15
            polyDiv(i) = poly1(i - 8)
        Next i
    End If
    div = polyDiv
End Function
