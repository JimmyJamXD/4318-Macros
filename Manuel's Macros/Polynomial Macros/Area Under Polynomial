Sub Area()
    Dim p7 As Double
    Dim p6 As Double
    Dim p5 As Double
    Dim p4 As Double
    Dim p3 As Double
    Dim p2 As Double
    Dim p1 As Double
    Dim p0 As Double
    Dim at As Integer
    Dim i As Integer
    
    Dim left As Double
    Dim right As Double
    
    Dim rng As range
    Set rng = Selection
    
    Dim poly1 As Variant
    ReDim poly1(0 To 7)
    
    p7 = rng.Cells(RowIndex:=1, ColumnIndex:=1)
    p6 = rng.Cells(RowIndex:=1, ColumnIndex:=2)
    p5 = rng.Cells(RowIndex:=1, ColumnIndex:=3)
    p4 = rng.Cells(RowIndex:=1, ColumnIndex:=4)
    p3 = rng.Cells(RowIndex:=1, ColumnIndex:=5)
    p2 = rng.Cells(RowIndex:=1, ColumnIndex:=6)
    p1 = rng.Cells(RowIndex:=1, ColumnIndex:=7)
    p0 = rng.Cells(RowIndex:=1, ColumnIndex:=8)
    left = rng.Cells(RowIndex:=1, ColumnIndex:=9)
    right = rng.Cells(RowIndex:=1, ColumnIndex:=10)
    poly1(0) = p7
    poly1(1) = p6
    poly1(2) = p5
    poly1(3) = p4
    poly1(4) = p3
    poly1(5) = p2
    poly1(6) = p1
    poly1(7) = p0
    
    at = 1
    For i = 0 To 7
        If poly1(i) <> 0 Then
            If i < 6 Then
                rng.Cells(RowIndex:=1, ColumnIndex:=at) = CStr(poly1(i)) + "x^" + CStr(7 - i)
            ElseIf i = 6 Then
                rng.Cells(RowIndex:=1, ColumnIndex:=at) = CStr(poly1(i)) + "x"
            ElseIf i = 7 Then
                rng.Cells(RowIndex:=1, ColumnIndex:=at) = poly1(i)
            End If
            at = at + 1
        End If
    Next i
    
    rng.Cells(RowIndex:=1, ColumnIndex:=at) = "Range"
    at = at + 1
    rng.Cells(RowIndex:=1, ColumnIndex:=at) = left
    at = at + 1
    rng.Cells(RowIndex:=1, ColumnIndex:=at) = right
    at = at + 1
    
    While at <= 10
        rng.Cells(RowIndex:=1, ColumnIndex:=at) = ""
        at = at + 1
    Wend
    
    rng.Cells(RowIndex:=2, ColumnIndex:=1) = "Area"
    rng.Cells(RowIndex:=2, ColumnIndex:=2) = areaUnder(left, right, p7, p6, p5, p4, p3, p2, p1, p0)
    
End Sub
