Public Function Add(num1 As Double, num2 As Double) As Double
    Add = num1 + num2
End Function

Public Function Subtract(num1 As Double, num2 As Double) As Double
    Subtract = num1 - num2
End Function

Public Function mulReal(a As Double, b As Double, c As Double, d As Double) As Double
    mulReal = (a * c) - (b * d)
End Function

Public Function mulImag(a As Double, b As Double, c As Double, d As Double) As Double
    mulImag = (a * d) + (b * c)
End Function

Public Function DivReal(a As Double, b As Double, c As Double, d As Double) As Double
    Dim c2 As Double
    Dim d2 As Double
    c2 = c * c
    d2 = d * d
    DivReal = ((a * c) + (b * d)) / (c2 + d2)
End Function

Public Function DivImag(a As Double, b As Double, c As Double, d As Double) As Double
    Dim c2 As Double
    Dim d2 As Double
    c2 = c * c
    d2 = d * d
    DivImag = ((-a * d) + (b * c)) / (c2 + d2)
End Function

Public Function magnitude(r As Double, i As Double) As Double
    Dim r2 As Double
    Dim i2 As Double
    Dim sum As Double
    r2 = r * r
    i2 = i * i
    sum = r2 + i2
    magnitude = Sqr(sum)
End Function

Public Function phase(r As Double, i As Double) As Double
    Dim div As Double
    div = i / r
    phase = Atn(div)
End Function

Public Function powerReal(r As Double, i As Double, j As Integer) As Double
    Dim tempReal As Double
    Dim tempImag As Double
    Dim oldReal As Double
    tempReal = r
    tempImag = i
    oldReal = r

    For k = 1 To j
        tempReal = mulReal(r, i, tempReal, tempImag)
        tempImag = mulImag(r, i, oldReal, tempImag)
        oldReal = tempReal
    Next k
    
    powerReal = tempReal
End Function

Public Function powerImag(r As Double, i As Double, j As Integer) As Double
    Dim tempReal As Double
    Dim tempImag As Double
    Dim oldReal As Double
    tempReal = r
    tempImag = i
    oldReal = r
    
    For k = 1 To j
        tempReal = mulReal(r, i, tempReal, tempImag)
        tempImag = mulImag(r, i, oldReal, tempImag)
        oldReal = tempReal
    Next k
    
    powerImag = tempImag
End Function

Public Function TF_Real(r1 As Double, i1 As Double, r2 As Double, i2 As Double, a As Double, theta As Double) As Double
    Dim tempR1 As Double
    Dim tempI1 As Double
    Dim tempR2 As Double
    Dim tempI2 As Double
    
    tempR1 = r1 - a
    tempI1 = i1
    tempR2 = r2 - (a * (Cos(theta)))
    tempI2 = i2 + a * (Sin(theta))
    
    TF_Real = DivReal(tempR1, tempI1, tempR2, tempI2)
End Function

Public Function TF_Imag(r1 As Double, i1 As Double, r2 As Double, i2 As Double, a As Double, theta As Double) As Double
    Dim tempR1 As Double
    Dim tempI1 As Double
    Dim tempR2 As Double
    Dim tempI2 As Double
    
    tempR1 = r1 - a
    tempI1 = i1
    tempR2 = r2 - (a * (Cos(theta)))
    tempI2 = i2 + a * (Sin(theta))
    
    TF_Imag = DivImag(tempR1, tempI1, tempR2, tempI2)
End Function ' - 11 lines
