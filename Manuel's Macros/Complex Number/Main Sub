Sub complex()
    Dim r1 As Double 'r1 is the first real number
    Dim i1 As Double 'i1 is the first imaginary number
    Dim r2 As Double 'r2 is the second real number
    Dim i2 As Double 'i2 is the second imaginary number
    Dim a As Double 'a is used for the TF (Transfer Function) function
    Dim theta As Double 'theta is used for the TF function
    Dim pi As Double 'pi = 3.14159265358979323846
    Dim H_real As Double 'real part of a transfer function
    Dim H_imag As Double 'imaginary part of a transfer function
    Dim j As Integer 'j is used for the for the power function
    Dim rng As range 'rng is the range for all the input values on excel
    
    Set rng = Selection
    
    r1 = rng.Cells(RowIndex:=1, ColumnIndex:=1)
    i1 = rng.Cells(RowIndex:=1, ColumnIndex:=2)
    r2 = rng.Cells(RowIndex:=1, ColumnIndex:=3)
    i2 = rng.Cells(RowIndex:=1, ColumnIndex:=4)
    j = rng.Cells(RowIndex:=1, ColumnIndex:=5)
    a = 0.9
    theta = 0
    pi = 3.14159265358979
    
    rng.Cells(RowIndex:=2, ColumnIndex:=1) = r1
    rng.Cells(RowIndex:=2, ColumnIndex:=2) = i1
    rng.Cells(RowIndex:=2, ColumnIndex:=3) = r2
    rng.Cells(RowIndex:=2, ColumnIndex:=4) = i2
    rng.Cells(RowIndex:=2, ColumnIndex:=5) = j
    
    rng.Cells(RowIndex:=1, ColumnIndex:=1) = "Real 1"
    rng.Cells(RowIndex:=1, ColumnIndex:=2) = "Imag 1"
    rng.Cells(RowIndex:=1, ColumnIndex:=3) = "Real 1"
    rng.Cells(RowIndex:=1, ColumnIndex:=4) = "Imag 2"
    rng.Cells(RowIndex:=1, ColumnIndex:=5) = "Power"
    
    rng.Cells(RowIndex:=3, ColumnIndex:=2) = "Real"
    rng.Cells(RowIndex:=3, ColumnIndex:=3) = "Imag"
    
    rng.Cells(RowIndex:=4, ColumnIndex:=1) = "Addition"
    rng.Cells(RowIndex:=4, ColumnIndex:=2) = Add(r1, r2)
    rng.Cells(RowIndex:=4, ColumnIndex:=3) = CStr(Add(i1, i2)) + "i"
    
    rng.Cells(RowIndex:=5, ColumnIndex:=1) = "Subtraction"
    rng.Cells(RowIndex:=5, ColumnIndex:=2) = Subtract(r1, r2)
    rng.Cells(RowIndex:=5, ColumnIndex:=3) = CStr(Subtract(i1, i2)) + "i"
    
    rng.Cells(RowIndex:=6, ColumnIndex:=1) = "Multiplication"
    rng.Cells(RowIndex:=6, ColumnIndex:=2) = mulReal(r1, i1, r2, i2)
    rng.Cells(RowIndex:=6, ColumnIndex:=3) = CStr(mulImag(r1, i1, r2, i2)) + "i"
    
    rng.Cells(RowIndex:=7, ColumnIndex:=1) = "Division"
    rng.Cells(RowIndex:=7, ColumnIndex:=2) = DivReal(r1, i1, r2, i2)
    rng.Cells(RowIndex:=7, ColumnIndex:=3) = CStr(DivImag(r1, i1, r2, i2)) + "i"
    
    rng.Cells(RowIndex:=8, ColumnIndex:=2) = "Mag"
    rng.Cells(RowIndex:=8, ColumnIndex:=3) = "Ang"
    
    rng.Cells(RowIndex:=9, ColumnIndex:=1) = "Mag*e^(i*Ang)"
    rng.Cells(RowIndex:=9, ColumnIndex:=2) = magnitude(r1, i1)
    rng.Cells(RowIndex:=9, ColumnIndex:=3) = phase(r1, i1)
    rng.Cells(RowIndex:=9, ColumnIndex:=4) = "Complex 1"
    
    rng.Cells(RowIndex:=10, ColumnIndex:=1) = "Mag*e^(i*Ang)"
    rng.Cells(RowIndex:=10, ColumnIndex:=2) = magnitude(r2, i2)
    rng.Cells(RowIndex:=10, ColumnIndex:=3) = phase(r2, i2)
    rng.Cells(RowIndex:=10, ColumnIndex:=4) = "Complex 2"
    
    rng.Cells(RowIndex:=11, ColumnIndex:=2) = "Power"
    rng.Cells(RowIndex:=11, ColumnIndex:=3) = "Funct"
    
    rng.Cells(RowIndex:=12, ColumnIndex:=1) = "(r1 + i1j)^" + CStr(j)
    rng.Cells(RowIndex:=12, ColumnIndex:=2) = powerReal(r1, i1, j - 1)
    rng.Cells(RowIndex:=12, ColumnIndex:=3) = CStr(powerImag(r1, i1, j - 1)) + "i"
    
    rng.Cells(RowIndex:=14, ColumnIndex:=2) = "Trans"
    rng.Cells(RowIndex:=14, ColumnIndex:=3) = "Funct"
    
    H_real = TF_Real(r1, i1, r2, i2, a, theta)
    H_imag = TF_Imag(r1, i1, r2, i2, a, theta)
    rng.Cells(RowIndex:=15, ColumnIndex:=1) = "H(0)"
    rng.Cells(RowIndex:=15, ColumnIndex:=2) = H_real
    rng.Cells(RowIndex:=15, ColumnIndex:=3) = CStr(H_imag) + "i"
    rng.Cells(RowIndex:=16, ColumnIndex:=1) = "Mag*e^(i*Ang)"
    rng.Cells(RowIndex:=16, ColumnIndex:=2) = magnitude(H_real, H_imag)
    rng.Cells(RowIndex:=16, ColumnIndex:=3) = phase(H_real, H_imag)
    
    theta = pi / 6
    H_real = TF_Real(r1, i1, r2, i2, a, theta)
    H_imag = TF_Imag(r1, i1, r2, i2, a, theta)
    rng.Cells(RowIndex:=18, ColumnIndex:=1) = "H(pi/6)"
    rng.Cells(RowIndex:=18, ColumnIndex:=2) = H_real
    rng.Cells(RowIndex:=18, ColumnIndex:=3) = CStr(H_imag) + "i"
    rng.Cells(RowIndex:=19, ColumnIndex:=1) = "Mag*e^(i*Ang)"
    rng.Cells(RowIndex:=19, ColumnIndex:=2) = magnitude(H_real, H_imag)
    rng.Cells(RowIndex:=19, ColumnIndex:=3) = phase(H_real, H_imag)
    
    theta = pi / 4
    H_real = TF_Real(r1, i1, r2, i2, a, theta)
    H_imag = TF_Imag(r1, i1, r2, i2, a, theta)
    rng.Cells(RowIndex:=21, ColumnIndex:=1) = "H(pi/4)"
    rng.Cells(RowIndex:=21, ColumnIndex:=2) = H_real
    rng.Cells(RowIndex:=21, ColumnIndex:=3) = CStr(H_imag) + "i"
    rng.Cells(RowIndex:=22, ColumnIndex:=1) = "Mag*e^(i*Ang)"
    rng.Cells(RowIndex:=22, ColumnIndex:=2) = magnitude(H_real, H_imag)
    rng.Cells(RowIndex:=22, ColumnIndex:=3) = phase(H_real, H_imag)
    
    theta = pi / 2
    H_real = TF_Real(r1, i1, r2, i2, a, theta)
    H_imag = TF_Imag(r1, i1, r2, i2, a, theta)
    rng.Cells(RowIndex:=24, ColumnIndex:=1) = "H(pi/2)"
    rng.Cells(RowIndex:=24, ColumnIndex:=2) = H_real
    rng.Cells(RowIndex:=24, ColumnIndex:=3) = CStr(H_imag) + "i"
    rng.Cells(RowIndex:=25, ColumnIndex:=1) = "Mag*e^(i*Ang)"
    rng.Cells(RowIndex:=25, ColumnIndex:=2) = magnitude(H_real, H_imag)
    rng.Cells(RowIndex:=25, ColumnIndex:=3) = phase(H_real, H_imag)
    
    theta = 3 * pi / 4
    H_real = TF_Real(r1, i1, r2, i2, a, theta)
    H_imag = TF_Imag(r1, i1, r2, i2, a, theta)
    rng.Cells(RowIndex:=27, ColumnIndex:=1) = "H(3*pi/4)"
    rng.Cells(RowIndex:=27, ColumnIndex:=2) = H_real
    rng.Cells(RowIndex:=27, ColumnIndex:=3) = CStr(H_imag) + "i"
    rng.Cells(RowIndex:=28, ColumnIndex:=1) = "Mag*e^(i*Ang)"
    rng.Cells(RowIndex:=28, ColumnIndex:=2) = magnitude(H_real, H_imag)
    rng.Cells(RowIndex:=28, ColumnIndex:=3) = phase(H_real, H_imag)
    
    theta = 5 * pi / 6
    H_real = TF_Real(r1, i1, r2, i2, a, theta)
    H_imag = TF_Imag(r1, i1, r2, i2, a, theta)
    rng.Cells(RowIndex:=30, ColumnIndex:=1) = "H(5*pi/6)"
    rng.Cells(RowIndex:=30, ColumnIndex:=2) = H_real
    rng.Cells(RowIndex:=30, ColumnIndex:=3) = CStr(H_imag) + "i"
    rng.Cells(RowIndex:=31, ColumnIndex:=1) = "Mag*e^(i*Ang)"
    rng.Cells(RowIndex:=31, ColumnIndex:=2) = magnitude(H_real, H_imag)
    rng.Cells(RowIndex:=31, ColumnIndex:=3) = phase(H_real, H_imag)
    
    theta = pi
    H_real = TF_Real(r1, i1, r2, i2, a, theta)
    H_imag = TF_Imag(r1, i1, r2, i2, a, theta)
    rng.Cells(RowIndex:=33, ColumnIndex:=1) = "H(pi)"
    rng.Cells(RowIndex:=33, ColumnIndex:=2) = H_real
    rng.Cells(RowIndex:=33, ColumnIndex:=3) = CStr(H_imag) + "i"
    rng.Cells(RowIndex:=34, ColumnIndex:=1) = "Mag*e^(i*Ang)"
    rng.Cells(RowIndex:=34, ColumnIndex:=2) = magnitude(H_real, H_imag)
    rng.Cells(RowIndex:=34, ColumnIndex:=3) = phase(H_real, H_imag)
End Sub ' - 22 lines
