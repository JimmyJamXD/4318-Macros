Sub Calculate_Current()
'
' Calculate_Current Macro
'

'
    Application.CutCopyMode = False
    ActiveCell.Formula2R1C1 = "=RC[-7]:R[18]C[-7]/RC[-6]:R[18]C[-6]"
    Range("H3").Select
End Sub
