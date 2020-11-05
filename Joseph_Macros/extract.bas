'extract macro can be used to separate measurements
'separate numberical values from units
'function used to separate numbers and string
'Strip(string,TRUE) = strips numbers
'Strip(string,FALSE) = strips string
Option Explicit
Public Function Strip(ByVal x As String, LeaveNums As Boolean) As Variant
Dim y As String
Dim s1 As String
Dim i As Long
    'i=1 to length of x using len(x)
    For i = 1 To Len(x)
        'extract substring y from string x
        y = Mid(x, i, 1)
        If LeaveNums = False Then
            If y Like "[A-Za-z ]" Then s1 = s1 & y 'False keeps Letters and spaces only
        Else
            If y Like "[0-9. ]" Then s1 = s1 & y   'True keeps Numbers and decimal points
        End If
    Next i
Strip = Trim(s1)
End Function

'extract macros sets up titles in above column
'may want to add row prior to running macro
Sub extract()
    Dim myRange As Range
    Dim myCell As Range
    Set myRange = Selection
    'set up table
    ActiveCell.Offset(-1, 0).Activate
    'make sure we dont overwrite other entries
    If IsEmpty(ActiveCell) Then
    ActiveCell.Value = "Original Entry"
    End If
    ActiveCell.Offset(0, 1).Activate
    If IsEmpty(ActiveCell) Then
    ActiveCell.Value = "Num Value"
    End If
    ActiveCell.Offset(0, 1).Activate
    If IsEmpty(ActiveCell) Then
    ActiveCell.Value = "Units"
    End If
    'return back to first cell
    ActiveCell.Offset(1, -2).Select
    For Each myCell In myRange
    'move right 1 cell
    ActiveCell.Offset(0, 1).Select
    'place number value
    ActiveCell.Value = "=Strip(RC[-1], TRUE)"
    'move right 1 cell again
    ActiveCell.Offset(0, 1).Select
    'place string value
    ActiveCell.Value = "=Strip(RC[-2], FALSE)"
    'return to original column and shift down 1 cell
    ActiveCell.Offset(1, -2).Select
    'move to next entry
    Next myCell
End Sub
