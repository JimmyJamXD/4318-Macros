'caclCost Macro
'used to calculate total cost including tax
'uses range of selected prices
Sub calcCost()
    Dim tax As Double
    Dim sum As Double
    Dim cost As Double
    'initialize sum as zero
    sum = 0
    Dim myRange As Range
    Dim myCell As Range
    Set myRange = Selection
    For Each myCell In myRange
        'total up values in selection
        sum = sum + myCell.Value
        ActiveCell.Offset(1, 0).Select
    Next myCell
    'show cost without tax calc
    ActiveCell.Value = "Cost w/o Tax"
    ActiveCell.Offset(1, 0).Activate
    sum = FormatCurrency(sum)
    ActiveCell.Value = sum
    'enter tax percentage
    tax = InputBox("Enter Value", "Enter Tax Percent")
    'convert tax percent to decimal
    tax = tax / 100
    'calculate cost with tax
    cost = sum + sum * tax
    cost = FormatCurrency(cost)
    ActiveCell.Offset(1, 0).Activate
    'show cost with tax
    ActiveCell.Value = "Cost w/ Tax"
    ActiveCell.Offset(1, 0).Activate
    ActiveCell.Value = cost
End Sub

'enterValues macro used to input prices
'asks users for unit cost and quanity
'shows unit cost, quantity and total cost
'enterValues used to set up calcCost macro
Sub enterValues()
    Dim unitCost As Double
    Dim quantity As Integer
    Dim cost As Double
    'set unitCost to 1
    unitCost = 1
    'set up table
    ActiveCell.Value = "Unit Cost"
    ActiveCell.Offset(0, 1).Activate
    ActiveCell.Value = "Quantity"
    ActiveCell.Offset(0, 1).Activate
    ActiveCell.Value = "Total Cost"
    'return to original column
    ActiveCell.Offset(1, -2).Select
    'entering 0 for unitCost
    'will stop the macro
    Do While unitCost <> 0
        unitCost = InputBox("Enter Cost of Unit", "Enter '0' to exit")
        'only prompts if unitCost not 0
        If unitCost <> 0 Then
            quantity = InputBox("Enter Quantity", "Enter '0' to go back")
            cost = unitCost * quantity
            'nothing done if unitcost or quantity
            'are equal to 0
            If cost <> 0 Then
                ActiveCell.Value = FormatCurrency(unitCost)
                ActiveCell.Offset(0, 1).Activate
                ActiveCell.Value = quantity
                ActiveCell.Offset(0, 1).Activate
                ActiveCell.Value = FormatCurrency(cost)
                ActiveCell.Offset(1, -2).Activate
            End If
        End If
    Loop
End Sub
