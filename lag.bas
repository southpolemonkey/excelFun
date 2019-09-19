' A lag like function to fill the blank cell from the closest value
' Last update late: 18/09/2019

Sub Lag()
    Dim target As Range
    Dim result As Range
    Set target = Application.InputBox("Select your target data", Type:=8)
    Set result = Application.InputBox("Select your result data", Type:=8)
    Dim LastR As Integer
    Dim LastC As Integer

    LastR = target.End(xlUp).row
    LastC = target.End(xlToLeft).Column

    Dim current_value As String
    current_value = Cstr(target.Cells(1, 1))

    For i = target.Rows(1).Row to target.Rows(1).Row + target.Rows.Count - 1
        If isEmpty(target.Cells(i, 1).Value) = True Then
            result.Cells(i ,1).Value = current_value            
        Else
            current_value = target.Cells(i, 1).Value
            result.Cells(i, 1).Value = current_value
        End If     
    Next
End Sub