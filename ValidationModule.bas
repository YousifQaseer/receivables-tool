
Attribute VB_Name = "ValidationModule"

Option Explicit

Public Function IsValidAmount(valueToCheck As Variant) As Boolean

    If IsNumeric(valueToCheck) Then
        If valueToCheck > 0 Then
            IsValidAmount = True
        Else
            IsValidAmount = False
        End If
    Else
        IsValidAmount = False
    End If

End Function
