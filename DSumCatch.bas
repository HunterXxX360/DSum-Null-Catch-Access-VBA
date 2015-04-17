Function DSumCatch(Expr As String, Domain As String, Optional Criteria) As Double
Dim Var As Variant
Dim Num As Double

Var = DSum(Expr, Domain, Criteria)

'   Check Null-values and numericality
If IsNull(Var) = True Then
    Num = 0
ElseIf IsNumeric(Num) And IsNull(Var) = False Then
    Num = Var
Else
    Num = 0
    Debug.Print "Var not numerical, therefore 0"
End If

DSumCatch = Num

End Function
