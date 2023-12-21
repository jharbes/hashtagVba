Attribute VB_Name = "Module1"
Function salario_com_imposto(num_hora_normal As Integer, num_hora_extra As Integer) As Double

    Dim salario As Double
    
    salario = num_hora_normal * Range("H6").Value + num_hora_extra * Range("H7").Value
    
    If salario <= 12000 Then
        salario_com_imposto = salario
    ElseIf salario <= 18000 Then
        salario_com_imposto = salario * 1.1
    Else
        salario_com_imposto = salario * 1.125
    End If

End Function

