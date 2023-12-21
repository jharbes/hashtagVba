Attribute VB_Name = "Módulo1"
Function salariocomimposto(qtd_normal As Double, qtd_extra As Double, preco_normal As Double, preco_extra As Double) As Double

salario = qtd_normal * preco_normal + qtd_extra * preco_extra

If salario <= 12000 Then

salariocomimposto = salario

ElseIf salario <= 18000 Then

salariocomimposto = salario * 1.1

Else

salariocomimposto = salario * 1.125

End If

End Function



Sub calcula_salario()

    Dim linha As Integer
    Dim ultima_linha As Integer
    
    ultima_linha = Range("B6").End(xlDown).Row
    
    For linha = 7 To ultima_linha
    
        Cells(linha, 5).Value = salariocomimposto(Cells(linha, 3).Value, Cells(linha, 4).Value, Range("H6").Value, Range("H7").Value)
    
    Next
    

End Sub
