Attribute VB_Name = "Module1"
Sub moedas()

Dim linha As Integer
Dim linha_moeda As Integer
Dim conversao As Double
Dim moeda As String
Dim total_brl As Double


ultima_linha = Range("K2").End(xlDown).Row
ultima_linha_moeda = Range("C3").End(xlDown).Row


For linha = 3 To ultima_linha

    Range("L" & linha).ClearContents

Next

For linha = 3 To ultima_linha

    moeda = Range("K" & linha).Value
    
    For linha_moeda = 4 To ultima_linha_moeda
    
        If Range("B" & linha_moeda).Value = moeda Then
            conversao = Range("C" & linha_moeda).Value
        End If
    
    Next
    
    total_brl = conversao * Range("J" & linha).Value
    Range("L" & linha).Value = total_brl

Next

End Sub
