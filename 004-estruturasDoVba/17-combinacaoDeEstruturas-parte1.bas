Attribute VB_Name = "Módulo1"



Sub calcula_bonus()

Dim linha As Integer
linha = 10

Do Until Range("B" & linha).Value = "" Or Range("C" & linha).Value = ""

    If Range("C" & linha).Value <= 0.9 Then
        Range("D" & linha).Value = 0
    ElseIf Range("C" & linha).Value < 1 Then
        Range("D" & linha).Value = 500
    Else
        Range("D" & linha).Value = 3000
    End If
    
    linha = linha + 1
    
Loop

End Sub

