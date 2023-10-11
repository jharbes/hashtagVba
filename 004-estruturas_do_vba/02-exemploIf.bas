Attribute VB_Name = "Módulo1"
Sub calcula_bonus()

If Cells(3, 3).Value >= 100000 Then

Range("D3").Value = 0.13 * Range("C3").Value

Else

Range("D3").Value = 0

End If

End Sub

Sub calcula_bonus2()

If Cells(3, 3).Value >= 100000 Then

    Range("D3").Value = 0.13 * Range("C3").Value

ElseIf Cells(3, 3).Value >= 70000 Then

    Range("D3").Value = 0.07 * Range("C3").Value

Else

    Range("D3").Value = 0

End If

End Sub

Sub salario()

If Cells(3, 3).Value = "RJ" Then

    Cells(3, 4).Value = 7000

ElseIf Cells(3, 3).Value = "SP" Then

    Cells(3, 4).Value = 5500

ElseIf Cells(3, 3).Value = "RS" Then

    Cells(3, 4).Value = 5000

Else

    Cells(3, 4).Value = 4000

End If

End Sub

Sub calcula_bonus3()

If Cells(3, 3).Value >= 50000 And Cells(3, 4).Value >= 0.75 Then

    Cells(3, 5).Value = 0.15 * Cells(3, 3).Value

Else

    Cells(3, 5).Value = 0
    
End If

End Sub

Sub calcula_bonus4()

If Cells(3, 3).Value >= 80000 Or Cells(3, 4).Value >= 8 Then

    Cells(3, 5).Value = 0.15 * Cells(3, 3).Value

Else

    Cells(3, 5).Value = 0
    
End If

End Sub


















