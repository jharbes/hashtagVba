'Estrutura If no VBA

'If com apenas Else

If Range("A2").Value >= 7 Then

	'Comandos executados caso a condicao seja verdadeira
	Range("B2").Value = "Aprovado"

Else

	'Comandos executados caso a condicao seja falsa
	Range("B2").Value = "Reprovado"

End If



'If com ElseIf e Else, podem ser usados quantos "ElseIf" forem necessarios

If Range("A2").Value >= 7 Then

	'Comandos executados caso a 1a condicao seja verdadeira
	Range("B2").Value = "Aprovado"
	
ElseIf Range("A2").Value >= 5 Then

	'Comandos executados caso a 2a condicao seja verdadeira
	Range("B2").Value = "Prova Final"

Else

	'Comandos executados caso a condicao seja falsa
	Range("B2").Value = "Reprovado"

End If



'If com o condicional And


Sub calcula_bonus3()

If Cells(3, 3).Value >= 50000 And Cells(3, 4).Value >= 0.75 Then

    Cells(3, 5).Value = 0.15 * Cells(3, 3).Value

Else

    Cells(3, 5).Value = 0
    
End If

End Sub



'If com o condicional Or


Sub calcula_bonus4()

If Cells(3, 3).Value >= 80000 Or Cells(3, 4).Value >= 8 Then

    Cells(3, 5).Value = 0.15 * Cells(3, 3).Value

Else

    Cells(3, 5).Value = 0
    
End If

End Sub