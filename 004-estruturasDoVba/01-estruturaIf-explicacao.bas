'Estrutura If no VBA

'If com apenas Else

If Range("A2").Value >= 7 Then

	'Comandos executados caso a condicao seja verdadeira
	Range("B2").Value = "Aprovado"

Else

	'Comandos executados caso a condicao seja falsa
	Range("B2").Value = "Reprovado"

End If



'If com ElseIf e Else

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