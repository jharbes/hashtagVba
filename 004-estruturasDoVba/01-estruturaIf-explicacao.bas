'Estrutura If no VBA

If Range("A2").Value >= 7 Then

	'Comandos executados caso a condicao seja verdadeira
	Range("B2").Value = "Aprovado"

Else

	'Comandos executados caso a condicao seja falsa
	Range("B2").Value = "Reprovado"

End If