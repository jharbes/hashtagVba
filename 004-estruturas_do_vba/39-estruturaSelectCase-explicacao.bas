'Estrutura select Case no VBA


Select Case Cells(2,1).Value

	Case >= 7
			
		Cells(2,2).Value = “Aprovado”

	Case >= 5

		Cells(2,2).Value = “Prova Final”

	Case Else

		Cells(2,2).Value= “Reprovado”

End Select
