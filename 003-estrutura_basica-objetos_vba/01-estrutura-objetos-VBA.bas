' Criando um código em VBA

Sub nome()		'Nome da SUBrotina, nao usar caracteres especiais

	'Aqui vai seu código

End Sub			'Fim da SUBrotina


'Objetos

	'-Células (Cells/Range)
	'-Abas (Sheets)
	'-Arquivos (Workbooks)
	'-Gráficos (ChartObjects)
	
'Variações:

	'-ActiveCell (Célula Ativa)
	'-Selection (Objeto selecionado: pode ser uma célula, um intervalo de células, gráfico, etc)
	'-ActiveSheet (Aba Ativa)
	'-ActiveWorkbook (Arquivo Ativo)
	
	
'• Ações:

Range(“A1”).Select ' Seleciona a célula A1

Sheets(“Base de Dados”).Activate ' Ativa a Aba ‘Base de Dados’

Cells(2, 1).Copy ' Copia a célula A2 (linha 2, coluna 1)

Selecion.Copy ' Copia as células selecionadas

ActiveCell.ClearContents '-> Apaga o valor da célula ativa


'• Propriedades:

Range(“A1”).Value = 22 ' Escreve 22 na célula A1

Cells(3, 3).Interior.Color = vbYellow ' Pinta a célula C3 (linha 3, coluna 3) de amarelo

Sheets(“Base de Dados”).Name = “Cadastro Funcionários” ' Renomeia a aba ‘Base de Dados’ para ‘Cadastro Funcionários’

Selection.Value = “Hashtag” ' Escreve ‘Hashtag’ nas células selecionadas

ActiveCell.Value = “Alon” ' Escreve ‘Alon’ na célula ativa



Sub minha_macro()

	Range(“A1”).Select
	ActiveCell.Value = “Alon”
	'comentários
	Sheets(“Base de Dados”).Activate
	Cells(3, 3).Value = “Diego”
	Cells(3, 3).Interior.Color = vbYellow

End Sub