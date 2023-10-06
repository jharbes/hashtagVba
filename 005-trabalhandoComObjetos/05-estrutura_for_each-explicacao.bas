'Estrutura For Each no VBA

For Each <variavel> in <conjunto>

	'Aqui vao os comandos que serao executados pra cada objeto do
	'conjunto (cada celula, cada aba, etc)
	
Next


'---------------------------------------------------
'Onde lemos <conjunto> temos os seguintes exemplos:

'Celulas: 	Range
'Abas: 		Sheets
'Arquivos:	Workbooks
'Graficos:	ChartObjects 


'Exemplo:

'celula é uma variavel do tipo Range, se fossemos declarar ficaria como:
Dim celula as Range

For Each celula in Range("A1:A10")
	
	'Vai escrever "VBA" nas celulas A1 até A10, primeiro na A1,
	'depois na A2, depois na A3, e assim por diante
	celula.Value = "VBA"
	
Next


'Exemplo 2:

'Se fossemos declarar a variavel chamada aba ficaria:
Dim aba as Worksheet

'Nesse caso o codigo vai percorrer CADA aba existente no arquivo escrevendo
'o valor "VBA" na celula A1
For Each aba in Sheets

	aba.Activate
	Cells(1,1).Value = "VBA"
	
Next