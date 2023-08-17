'Estrutura de Repetiçao Do While no VBA

Do While '<condicao>

	'Aqui vao os comandos que vao ser executados até que a
	'condicao seja considerada falsa (deixe de ser atendida)

Loop



Dim linha as Integer
linha = 1

'Do while funciona conforme o tradicional while, o loop vai rodar enquanto a
'condicao for considerada verdadeira (logica invertida em relacao ao Do Until)
Do While linha <= 10 

	'Aqui vao os comandos que vao ser executados até que a
	'condicao deixe de ser atendida (seja considerada falsa)
	Cells(linha,4).Value = Cells(linha,3).Value * 0.7
	linha = linha + 1

Loop