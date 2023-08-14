'Estrutura Do Until no VBA

Do Until '<condicao>

	'Aqui vao os comandos que vao ser executados até que a
	'condicao seja atendida

Loop





Dim linha as Integer
linha = 1

'Observe que o Do Until tem a logica invertida em relacao ao while, ou seja, quando a condicao se
'tornar verdadeira ele para de executar
Do Until linha > 10 

	'Aqui vao os comandos que vao ser executados até que a
	'condicao seja atendida
	Cells(linha,4).Value = Cells(linha,3).Value * 0.7
	linha = linha + 1

Loop