Attribute VB_Name = "Module1"
'Para entrar no VBA vamos na aba 'Developer' do Excel e clicamos em 'Visual Basic'
'ou ent�o usamos a tecla de atalho ALT + F11

'Para inserir um m�dulo vamos no Menu Superior -> 'Insert' -> 'Module'

Sub escreve()

    Range("C8").Value = "Jo�o"
    

End Sub


'Para executar o c�digo usamos a tecla de atalho 'F5' ou vamos no Menu Superior de �cones -> 'Rub Sub/User Form'



'Escrevendo em v�rias c�lulas
'Observe que ele s� executa a macro onde est� o cursor do mouse, ou seja, n�o executar� ambas as
'macros 'escreve' e 'melhor_canal


Sub melhor_canal()

    Range("B13:H15").Value = "Hashtag"
    

End Sub


'Escrevendo em uma c�lula com a fun��o Cells

Sub melhor_professor()


'Com a fun��o Cells passaremos o n�mero da linha e a coluna em formato num�rico:
'(linha,coluna) da c�lula
Cells(8, 3).Value = "Jo�o"

    

End Sub
