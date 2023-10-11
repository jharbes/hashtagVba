Attribute VB_Name = "Module1"
'Para entrar no VBA vamos na aba 'Developer' do Excel e clicamos em 'Visual Basic'
'ou então usamos a tecla de atalho ALT + F11

'Para inserir um módulo vamos no Menu Superior -> 'Insert' -> 'Module'

Sub escreve()

    Range("C8").Value = "João"
    

End Sub


'Para executar o código usamos a tecla de atalho 'F5' ou vamos no Menu Superior de ícones -> 'Rub Sub/User Form'



'Escrevendo em várias células
'Observe que ele só executa a macro onde está o cursor do mouse, ou seja, não executará ambas as
'macros 'escreve' e 'melhor_canal


Sub melhor_canal()

    Range("B13:H15").Value = "Hashtag"
    

End Sub
