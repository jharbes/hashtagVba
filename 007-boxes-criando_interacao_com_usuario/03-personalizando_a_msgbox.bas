Attribute VB_Name = "Module2"
Sub deseja_executar2()

'MsgBox texto, botões, título, helpfile, contexto
'ou
'Variável = (MsgBox texto, botões, título, helpfile, contexto)


'msgbox com a opção de sim ou não
'a reposta (sim ou nao) será armazenada na variavel reposta
'para possibilidade de usar posteriormente

'para acrescentar um novo visual ao msgbox usaremos o caractere
'(+) logo apos o visual que ja existe:
resposta = MsgBox("Deseja executar a macro?", vbYesNo + vbQuestion, "Confirmação")

'observe que a variavel resposta recebe o valor 6 caso a
'a resposta seja sim e 7 caso a resposta seja nao, conforme
'tabela padrao de saidas das msgbox salva em imagem

If resposta = 6 Then
    
    'observe que a msgbox aparece na tela mesmo registrando a variavel
    'junto do comando do msgbox
    resposta1 = MsgBox("Macro executada com sucesso!", vbInformation, "Executado")
    
Else

    resposta2 = MsgBox("Macro cancelada!", vbCritical, "Macro não executada!")

End If



End Sub
