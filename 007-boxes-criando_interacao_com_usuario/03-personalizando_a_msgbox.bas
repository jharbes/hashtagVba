Attribute VB_Name = "Module2"
Sub deseja_executar2()

'MsgBox texto, bot�es, t�tulo, helpfile, contexto
'ou
'Vari�vel = (MsgBox texto, bot�es, t�tulo, helpfile, contexto)


'msgbox com a op��o de sim ou n�o
'a reposta (sim ou nao) ser� armazenada na variavel reposta
'para possibilidade de usar posteriormente

'para acrescentar um novo visual ao msgbox usaremos o caractere
'(+) logo apos o visual que ja existe:
resposta = MsgBox("Deseja executar a macro?", vbYesNo + vbQuestion, "Confirma��o")

'observe que a variavel resposta recebe o valor 6 caso a
'a resposta seja sim e 7 caso a resposta seja nao, conforme
'tabela padrao de saidas das msgbox salva em imagem

If resposta = 6 Then
    
    'observe que a msgbox aparece na tela mesmo registrando a variavel
    'junto do comando do msgbox
    resposta1 = MsgBox("Macro executada com sucesso!", vbInformation, "Executado")
    
Else

    resposta2 = MsgBox("Macro cancelada!", vbCritical, "Macro n�o executada!")

End If



End Sub
