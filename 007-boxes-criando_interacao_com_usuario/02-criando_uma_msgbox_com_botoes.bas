Attribute VB_Name = "Module1"
Sub mensagem()

MsgBox ("Seja bem vindo!")


End Sub


Sub mensagem_nome()

MsgBox ("Bom dia " & Application.UserName & "!")


End Sub


Sub deseja_executar()

'msgbox com a opção de sim ou não
'a reposta (sim ou nao) será armazenada na variavel reposta
'para possibilidade de usar posteriormente
resposta = MsgBox("Deseja executar a macro?", vbYesNo)

'observe que a variavel resposta recebe o valor 6 caso a
'a resposta seja sim e 7 caso a resposta seja nao, conforme
'tabela padrao de saidas das msgbox salva em imagem

If resposta = 6 Then

    MsgBox ("Macro executada com sucesso!")

Else

    MsgBox ("Macro cancelada!")

End If



End Sub

