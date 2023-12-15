Attribute VB_Name = "Module1"
Sub compila()

resposta = MsgBox("deseja realmente executar a macro?", vbYesNo + vbQuestion, "CONFIRMAÇÃO")

If reposta = 6 Then

    tipo_de_carro = InputBox("Deseja compilar os carros Novos ou Usados?", "TIPO DE CARRO", "Novo/Usado")
    

End If


End Sub
