Attribute VB_Name = "Module6"
Sub cadastra_venda()

Dim linha As Integer
Dim ultima_linha As Integer
Dim data1 As Date
Dim modelo As String
Dim opcionais As String

resposta_execucao = MsgBox("Deseja rodar a macro?", vbYesNo + vbQuestion, "Confirma��o")

ultima_linha = 4

If resposta_execucao = 6 Then
    
    
    data1 = CDate(InputBox("Entre com a data da venda:", "Data"))
    modelo = InputBox("Entre com o modelo do carro:", "Carro")
    preco = InputBox("Entre com o pre�o:", "Pre�o")
    opcionais = InputBox("Digite os opcionais:", "Opcionais")
    
    Cells(ultima_linha, 2).Value = data1
    Cells(ultima_linha, 3).Value = modelo
    Cells(ultima_linha, 4).Value = Format(preco, "Currency")
    Cells(ultima_linha, 5).Value = opcionais
    
    resposta1 = MsgBox("Macro executada com sucesso!", vbInformation)

Else

    resposta1 = MsgBox("Opera��o Cancelada", vbInformation)


End If





End Sub
