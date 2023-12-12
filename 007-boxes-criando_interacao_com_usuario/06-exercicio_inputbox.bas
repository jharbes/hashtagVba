Attribute VB_Name = "Module5"
Sub cadastra_produto()

Dim linha As Integer
Dim ultima_linha As Integer

resposta_execucao = MsgBox("Deseja rodar a macro?", vbYesNo + vbQuestion, "Confirmação")


If resposta_execucao = 6 Then
        
    'Precisamos somar 1 à variavel pois ela para na ultima linha
    'preenchida e nao a primeira nao preenchida
    ultima_linha = Range("B6").End(xlDown).Row + 1
    
    
    produto = InputBox("Entre com o produto:", "Produto")
    preco_unitario = InputBox("Entre com o preco_unitario:", "Produto")
    estoque = InputBox("Entre com o quantidade em estoque:", "Produto")
    
    Cells(ultima_linha, 2).Value = produto
    Cells(ultima_linha, 3).Value = Format(preco_unitario, "Currency")
    Cells(ultima_linha, 4).Value = estoque
    
    resposta1 = MsgBox("Macro executada com sucesso!", vbInformation)

Else

    resposta1 = MsgBox("Operação Cancelada", vbInformation)


End If





End Sub
