Attribute VB_Name = "Module1"
Sub registra_venda()


marca = InputBox("Digite o nome da marca:")

Sheets("Dados").Activate

'descobre a linha onde está a marca que o usuario digitou
linha_marca = Cells.Find(marca).Row
preco_marca = Cells(linha_marca, 2).Value


Sheets("Vendas Diárias").Activate


'utiliza o path do arquivo atual onde está a macro
'logo os dois arquivos devem estar no mesmo diretório
Workbooks.Open (ThisWorkbook.Path & "\09-exercicio_estoque-estoque-resolucao.xlsm")

'localiza em qual linha do arquivo estoque que está a marca selecionada
linha_estoque = Cells.Find(marca).Row
quantidade_estoque = Cells(linha_estoque, 2).Value


If quantidade_estoque > 0 Then

    Cells(linha_estoque, 2).Value = Cells(linha_estoque, 2) - 1

End If

ActiveWorkbook.Save
ActiveWorkbook.Close


linha_disponivel = Range("A1").End(xlDown).Row + 1

Cells(linha_disponivel, 1).Value = linha_disponivel - 1
Cells(linha_disponivel, 2).Value = Date
Cells(linha_disponivel, 3).Value = marca
Cells(linha_disponivel, 4).Value = preco_marca
Cells(linha_disponivel, 5).Value = quantidade_estoque

If quantidade_estoque > 0 Then

    Cells(linha_disponivel, 6).Value = "Disponível"
    
Else

    Cells(linha_disponivel, 6).Value = "Indisponível"

End If


End Sub
