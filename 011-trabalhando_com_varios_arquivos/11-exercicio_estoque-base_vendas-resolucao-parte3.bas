Attribute VB_Name = "Module1"
Sub registra_venda()


marca = InputBox("Digite o nome da marca:")

Sheets("Dados").Activate

'descobre a linha onde est� a marca que o usuario digitou
linha_marca = Cells.Find(marca).Row
preco_marca = Cells(linha_marca, 2).Value


Sheets("Vendas Di�rias").Activate


'utiliza o path do arquivo atual onde est� a macro
'logo os dois arquivos devem estar no mesmo diret�rio
Workbooks.Open (ThisWorkbook.Path & "\09-exercicio_estoque-estoque-resolucao.xlsm")

'localiza em qual linha do arquivo estoque que est� a marca selecionada
linha_estoque = Cells.Find(marca).Row
quantidade_estoque = Cells(linha_estoque, 2).Value


If quantidade_estoque > 0 Then

    Cells(linha_estoque, 2).Value = Cells(linha_estoque, 2) - 1

End If

ActiveWorkbook.Save
ActiveWorkbook.Close


End Sub
