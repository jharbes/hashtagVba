Attribute VB_Name = "Module1"
Sub registra_venda()


marca = InputBox("Digite o nome da marca:")

Sheets("Dados").Activate

'descobre a linha onde est� a marca que o usuario digitou
linha_marca = Cells.Find(marca).Row

preco_marca = Cells(linha_marca, 2).Value


End Sub
