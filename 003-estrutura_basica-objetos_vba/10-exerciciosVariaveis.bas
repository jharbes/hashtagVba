Attribute VB_Name = "Module1"
Sub calculo_faturamento()

Dim faturamento As Double
Dim impostoSobreFaturamento As Double
Dim custoSobreProdutoVendido As Double
Dim despesasOperacionais As Double
Dim despesasFinanceiras As Double
Dim outrasDespesas As Double

faturamento = Cells(2, 3).Value
impostoSobreFaturamento = Cells(3, 3)
custoSobreProdutoVendido = Cells(4, 3)
despesasOperacionais = Cells(5, 3)
outrasDespesas = Cells(6, 3)

lucro = faturamento - impostoSobreFaturamento - custoSobreProdutoVendido - despesasFinanceiras - outrasDespesas
margem = lucro / faturamento

Cells(9, 3) = lucro
Cells(10, 3) = margem



End Sub
