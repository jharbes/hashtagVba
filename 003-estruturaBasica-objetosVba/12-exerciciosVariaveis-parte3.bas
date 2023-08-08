Attribute VB_Name = "Module1"
Sub calculo_faturamento()

'Ativando a aba variáveis para que seja utilizada ela para rodar o código
Sheets("Resultado Operacional").Activate

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

'As linhas de baixo formatam as celulas C9 e C10 como dinheiro e percentual
'respectivamente, para descobrir sua formatacao gravamos a macro desses passos
'Range("C9").Select
'Selection.Style = "Currency"
'Range("C10").Select
'Selection.Style = "Percent"

'ou

Range("C9").Style = "Currency"
Range("C10").Style = "Percent"

End Sub


Sub cadastro_produto()


'Ativando a aba variáveis para que seja utilizada ela para rodar o código
Sheets("Variáveis").Activate

Cells(12, 2).Value = "Pastel de Queijo"



End Sub
