Attribute VB_Name = "Module1"
Sub Registro()

'Declaracao de variaveis não são obrigatorias, porem sao bastante indicadas
'pois farao com que a performance da macro seja melhorada
'e a organizacao do codigo seja melhorada

Dim linha As Integer
Dim produto As String

produto = "Guaraná Natural"

linha = 11

Cells(linha, 2).Value = produto
Cells(linha, 3).Value = 1.5
Cells(linha, 4).Value = 10000


End Sub
