Attribute VB_Name = "Module3"
Sub arrumar_textos()

'Quando quisermos utilizar funcoes nativas do Excel precisaremos
'escrever o WorksheetFunction antes
'Exemplo: WorksheetFunction.Trim()


Dim linha As Integer
Dim ultima_linha As Integer
Dim nome As String
Dim sobrenome As String
Dim nome_completo As String

ultima_linha = Range("B2").End(xlDown).Row


For linha = 3 To ultima_linha

nome = WorksheetFunction.Trim(Cells(linha, 2).Value)
sobrenome = WorksheetFunction.Trim(Cells(linha, 3).Value)
nome_completo = nome & " " & sobrenome
Cells(linha, 4).Value = nome_completo

Next

'Data atual
Range("F3").Value = Date

'Hora atual
Range("H3").Value = Time

'Data e Hora atual
Range("F6").Value = Now()



End Sub
