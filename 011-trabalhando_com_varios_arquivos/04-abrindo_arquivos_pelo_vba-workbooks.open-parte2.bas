Attribute VB_Name = "Module1"
Sub registrar_funcionario()

'On Error Resume Next

Dim resposta As Integer
Dim resposta2 As Integer
Dim nome As String
Dim area As String
Dim salario As Double
Dim idade As Integer


resposta = MsgBox("Deseja rodar a macro?", vbYesNo + vbQuestion, "EXECUTAR MACRO")


If resposta = 6 Then
    
    nome = InputBox("Qual o nome do funcion�rio?", "NOME DO FUNCION�RIO")
    area = InputBox("Qual a �rea do funcion�rio?", "�REA DO FUNCION�RIO")
    salario = InputBox("Qual o sal�rio do funcion�rio?", "SAL�RIO DO FUNCION�RIO")
    idade = InputBox("Qual a idade do funcion�rio?", "IDADE DO FUNCION�RIO")
    
    linha_resumo = Range("A1").End(xlDown).Row + 1
    
    Cells(linha_resumo, 1).Value = nome
    Cells(linha_resumo, 2).Value = area
    Cells(linha_resumo, 3).Value = salario
    Cells(linha_resumo, 4).Value = idade
    
    'salvando a path do arquivo atual
    caminho_do_arquivo = ThisWorkbook.Path
    
    'o outro arquivo deve estar no mesmo diretorio do que o arquivo original
    Workbooks.Open (caminho_do_arquivo & "\02-exercicio_arquivos-explicacao-areas.xlsm")
    
    
Else

    resposta2 = MsgBox("Execu��o da Macro Cancelada!", vbInformation, "EXECU��O CANCELADA")

End If

End Sub
