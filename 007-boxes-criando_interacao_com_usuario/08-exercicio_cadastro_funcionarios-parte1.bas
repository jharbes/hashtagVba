Attribute VB_Name = "Module1"
Sub novo_funcionario()


Dim ultima_linha_branco As Integer

ultima_linha_branco = Range("A1").End(xlDown).Row + 1

resposta_execucao = MsgBox("Deseja incluir um funcion�rio?", vbYesNo + vbQuestion, "Confirma��o")


If resposta_execucao = 6 Then
    
    nome = InputBox("Digite o nome completo do funcion�rio:", "NOME COMPLETO")
    area = InputBox("Digite a �rea em que o funcion�rio ir� atuar:", "�REA")
    salario = InputBox("Preencha o sal�rio do funcion�rio:", "SAL�RIO")
    
    Cells(ultima_linha_branco, 1).Value = nome
    Cells(ultima_linha_branco, 2).Value = area
    Cells(ultima_linha_branco, 3).Value = Format(salario, "Currency")
    
    
    'Coloca em ordem alfab�tica
    
    
    resposta1 = MsgBox("Macro executada com sucesso!", vbInformation)

Else

    resposta1 = MsgBox("Opera��o Cancelada", vbInformation)


End If



End Sub
