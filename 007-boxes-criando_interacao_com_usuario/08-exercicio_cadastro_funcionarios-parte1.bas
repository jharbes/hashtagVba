Attribute VB_Name = "Module1"
Sub novo_funcionario()


Dim ultima_linha_branco As Integer

ultima_linha_branco = Range("A1").End(xlDown).Row + 1

resposta_execucao = MsgBox("Deseja incluir um funcionário?", vbYesNo + vbQuestion, "Confirmação")


If resposta_execucao = 6 Then
    
    nome = InputBox("Digite o nome completo do funcionário:", "NOME COMPLETO")
    area = InputBox("Digite a área em que o funcionário irá atuar:", "ÁREA")
    salario = InputBox("Preencha o salário do funcionário:", "SALÁRIO")
    
    Cells(ultima_linha_branco, 1).Value = nome
    Cells(ultima_linha_branco, 2).Value = area
    Cells(ultima_linha_branco, 3).Value = Format(salario, "Currency")
    
    
    'Coloca em ordem alfabética
    
    
    resposta1 = MsgBox("Macro executada com sucesso!", vbInformation)

Else

    resposta1 = MsgBox("Operação Cancelada", vbInformation)


End If



End Sub
