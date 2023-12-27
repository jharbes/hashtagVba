Attribute VB_Name = "Module1"
Sub registrar_funcionario()

'aumentar a velocidade da macro desligando a visualizacao da macro rodando
'importante desligar apos fim do codigo
Application.ScreenUpdating = False

'aumentar a velocidade da macro desligando o calculo automatico na tabela do excel
'importante desligar apos fim do codigo
Application.Calculation = xlCalculationManual


On Error Resume Next

Dim resposta As Integer
Dim resposta2 As Integer
Dim nome As String
Dim area As String
Dim salario As Double
Dim idade As Integer
Dim linha_resumo As Integer
Dim linha_resumo2 As Integer
Dim wb As Workbook


resposta = MsgBox("Deseja rodar a macro?", vbYesNo + vbQuestion, "EXECUTAR MACRO")


If resposta = 6 Then
    
    nome = InputBox("Qual o nome do funcionário?", "NOME DO FUNCIONÁRIO")
    
    area = InputBox("Qual a área do funcionário?", "ÁREA DO FUNCIONÁRIO")
    Do Until area = "Industrial" Or area = "Administrativo" Or area = "Logística" Or area = "Comercial"
    
        area = InputBox("Área inválida! Opções: Industrial/Administrativo/Logística/Comercial", "ÁREA DO FUNCIONÁRIO")
    
    Loop
    
    salario = InputBox("Qual o salário do funcionário?", "SALÁRIO DO FUNCIONÁRIO")
    Do Until Err.Number = 0
        
        Err.Clear ' Limpa o erro
        MsgBox "Por favor, insira apenas valores numéricos.", vbExclamation, "Erro de Entrada"
        salario = InputBox("Qual o salário do funcionário?", "SALÁRIO DO FUNCIONÁRIO")
        numero = CDbl(salario)
        
        If Err.Number = 0 Then
            Err.Clear ' Limpa o erro
        End If
    
    Loop
    
    idade = InputBox("Qual a idade do funcionário?", "IDADE DO FUNCIONÁRIO")
    Do Until Err.Number = 0
        
        Err.Clear ' Limpa o erro
        MsgBox "Por favor, insira apenas valores numéricos.", vbExclamation, "Erro de Entrada"
        idade = InputBox("Qual a idade do funcionário?", "IDADE DO FUNCIONÁRIO")
        numero = CDbl(idade)
        
        If Err.Number = 0 Then
            Err.Clear ' Limpa o erro
        End If
    
    Loop
    
    linha_resumo = Range("A1").End(xlDown).Row + 1
    
    Cells(linha_resumo, 1).Value = nome
    Cells(linha_resumo, 2).Value = area
    Cells(linha_resumo, 3).Value = salario
    Cells(linha_resumo, 4).Value = idade
    
    'salvando a path do arquivo atual
    caminho_do_arquivo = ThisWorkbook.Path
    
    'o outro arquivo deve estar no mesmo diretorio do que o arquivo original
    Set wb = Workbooks.Open(caminho_do_arquivo & "\02-exercicio_arquivos-explicacao-areas.xlsm")
    
    
    Sheets(area).Activate
    linha_resumo2 = Range("A1").End(xlDown).Row + 1
    
    Cells(linha_resumo2, 1).Value = nome
    Cells(linha_resumo2, 2).Value = area
    Cells(linha_resumo2, 3).Value = salario
    Cells(linha_resumo2, 4).Value = idade
    
    'salva o arquivo das areas
    wb.Save
    
    'fecha o arquivo das areas
    wb.Close
    
    resposta2 = MsgBox("Macro executada com sucesso!", vbInformation)
    
Else

    resposta2 = MsgBox("Execução da Macro Cancelada!", vbInformation, "EXECUÇÃO CANCELADA")

End If

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub
