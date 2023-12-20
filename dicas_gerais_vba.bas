Worksheets("NomeDaPlanilha").Cells(linha, coluna).NumberFormat = "General"
Worksheets("NomeDaPlanilha").Cells(linha, coluna).Value = SeuNumero



'--------------------------------------------------------------------------------------------
'Como pular linha no codigo mas manter o entendimento do codigo inicial
'Ultilizando espaço + underline
'Exemplo:
resposta = MsgBox("deseja realmente executar a macro?", vbYesNo + vbQuestion, "CONFIRMAÇÃO")

resposta = MsgBox("deseja realmente executar a macro?", _
 vbYesNo + vbQuestion, "CONFIRMAÇÃO")



'--------------------------------------------------------------------------------------------

'utilizando goto em vba
'ao cair na linha do GoTo ele automaticamente sera redirecionado para a linha de codigo onde está
'o nome do GoTo (alon:)
alon:
	tipo = inputbox("Teste")
	
GoTo alon

'--------------------------------------------------------------------------------------------

'aumentar a velocidade da macro desligando a visualizacao da macro rodando

Application.ScreenUpdating = False

'Todo o codigo da macro aqui

Application.ScreenUpdating = True


'--------------------------------------------------------------------------------------------

'aumentar a velocidade da macro desligando o calculo automatico na tabela do excel
'calculo automatico é quando todas as celulas com formulas sao recalculadas quando alguma açao no excel é feita

Application.Calculation = xlCalculationManual

'Todo o codigo aqui

Application.Calculation = xlCalculationAutomatic


'podemos combinar essa e a anterior para agilizarmos a macro
'--------------------------------------------------------------------------------------------