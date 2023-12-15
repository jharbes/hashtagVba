Worksheets("NomeDaPlanilha").Cells(linha, coluna).NumberFormat = "General"
Worksheets("NomeDaPlanilha").Cells(linha, coluna).Value = SeuNumero


'Como pular linha no codigo mas manter o entendimento do codigo inicial
'Ultilizando espaço + underline
'Exemplo:
resposta = MsgBox("deseja realmente executar a macro?", vbYesNo + vbQuestion, "CONFIRMAÇÃO")

resposta = MsgBox("deseja realmente executar a macro?", _
 vbYesNo + vbQuestion, "CONFIRMAÇÃO")
