'Application = Excel em si

'Workbooks = Cada arquivo do excel

'Sheets = Cada aba do excel

'Range/Cells = Cada célula ou conjunto de células de uma aba



'--------------------------------------------------------------------------------------------

'manipular elementos da planilha sem precisar alterar a planilha ativa
Worksheets("NomeDaPlanilha").Cells(linha, coluna).NumberFormat = "General"
Worksheets("NomeDaPlanilha").Cells(linha, coluna).Value = SeuNumero

'ou

Sheets("Planilha2").Cells(1, 2).Value = "Alon"



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

'tratamento de erros "resume next"

'indica que caso haja alguem erro basta seguir em frente rodando
'o restante da macro (coloca no inicio da macro)
On Error Resume Next



'tratamento de erros "goto label"

sub servicos()

on Error GoTo Tratar

Exit Sub

Tratar:
Msgbox("Funcionário sem serviço especificado, favor tratar!")

End Sub


'--------------------------------------------------------------------------------------------

'encerra todas as instancias do excel
Application.Quit 

'Retira as instancias de excel do full screen (tela cheia, sem os botoes e menus, nao é retirar de maximizar)
Application.DisplayFullScreen = False


'--------------------------------------------------------------------------------------------

'salvando a path do arquivo atual
caminho_do_arquivo = ThisWorkbook.Path
	
'Abre o arquivo em questão
Workbooks.Open(path)

Set wb = Workbooks.Open(path) 'seta a variavel wb como sendo o workbook aberto
'Salva o arquivo em questao
'salva o arquivo das areas
wb.Save
    
'CUIDADO!!
ThisWorkbook.Save 'arquivo onde está escrita a macro
ActiveWorkbook.Save 'arquivo sendo utilizado no momento

'Fechar o arquivo
wb.Close

ThisWorkbook.Close 'arquivo onde está escrita a macro
ActiveWorkbook.Close 'arquivo sendo utilizado no momento