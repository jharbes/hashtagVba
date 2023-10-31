Attribute VB_Name = "Module1"
Sub imprime_folhas_rosto()

Dim linha_inicio As Integer
Dim linha_fim As Integer
Dim linha_atual As Integer
Dim planejador As String

Dim pn As String
Dim descricao As String
Dim wbs As String
Dim projeto_com_remessa As String
Dim projeto As String
Dim remessa As String
Dim ordem As String
Dim tr As String
Dim data_necessidade As String
Dim data_atual As String

'Cria a data atual no formato dd/mm/yyyy para ser utilizada como
'data de abertura de ordem
data_atual = Format(Date, "dd/mm/yyyy")


Sheets("planilha_ordens").Activate

'Salva o nome do planejador
planejador = Range("I13").Value

'Salva o caminho do diretorio para salvar as folhas de rosto
diretorio_salvar = Range("I16").Value


linha_inicio = 2

'Encontra qual é a ultima linha preenchida da tabela "planilha_ordens"
linha_fim = Range("A1").End(xlDown).Row

'Percorre as linhas preenchidas da tabela "planilha_ordens"
For linha_atual = linha_inicio To linha_fim
    
    'Salva todos os dados de cada linha por vez da tabela "planilha_ordens"
    pn = Cells(linha_atual, 1).Value
    descricao = Cells(linha_atual, 2).Value
    wbs = Cells(linha_atual, 3).Value
    projeto_com_remessa = Cells(linha_atual, 4).Value
    posicao_hashtag = InStr(projeto_com_remessa, "#")
    projeto = Trim(Left(projeto_com_remessa, posicao_hashtag - 1))
    remessa = Mid(projeto_com_remessa, posicao_hashtag)
    ordem = Cells(linha_atual, 5).Value
    tr = Cells(linha_atual, 6).Value
    data_necessidade = Cells(linha_atual, 7).Value
    
    'Altera a planilha ativa para a "folha_de_rosto_modelo"
    Sheets("folha_de_rosto_modelo").Activate
    
    'Preenche a folha de rosto
    Range("C1").Value = planejador
    Range("H1").Value = data_atual
    Range("L1").Value = data_necessidade
    Range("D13").Value = projeto
    Range("D15").Value = tr
    Range("K13").Value = remessa
    Range("K15").Value = ordem
    Range("D17").Value = pn
    Range("D19").Value = descricao
    
    'Imprime/Salva em formato PDF a folha de rosto no diretorio indicado
    Sheets("folha_de_rosto_modelo").ExportAsFixedFormat Type:=xlTypePDF, Filename:=diretorio_salvar & "folha_de_rosto_ordem_" & ordem & ".pdf"
    
    'Retorna a planilha ativa para "planilha_ordens" para continuar o processo até a ultima linha preenchida
    Sheets("planilha_ordens").Activate

Next



End Sub


Sub limpa_registros()


'Apaga todos os registros da aba "planilha_ordens"
Sheets("planilha_ordens").Activate
Range("A2:G2").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Clear
Range("A2").Select


End Sub
