Attribute VB_Name = "Module1"
Sub cadastrar_lote_minha_solucao()

Dim linha As Integer

Dim ultima_linha_origem As Integer
Dim ultima_coluna_origem As Integer

Dim ultima_linha_destino As Integer
Dim ultima_coluna_destino As Integer

Sheets("Lote de funcionários").Activate

ultima_linha_origem = Range("A1").End(xlDown).Row
ultima_coluna_origem = Range("A1").End(xlToRight).Column

'No VBA o indice inicial é opcional, nesse caso estamos colocando como 1,
'mas caso nada fosse declarado o indice inicial seria 0 (zero)
ReDim matriz_origem(1 To ultima_linha_origem, 1 To ultima_coluna_origem) As String


'Aqui capturamos os dados dos funcionarios na variavel "matriz_origem" que estao na
'aba "Lote de funcionários"
For linha = 2 To ultima_linha_origem

    For coluna = 1 To ultima_coluna_origem
    
        matriz_origem(linha - 1, coluna) = Cells(linha, coluna).Value
    
    Next

Next

'Alteramos a aba ativa
Sheets("Cadastro").Activate

'descobrimos a ultima linha em branco
ultima_linha_destino = Range("A1").End(xlDown).Row + 1
ultima_coluna_destino = Range("A1").End(xlToRight).Column


'Agora com esse for preenchemos a aba "Cadastro" com os funcionarios encontrados na aba
'"Lote de funcionários"
For linha = ultima_linha_destino To ultima_linha_destino + ultima_linha_origem - 2
	
	'Para cada linha nossa primeira açao será copiar a formatação da tabela geral para a nova
	'linha a ser preenchida
    Range("A2:E2").Select
    Selection.Copy
    Range("A" & linha & ":E" & linha).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
	
	'Agora preenchemos a planilha aplicando uma logica com "Select Case" para acerto das colunas
    For coluna = 1 To ultima_coluna_destino
    
        Select Case coluna
        
            Case 1
            
                Cells(linha, coluna) = matriz_origem(linha - ultima_linha_destino + 1, coluna)
                    
            Case 2
            
                Cells(linha, coluna) = matriz_origem(linha - ultima_linha_destino + 1, 7)
                
            Case 3
            
                Cells(linha, coluna) = matriz_origem(linha - ultima_linha_destino + 1, coluna)
                
            Case 4
            
                Cells(linha, coluna) = matriz_origem(linha - ultima_linha_destino + 1, 6)
                
            Case 5
            
                Cells(linha, coluna) = matriz_origem(linha - ultima_linha_destino + 1, coluna)
                
        End Select
    
    Next

Next

End Sub
