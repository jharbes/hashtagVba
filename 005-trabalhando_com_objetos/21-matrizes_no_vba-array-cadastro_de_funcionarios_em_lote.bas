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

ReDim matriz_origem(1 To ultima_linha_origem, 1 To ultima_coluna_origem) As String


For linha = 2 To ultima_linha_origem

    For coluna = 1 To ultima_coluna_origem
    
        matriz_origem(linha - 1, coluna) = Cells(linha, coluna).Value
    
    Next

Next

Sheets("Cadastro").Activate

ultima_linha_destino = Range("A1").End(xlDown).Row + 1
ultima_coluna_destino = Range("A1").End(xlToRight).Column


For linha = ultima_linha_destino To ultima_linha_destino + ultima_linha_origem - 2

    Range("A2:E2").Select
    Selection.Copy
    Range("A" & linha & ":E" & linha).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False

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
