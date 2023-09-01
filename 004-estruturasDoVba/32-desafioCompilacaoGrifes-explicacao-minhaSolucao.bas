Attribute VB_Name = "Module1"
Sub compilar_grifes()

Sheets("Produtos").Activate

'Limpa as células da tabela resultado
Range("E6:O11").ClearContents
Range("D5").Select
Selection.Copy
Range("E6:O11").Select
Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
    SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False

Dim grife As String
Dim status As String
Dim estoque_minimo As Integer
Dim ultima_linha_base As Integer
Dim linha As Integer
Dim codigo As String
Dim cor As String
Dim linha_cor As Integer
Dim estoque As Integer
Dim estoque_amarelo As Integer
Dim estoque_branco As Integer
Dim estoque_azul As Integer
Dim estoque_rosa As Integer
Dim estoque_verde_esmeralda As Integer
Dim estoque_vermelho As Integer

estoque_amarelo = 0
estoque_branco = 0
estoque_azul = 0
estoque_rosa = 0
estoque_verde_esmeralda = 0
estoque_vermelho = 0




grife = Range("B1").Value
status = Range("B2").Value
estoque_minimo = Range("B3").Value


Sheets("Base").Activate

ultima_linha_base = Range("A1").End(xlDown).Row

For linha = 2 To ultima_linha_base
    
    If Range("D" & linha).Value = grife And Range("G" & linha).Value = status Then
    
        codigo = Range("A" & linha).Value
        cor = Range("C" & linha).Value
        estoque = Range("F" & linha).Value
        
        If cor = "AMARELO" Then
            Sheets("Produtos").Activate
            
            linha_cor = 6
            coluna = 5
            
            Do Until Cells(linha_cor, coluna) = ""
                coluna = coluna + 1
            Loop
            
            estoque_amarelo = estoque_amarelo + estoque
            
            Cells(linha_cor, coluna) = codigo
            Cells(linha_cor, 15) = estoque_amarelo
            
            
        ElseIf cor = "BRANCO" Then
            Sheets("Produtos").Activate
            
            linha_cor = 7
            coluna = 5
            
            Do Until Cells(linha_cor, coluna) = ""
                coluna = coluna + 1
            Loop
            
            estoque_branco = estoque_branco + estoque
            
            Cells(linha_cor, coluna) = codigo
            Cells(linha_cor, 15) = estoque_branco
            
        ElseIf cor = "AZUL" Then
            Sheets("Produtos").Activate
            
            linha_cor = 8
            coluna = 5
            
            Do Until Cells(linha_cor, coluna) = ""
                coluna = coluna + 1
            Loop
            
            estoque_azul = estoque_azul + estoque
            
            Cells(linha_cor, coluna) = codigo
            Cells(linha_cor, 15) = estoque_azul
            
        ElseIf cor = "ROSA" Then
            Sheets("Produtos").Activate
            
            linha_cor = 9
            coluna = 5
            
            Do Until Cells(linha_cor, coluna) = ""
                coluna = coluna + 1
            Loop
            
            estoque_rosa = estoque_rosa + estoque
            
            Cells(linha_cor, coluna) = codigo
            Cells(linha_cor, 15) = estoque_rosa
            
        ElseIf cor = "VERDE ESMERALDA" Then
            Sheets("Produtos").Activate
            
            linha_cor = 10
            coluna = 5
            
            Do Until Cells(linha_cor, coluna) = ""
                coluna = coluna + 1
            Loop
            
            estoque_verde_esmeralda = estoque_verde_esmeralda + estoque
            
            Cells(linha_cor, coluna) = codigo
            Cells(linha_cor, 15) = estoque_verde_esmeralda
            
        ElseIf cor = "VERMELHO" Then
            Sheets("Produtos").Activate
            
            linha_cor = 11
            coluna = 5
            
            Do Until Cells(linha_cor, coluna) = ""
                coluna = coluna + 1
            Loop
            
            estoque_vermelho = estoque_vermelho + estoque
            
            Cells(linha_cor, coluna) = codigo
            Cells(linha_cor, 15) = estoque_vermelho
            
        End If
        
        Sheets("Base").Activate
        
        
        
        
    
    End If

Next


Sheets("Produtos").Activate

For linha = 6 To Range("O6").End(xlDown).Row

    If Range("O" & linha).Value < estoque_minimo Then
    
        Range("E" & linha & ":O" & linha).Interior.Color = RGB(255, 0, 0)
    
    End If
    
Next

End Sub
