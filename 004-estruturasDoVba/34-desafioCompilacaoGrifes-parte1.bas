Attribute VB_Name = "Module2"
Sub compilar_grifes_aprimorado()

Sheets("Produtos").Activate

'Limpa as células da tabela resultado
Range("E6:O11").ClearContents

Dim grife As String
Dim status As String
Dim estoque_minimo As Integer
Dim ultima_linha_base As Integer
Dim linha As Integer
Dim codigo As String
Dim cor As String
Dim linha_cor As Integer
Dim estoque As Integer




grife = Range("B1").Value
status = Range("B2").Value
estoque_minimo = Range("B3").Value


For linha_cor = 6 To Range("D6").End(xlDown).Row

    estoque_cor = 0
    cor_atual = Range("C" & linha_cor).Value
    

    Sheets("Base").Activate
    
    ultima_linha_base = Range("A1").End(xlDown).Row
    
    For linha = 2 To ultima_linha_base
        
        If Range("D" & linha).Value = grife And Range("G" & linha).Value = status And Range("C" & linha).Value = cor_atual Then
        
            codigo = Range("A" & linha).Value
            estoque = Range("F" & linha).Value
            
            Sheets("Produtos").Activate
            
            coluna = 5
            
            Do Until Cells(linha_cor, coluna) = ""
                coluna = coluna + 1
            Loop
            
            estoque_cor = estoque_cor + estoque
            
            Cells(linha_cor, coluna) = codigo
            Cells(linha_cor, 15) = estoque_cor
                
            
            Sheets("Base").Activate
            
            
            
            
        
        End If
    
    Next

Next

    
Sheets("Produtos").Activate
    
For linha = 6 To Range("O6").End(xlDown).Row

    If Range("O" & linha).Value < estoque_minimo Then
    
        Range("E" & linha & ":O" & linha).Interior.Color = RGB(255, 0, 0)
    
    End If
    
Next
    


End Sub

