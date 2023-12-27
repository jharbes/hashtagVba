Attribute VB_Name = "Module1"
Sub registrar_venda()

'aumentar a velocidade da macro desligando a visualizacao da macro rodando
'importante desligar apos fim do codigo
Application.ScreenUpdating = False

'aumentar a velocidade da macro desligando o calculo automatico na tabela do excel
'importante desligar apos fim do codigo
Application.Calculation = xlCalculationManual


Dim linha As Integer
Dim ultima_linha As Integer
Dim preco As Double
Dim quantidade_estoque As Integer

resposta = MsgBox("Deseja rodar a macro?", vbYesNo + vbQuestion, "EXECUTAR MACRO")


Sheets("Dados").Activate
ultima_linha = Range("A1").End(xlDown).Row
ReDim array_marcas(ultima_linha - 1)

For linha = 2 To ultima_linha
    
    array_marcas(linha - 1) = Cells(linha, 1).Value

Next

If resposta = 6 Then

    marca = InputBox("Qual marca da moto vendida?", "MARCA DA MOTO")
    Do Until IsInArray(marca, array_marcas)
    
        MsgBox ("Marca não reconhecida, favor tentar novamente!")
        marca = InputBox("Qual marca da moto vendida?", "MARCA DA MOTO")
    
    Loop
    
    For linha = 2 To ultima_linha
        
        If marca = Cells(linha, 1) Then
        
            preco = Cells(linha, 2).Value
        
        End If
    
    Next
    
    Data = Date
    
    Set wb_estoque = Workbooks.Open(ThisWorkbook.Path & "\08-exercicio_estoque-explicacao-estoque.xlsm")
    
    ultima_linha = Range("A1").End(xlDown).Row
    
    For linha = 2 To ultima_linha
    
        If marca = Cells(linha, 1) Then
        
            quantidade_estoque = Cells(linha, 2).Value
        
        End If
    
    Next
    
    wb_estoque.Close
    
    Sheets("Vendas Diárias").Activate

    ultima_linha = Range("A1").End(xlDown).Row + 1
    
    Cells(ultima_linha, 1).Value = Cells(ultima_linha - 1, 1).Value + 1
    Cells(ultima_linha, 2).Value = Data
    Cells(ultima_linha, 3).Value = marca
    Cells(ultima_linha, 4).Value = preco
    Cells(ultima_linha, 5).Value = quantidade_estoque
    
    If quantidade_estoque > 0 Then
    
        Cells(ultima_linha, 6).Value = "Disponível"
    
    Else
    
        Cells(ultima_linha, 6).Value = "Indisponível"
    
    End If
    
    resposta2 = MsgBox("Macro executada com sucesso!", vbInformation, "EXECUÇÃO COM SUCESSO")
    
Else

    resposta2 = MsgBox("Execução cancelada com sucesso!", vbInformation, "EXECUÇÃO CANCELADA")

End If


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub


Function IsInArray(valToBeFound As Variant, arr As Variant) As Boolean

    Dim element As Variant
    
    On Error Resume Next ' Em caso de erro (por exemplo, se arr não for um array), a próxima linha causará erro
    
    For Each element In arr
        If element = valToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next element
    
    IsInArray = False ' Se o valor não for encontrado
End Function
