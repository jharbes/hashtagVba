Attribute VB_Name = "Module1"
Sub compila_concessionarias()

Dim numero_concessionarias As Integer

Sheets("Concessionárias").Activate
numero_concessionarias = Range("A1").End(xlDown).Row - 1

Dim ultima_linha As Integer
Dim posicao_hifen As Integer
Dim concessionaria As String
Dim linha As Integer
Dim linha2 As Integer
Dim valor As Double
ReDim array_concessionarias(1 To numero_concessionarias) As String

For linha = 2 To numero_concessionarias + 1
    
    posicao_hifen = InStr(Cells(linha, 1).Value, "-")
    concessionaria = Mid(Cells(linha, 1).Value, posicao_hifen + 2)
    array_concessionarias(linha - 1) = concessionaria

Next



Sheets("Resumo").Activate

ultima_linha = Range("A1").End(xlDown).Row

resposta = MsgBox("Deseja rodar a macro?", vbQuestion + vbYesNo, "EXECUTAR MACRO")


If resposta = 6 Then
    

    novos_usados = InputBox("Deseja compilar os carros Novos ou Usados?", "CONFIRMAÇÃO", "Novos/Usados")
    
    If novos_usados = "Novos" Or novos_usados = "Usados" Then
    
        'Apaga todos os registros das abas de concessionarias
        For linha = 1 To numero_concessionarias
    
            Sheets(array_concessionarias(linha) & " - Novos").Activate
            Range("A2:F2").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Clear
            Range("A2").Select
            
            Sheets(array_concessionarias(linha) & " - Usados").Activate
            Range("A2:F2").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Clear
            Range("A2").Select
    
        Next
            
        Sheets("Resumo").Activate
        For linha = 2 To ultima_linha
            
            Sheets("Resumo").Activate
            If novos_usados = "Novos" Then
            
                situacao = "Novo"
                
            Else
            
                situacao = "Usado"
            
            End If
                
            For linha2 = 1 To numero_concessionarias
                unidade = Mid(Cells(linha, 1).Value, InStr(Cells(linha, 1).Value, "-") + 2)
                If unidade = array_concessionarias(linha2) And Cells(linha, 6).Value = situacao Then
                
                    Data = Cells(linha, 2).Value
                    quantidade = Cells(linha, 3).Value
                    carro = Cells(linha, 4).Value
                    valor = Cells(linha, 5).Value
                    tipo = Cells(linha, 6).Value
                    
                    Sheets(array_concessionarias(linha2) & " - " & novos_usados).Activate
                    
                    If Range("A1").End(xlDown).Row <> 1048576 Then
                        linha3 = Range("A1").End(xlDown).Row + 1
                    Else
                        linha3 = 2
                    End If
                    
                    Cells(linha3, 1).Value = unidade
                    Cells(linha3, 2).Value = Data
                    Cells(linha3, 3).Value = quantidade
                    Cells(linha3, 4).Value = carro
                    Cells(linha3, 5).Value = Format(valor, "Currency")
                    Cells(linha3, 6).Value = tipo
                    
                End If
                
        
            Next
        
        Next
        
    Else
    
        resposta2 = MsgBox("Opção inexistente, tente novamente", vbExclamation, "ERRO!")
    
    End If

Else

    resposta1 = MsgBox("Execução cancelada!", vbInformation, "EXECUÇÃO")

End If



End Sub
