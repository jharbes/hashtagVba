Attribute VB_Name = "Module1"

Sub compilar_extracoes()

resposta = MsgBox("Deseja rodar a macro?", vbOKCancel, "EXECUTA MACRO")

Sheets("Base").Activate

If resposta = 1 Then
    
    'capturando os meses em array
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Base")

    Dim rng As Range
    Set rng = ws.Range("A2:A" & Range("A1").End(xlDown).Row) ' Começa na segunda linha, altere A100 conforme necessário

    Dim cell As Range
    Dim valoresUnicos As Collection
    Set valoresUnicos = New Collection

    ' Adicionando valores únicos na coleção
    On Error Resume Next ' Ignora o erro quando tenta adicionar um item duplicado
    For Each cell In rng
        If Not IsEmpty(cell.Value) Then
            valoresUnicos.Add cell.Value, CStr(cell.Value)
        End If
    Next cell
    On Error GoTo 0 ' Restaura o comportamento normal de erro

    ' Copiando os valores da coleção para um array
    Dim array_mes() As Variant
    If valoresUnicos.Count > 0 Then
        ReDim array_mes(1 To valoresUnicos.Count)
        Dim i As Integer
        For i = 1 To valoresUnicos.Count
            array_mes(i) = valoresUnicos(i)
        Next i

        ' Opção para imprimir os valores únicos no Immediate Window
        For i = 1 To UBound(array_mes)
            Debug.Print array_mes(i)
        Next i
    End If
    
    
    'capturando as plataformas em array
    Dim ws2 As Worksheet
    Set ws2 = ThisWorkbook.Sheets("Base")

    Dim rng2 As Range
    Set rng2 = ws.Range("C2:C" & Range("A1").End(xlDown).Row) ' Começa na segunda linha, altere A100 conforme necessário

    Dim cell2 As Range
    Dim valoresUnicos2 As Collection
    Set valoresUnicos2 = New Collection

    ' Adicionando valores únicos na coleção
    On Error Resume Next ' Ignora o erro quando tenta adicionar um item duplicado
    For Each cell2 In rng2
        If Not IsEmpty(cell2.Value) Then
            valoresUnicos2.Add cell2.Value, CStr(cell2.Value)
        End If
    Next cell2
    On Error GoTo 0 ' Restaura o comportamento normal de erro

    ' Copiando os valores da coleção para um array
    Dim array_plataforma() As Variant
    If valoresUnicos2.Count > 0 Then
        ReDim array_plataforma(1 To valoresUnicos2.Count)
        Dim i2 As Integer
        For i2 = 1 To valoresUnicos2.Count
            array_plataforma(i2) = valoresUnicos2(i2)
        Next i2

        ' Opção para imprimir os valores únicos no Immediate Window
        For i2 = 1 To UBound(array_plataforma)
            Debug.Print array_plataforma(i2)
        Next i2
    End If
    
    
    Dim linha As Integer
    Dim linha2 As Integer
    
    For linha = LBound(array_mes) To UBound(array_mes)
        
        Range("A1").Select
        Selection.AutoFilter
        ActiveSheet.Range("$A$1:$D$" & Range("A1").End(xlDown).Row).AutoFilter Field:=1, Criteria1:=array_mes(linha)
        
        For linha2 = LBound(array_plataforma) To UBound(array_plataforma)
        
            ActiveSheet.Range("$A$1:$D$" & Range("A1").End(xlDown).Row).AutoFilter Field:=3, Criteria1:=array_plataforma(linha2)
            Range("D2").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
            Sheets(array_mes(linha)).Activate
            
            For coluna = 1 To Range("A1").End(xlToRight).Column
                
                If Cells(1, coluna).Value = array_plataforma(linha2) Then
                
                    Cells(2, coluna).Select
                    ActiveSheet.Paste
                            
                End If
                
            Next
            
            Sheets("Base").Activate
            
        Next
        
    Next
    
    Sheets("Base").Activate
    ActiveSheet.ShowAllData
    Selection.AutoFilter
    
    MsgBox "Macro executada com sucesso!", vbInformation, "SUCESSO!"

Else

    MsgBox "Execução abortada!", vbExclamation, "EXECUÇÃO ABORTADA"

End If



End Sub



Sub limpar_registros()

    For Each aba In ThisWorkbook.Sheets
    
        If aba.Index > 1 Then
        
            aba.Activate
            Range("B2:H1048576").ClearContents
        
        End If
    
    Next
    
    Sheets("Base").Activate

End Sub
