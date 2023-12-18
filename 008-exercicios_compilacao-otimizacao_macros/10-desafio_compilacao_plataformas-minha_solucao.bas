Attribute VB_Name = "Module1"
Sub compilar_extracoes()

resposta = MsgBox("Deseja rodar a macro?", vbOKCancel, "EXECUTA MACRO")

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
    
    
    
    
 
    Sheets("Base").Activate

Else

    MsgBox "Execução abortada!", vbExclamation, "EXECUÇÃO ABORTADA"

End If



End Sub
