Attribute VB_Name = "Module1"
Sub compila()

resposta = MsgBox("deseja realmente executar a macro?", _
vbYesNo + vbQuestion, "CONFIRMAÇÃO")

If resposta = 6 Then

    tipo_de_carro = InputBox("Deseja compilar os carros Novos ou Usados?", "TIPO DE CARRO", "Novo/Usado")
    
    Sheets("Concessionárias").Activate
    
    For linha = 2 To Range("A2").End(xlDown).Row
    
        concessionaria = Cells(linha, 1).Value
        Sheets("Resumo").Activate
        
        'Filtrando a tabela resumo
        ultima_linha_resumo = Range("A1").End(xlDown).Row
        ActiveSheet.Range("$A$1:$F$" & ultima_linha_resumo).AutoFilter Field:=1, Criteria1:= _
        concessionaria
        ActiveSheet.Range("$A$1:$F$" & ultima_linha_resumo).AutoFilter Field:=6, Criteria1:=tipo_de_carro
        
        'Copiando e colando para a tabela designada
        'Para cada concessionária no tipo escolhido
        Range("A1").Select
        Range(Selection, Selection.End(xlToRight)).Select
        Range(Selection, Selection.End(xlDown)).Select
        Application.CutCopyMode = False
        Selection.Copy
        
        'Colando para a aba devida
        nome_concessionaria_isolado = Mid(concessionaria, 7)
        Sheets(nome_concessionaria_isolado & " - " & tipo_de_carro & "s").Select
        Range("A1").Select
        ActiveSheet.Paste
        
    
        Sheets("Concessionárias").Activate
    
    Next

End If


End Sub
