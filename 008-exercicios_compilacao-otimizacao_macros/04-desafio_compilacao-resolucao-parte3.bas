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
        
        ultima_linha_resumo = Range("A2").End(xlDown).Row
        ActiveSheet.Range("$A$1:$F$" & ultima_linha_resumo).AutoFilter Field:=1, Criteria1:= _
        concessionaria
        ActiveSheet.Range("$A$1:$F$" & ultima_linha_resumo).AutoFilter Field:=6, Criteria1:=tipo_de_carro
    
        
    
    
    
        Sheets("Concessionárias").Activate
    
    Next

End If


End Sub
