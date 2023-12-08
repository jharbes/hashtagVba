Attribute VB_Name = "Module3"
Sub verifica_status()

Dim estoque_atual As Integer
Dim estoque_minimo As Integer

estoque_atual = Range("C7").Value
estoque_minimo = Range("D7").Value

resposta = MsgBox("Deseja realmente rodar a macro?", vbYesNo + vbCritical, "Confirma a execução da macro")

If resposta = 6 Then

    If estoque_atual >= estoque_minimo Then
        
        Range("E7").Value = "Produto OK"
    
    Else
    
        Range("E7").Value = "Produto em Falta"
        
    End If
    
    resposta1 = MsgBox("Macro executada com sucesso!", vbInformation)
    
Else

    resposta2 = MsgBox("Execução da macro abortada!", vbExclamation)

End If


End Sub
