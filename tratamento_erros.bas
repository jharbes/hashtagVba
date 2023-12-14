On Error GoTo ErrorHandler
' Seu código aqui
Exit Sub
ErrorHandler:
MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical
