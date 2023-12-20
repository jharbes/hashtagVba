Attribute VB_Name = "Module1"
Sub percentual_finalizado()

'indica que caso haja alguem erro basta seguir em frente rodando
'o restante da macro
On Error Resume Next

Dim linha As Integer
Dim ultima_linha As Integer

linha = 2

Do Until Cells(linha, 1).Value = ""
    
    Cells(linha, 4).Value = Cells(linha, 3) / Cells(linha, 2)
    
    linha = linha + 1

Loop


End Sub
