Attribute VB_Name = "Módulo1"
Sub compilacao_plataformas()



For Each aba In ThisWorkbook.Sheets
If aba.Index > 1 Then


aba.Activate

Range("B2:H10000").ClearContents
End If

Next

Sheets("Base").Activate

linha = 2

Do Until Cells(linha, 1).Value = ""

mes = Cells(linha, 1).Value
plataforma = Cells(linha, 3).Value
volume = Cells(linha, 4).Value

Sheets(mes).Activate


coluna_plataforma = Cells.Find(plataforma).Column
linha_plataforma = Cells(100000, coluna_plataforma).End(xlUp).Row + 1

Cells(linha_plataforma, coluna_plataforma).Value = volume

Sheets("Base").Activate

linha = linha + 1

Loop

End Sub

Sub limpa_abas()


For Each aba In ThisWorkbook.Sheets
If aba.Index > 1 Then


aba.Activate

Range("B2:H10000").ClearContents
End If

Next
Sheets("Base").Activate


End Sub
