Attribute VB_Name = "Módulo1"
Sub limpar_filtros()

On Error Resume Next

For Each aba In ThisWorkbook.Sheets

    aba.ShowAllData

Next

End Sub
