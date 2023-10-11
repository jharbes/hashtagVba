Attribute VB_Name = "Módulo1"
Sub gravacao_padronizado()

For Each aba In ThisWorkbook.Sheets
aba.Activate

    For Each grafico In ActiveSheet.ChartObjects
    
        With ActiveSheet.Shapes(grafico.Name).Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorAccent4
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0.8000000119
            .Transparency = 0
            .Solid
        End With
        ActiveSheet.Shapes(grafico.Name).Height = 226.7716535433
        ActiveSheet.Shapes(grafico.Name).Width = 425.1968503937
        
    Next

Next

End Sub

