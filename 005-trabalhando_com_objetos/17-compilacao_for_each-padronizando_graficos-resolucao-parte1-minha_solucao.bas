Attribute VB_Name = "Module2"
Sub padronizando_todos_graficos()

Dim aba As Worksheet
Dim grafico As ChartObject

'ActiveSheet = aba atual de utilizacao

For Each aba In ThisWorkbook.Sheets

    aba.Activate
    
    For Each grafico In aba.ChartObjects
    'ou
    'For Each grafico In ActiveSheet.CharObjects
    
    ActiveSheet.ChartObjects(grafico.Name).Activate
    With ActiveSheet.Shapes(grafico.Name).Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent4
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0.8000000119
        .Transparency = 0
        .Solid
    End With
    ActiveSheet.Shapes(grafico.Name).Width = 310
    ActiveSheet.Shapes(grafico.Name).Height = 230
    Application.CommandBars("Format Object").Visible = False
    
    Next

Next

Sheets("Base Vendas").Activate
Range("A2").Select


End Sub
