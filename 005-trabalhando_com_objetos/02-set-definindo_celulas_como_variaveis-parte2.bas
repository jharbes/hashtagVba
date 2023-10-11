Attribute VB_Name = "Módulo1"
Sub formatar()
Attribute formatar.VB_ProcData.VB_Invoke_Func = " \n14"
'
' formatar Macro
'

'
    Range("D2").Select
    Selection.Font.Bold = True
    Selection.Font.Italic = True
    Selection.Font.Underline = xlUnderlineStyleSingle
    Selection.NumberFormat = "$ #,##0.00"
End Sub
Sub venda_minima()

'Vamos setar uma variavel como o intervalo de celulas desejado
Dim intervalo As Range

Set intervalo = Range("D2:D10")

intervalo.Value = 5000
intervalo.Font.Bold = True
intervalo.Font.Italic = True
intervalo.Font.Underline = xlUnderlineStyleSingle
intervalo.NumberFormat = "$ #,##0.00"

End Sub














