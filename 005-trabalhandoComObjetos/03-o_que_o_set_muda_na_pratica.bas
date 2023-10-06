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



Sub cadastro_vendedor()

Dim celula As Range

'o set informa que o valor a ser atribuido a variavel celula devera
'ir para a celula A10, na ausencia dele ele entederá diferente, entende
'que o valor da variavel celula é o informado em A10
Set celula = Range("A10")

celula = "Izabelle"

End Sub














