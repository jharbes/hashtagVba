Attribute VB_Name = "Module1"
Sub empresa()

'Gasto de salário da empresa?

'Observe que por meio do VBA o que será inserido na célula
'não será a formula mas sim apenas o valor puro
Range("F3").Value = WorksheetFunction.Sum(Range("D:D"))

End Sub

Sub media1()

'Média dos salários da empresa

media = WorksheetFunction.Average(Range("D:D"))
Cells(6, 6) = media


End Sub

Sub estagiario()

'Quantos estagiarios existem na empresa?

'Formula no excel: CONT.SE(C:C;"Estagiário")

Cells(9, 6).Value = WorksheetFunction.CountIf(Range("C:C"), "Estagiário")


End Sub
