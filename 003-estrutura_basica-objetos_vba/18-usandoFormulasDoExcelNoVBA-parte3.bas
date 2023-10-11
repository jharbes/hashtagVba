Attribute VB_Name = "Module1"
Sub empresa()

'Gasto de sal�rio da empresa?

'Observe que por meio do VBA o que ser� inserido na c�lula
'n�o ser� a formula mas sim apenas o valor puro
Range("F3").Value = WorksheetFunction.Sum(Range("D:D"))

End Sub

Sub media1()

'M�dia dos sal�rios da empresa

media = WorksheetFunction.Average(Range("D:D"))
Cells(6, 6) = media


End Sub

Sub estagiario()

'Quantos estagiarios existem na empresa?

'Formula no excel: CONT.SE(C:C;"Estagi�rio")

Cells(9, 6).Value = WorksheetFunction.CountIf(Range("C:C"), "Estagi�rio")


End Sub
