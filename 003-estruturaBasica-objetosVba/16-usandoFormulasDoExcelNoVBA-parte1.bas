Attribute VB_Name = "Module1"
Sub empresa()

'Observe que por meio do VBA o que ser� inserido na c�lula
'n�o ser� a formula mas sim apenas o valor puro
Range("F3").Value = WorksheetFunction.Sum(Range("D:D"))


End Sub
