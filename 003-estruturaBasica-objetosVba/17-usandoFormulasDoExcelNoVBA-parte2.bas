Attribute VB_Name = "Module1"
Sub empresa()

'Observe que por meio do VBA o que será inserido na célula
'não será a formula mas sim apenas o valor puro
Range("F3").Value = WorksheetFunction.Sum(Range("D:D"))

End Sub

Sub media1()

media = WorksheetFunction.Average(Range("D:D"))
Cells(6, 6) = media


End Sub
