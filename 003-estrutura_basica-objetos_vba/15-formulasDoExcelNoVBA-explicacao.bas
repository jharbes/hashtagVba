'Usando Fórmulas do Excel no VBA

WorksheetFunction.Sum(...) 'onde:

WorksheetFunction. 'Para chamar as formulas do excel
Sum 'Nome da fórmula (em ingles), exemplos: Sum, Average, Max, Countif,Vlookup
(...) 'Argumentos da formula (com a linguagem do VBA), exemplos: Range("A1:B3").Value, Cells(7,4).Value, etc


'Ex:

Range("A1").Value = WorksheetFunction.Sum(Range("A2:A5").Value)

media = WorksheetFunction.Average(Range("B7:C10").Value)