Attribute VB_Name = "Módulo1"
Sub matrizes()

'Aqui estamos declarando a variavel matriz como "Variant" o que significa
'que ela pode assumir qualquer tipo
Dim matriz As Variant


'No VBA o indice inicial é opcional, nesse caso estamos colocando como 1,
'mas caso nada fosse declarado o indice inicial seria 0 (zero)
'ReDim matriz(1 To 27)

'For linha = 1 To 27
'    matriz(linha) = Cells(linha + 1, 1).Value
'Next

'matriz (1)

'observe que ao alimentar a variavel matriz com Range ele criará
'uma matriz com o numero de dimensoes de acordo com o range escolhido
'(nesse caso duas dimensoes)
matriz = Range("A2:B28").Value

Debug.Print matriz(2, 1)

estado = Cells(2, 1).Value

End Sub
