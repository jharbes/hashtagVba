Attribute VB_Name = "Módulo1"
Sub uboundVBA()

'UBOUND = Up Bound (limite superior da matriz)
'LBOUND = Low Bound (limite inferior da matriz)

Dim matrizNomes As Variant


'A variavel matrizNomes recebe os valores da coluna A, da linha 2 até a linha 201
matrizNomes = Range("A2:A201").Value


'o Debug.Print printa na verificacao imediata o que foi solicitado
Debug.Print UBound(matrizNomes)
Debug.Print LBound(matrizNomes)


'Nesse caso ao chamar a funcao LBound() ela colocará o valor 1 (limite inferior da matriz)
' e ao chamar a funcao UBound() ele colocará o valor 200 (limite superior da matriz)
'funcoes boas para manipular matrizes cujos limites sao desconhecidos de inicio
For i = LBound(matrizNomes) To UBound(matrizNomes)
    Cells(i + 1, 2).Value = i
Next i

End Sub
