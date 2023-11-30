Attribute VB_Name = "Module1"



Sub capturar_codigo()

'Exemplo funcao Left

Dim linha As Integer
Dim ultima_linha As Integer
Dim codigo As String

Sheets("Fórmulas de Texto - Parte 1").Activate

ultima_linha = Range("B2").End(xlDown).Row

For linha = 3 To ultima_linha

'codigo recebe os oito caracteres a esquerda da string
codigo = Left(Cells(linha, 2).Value, "8")
Cells(linha, 4).Value = codigo


Next

End Sub




Sub capturar_estado()

'Exemplo funcao Right

Dim linha As Integer
Dim ultima_linha As Integer
Dim estado As String

Sheets("Fórmulas de Texto - Parte 2").Activate

ultima_linha = Range("B2").End(xlDown).Row

For linha = 3 To ultima_linha

'estado recebe os ultimos dois caracteres a direita
estado = Right(Cells(linha, 2).Value, "2")
Cells(linha, 4).Value = estado


Next


End Sub



Sub capturar_descricao()

'Exemplo funcao Mid

' Mid("string",x,y)
' x = a partir de que caractere pegar o resultado
' y = numero de caracateres a partir de x (opcional)

Dim linha As Integer
Dim ultima_linha As Integer
Dim descricao As String

Sheets("Fórmulas de Texto - Parte 3").Activate

ultima_linha = Range("B2").End(xlDown).Row

For linha = 3 To ultima_linha

'descricao recebe os caracteres da string excluindo os 11 primeiros
'da esquerda (ou todos a partir do 12o caractere
descricao = Mid(Cells(linha, 2).Value, 12)
Cells(linha, 4).Value = descricao


Next


End Sub
