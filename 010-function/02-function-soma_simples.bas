Attribute VB_Name = "Module1"
Function soma_simples(x As Double, y As Double) As Double
    
    MsgBox ("Início do uso da formula soma_simples")
    soma_simples = x + y

End Function

'Depois podemos usar no proprio excel essa function utilizando
'=soma_simples(arg1,arg2) direto na propria celula ou barra
'de formulas

'Para visualizar as informações sobre a formula existem duas maneiras:

'Podemos escrever na celula a forma e depois de abrir os parenteses
'ex: "=soma_simples(" apertaremos CTRL + SHIFT + A

'Também podemos ir na barra de fórmulas colocar a fórmula
'"=soma simples(" e depois clicar em "fx"


