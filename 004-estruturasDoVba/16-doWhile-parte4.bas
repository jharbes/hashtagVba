Attribute VB_Name = "Module3"

Sub repeticao2()

Dim linha As Integer
linha = 1

Do While linha <= 10
    
    'Cells(linha, 5).Value = "VBA"
    
    'Abaixo faremos o script utilizando Range em vez de Cells
    'Utilizaremos o operador de concatenacao do VBA "&" para
    'gerar o número da celula correspondente necessaria
    Range("E" & linha).Value = "VBA"
    
    linha = linha + 1

Loop

End Sub


Sub tempo_de_empresa()

Dim linha As Integer
linha = 3


'Vamos escrever até a linha 18 o tempo de empresa de
'cada funcionario
Do While Cells(linha, 2) <> "" And Cells(linha, 3) <> ""
    
    'A formula Date retorna a data atual, entao subtraimos da
    'data de contratacao para saber o tempo de empresa
    'inicialmente em dias, depois dividiremos por 365 para
    'ter o valor em anos
    idade = (Date - Cells(linha, 3).Value) / 365
    
    'Aqui utilizaremos a funcao RoundDown que eh proveniente do
    'excel para arredondar a divisao de cima,
    'por isso ela sera precedida pela WorksheetFunction
    'seus argumentos serao o valor a ser arredondado e o numero
    'de casas decimais desejado após o arredondamento
    idade = WorksheetFunction.RoundDown(idade, 0)
    
    Cells(linha, 4).Value = idade
    
    linha = linha + 1

Loop


End Sub
