Attribute VB_Name = "Módulo1"
Sub cria_abas()


'declarando um array ou vetor no VBA (de Strings)
Dim array_estados(1 To 5) As String

'declarando as outras variaveis utilizadas (opcional)
Dim contador As Integer
Dim nova_aba As Worksheet


'fazendo cada elemento do array receber os valores contidos
'na coluna A da primeira aba do Excel
For contador = 2 To 6

    array_estados(contador - 1) = Cells(contador, 1)
    
Next


'Usando o For para criar as abas desejadas cujo nome estão
'dentro do vetor array_estados
For contador = 1 To 5
    
    'As duas linhas abaixo sao diferentes, pois a que possui a instrucao
    'Before vai criar as abas ANTES da aba atual e a instrucao que possui o
    'After vai criar as abas DEPOIS da aba atual
	'IMPORTANTE: Se Before e After forem omitidos, a nova planilha será inserida antes da planilha ativa.
    
    'Set nova_aba = ThisWorkbook.Worksheets.Add(Before:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
    Set nova_aba = ThisWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
    
    nova_aba.Name = array_estados(contador)
    
Next

Sheets("Dados").Activate

End Sub
