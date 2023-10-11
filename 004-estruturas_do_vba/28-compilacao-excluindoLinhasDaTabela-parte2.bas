Attribute VB_Name = "Module1"
Sub compila()

Dim linha As Integer

ultima_linha = Range("A1").End(xlDown).Row

For linha = 2 To ultima_linha

    If Range("F" & linha).Value = "Antigo" Then
        Rows(linha).Delete
        'Rows(linha & ":" & linha).Delete  'outra opcao para o mesmo comando
        
        'Subtrairemos 1 da linha toda vez que o If verdadeiro porque
        'sempre que houver a exclusao da linha a proxima linha mudara
        'sua numeracao (se apagar a linha 12 a 13 passara a ser 12)
        'sendo assim teremos que subtrair 1 da linha para que a nova
        'linha 12 nao fique sem tratamento, caso isso nao seja feito o
        'codigo nao deletara quando houver dois "Antigos" em sequencia
        linha = linha - 1
    End If

Next


End Sub
