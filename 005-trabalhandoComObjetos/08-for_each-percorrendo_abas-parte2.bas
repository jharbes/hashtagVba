Attribute VB_Name = "Módulo1"
Sub percorre_celulas()

'observe que a declaracao da variavel "celula" é opcional
Dim celula As Range

For Each celula In Range("G1:G10")

    celula.Value = "VBA"

Next

End Sub



Sub percorre_abas()

'observe que a declaracao da variavel "aba" é opcional
Dim aba As Worksheet

'observe que a informacao "Sheets" dada na linha abaixo já
'corresponderá a TODAS as ABAS do arquivo
'CUIDADO**: A informacao "Sheets" funcionará para TODOS os arquivos
'de Excel abertos no computador e nao apenas no presente arquivo
'para que sejam utilizadas APENAS AS ABAS DO PRESENTE ARQUIVO
'devemos alterar a informação para "ThisWorkbook.Sheets"
For Each aba In ThisWorkbook.Sheets
    
    aba.Activate
    Cells(2, 1).Value = "Alon"

Next

End Sub




















