Attribute VB_Name = "Module1"
Sub csv_to_columns()
Attribute csv_to_columns.VB_ProcData.VB_Invoke_Func = " \n14"

'Transformando CSV (ou arquivos onde o separador seja algum caractere) para COLUNAS


'Selecione a coluna em questão (geralmente a coluna A do Excel

'Va no Menu Superior -> 'Data/Dados' -> 'Texto para Colunas/Text to Columns'

'Geralmente será a opção 'Delimited/Delimitado' para delimitação por algum tipo de caractere especial
'(como a virgula)

'Opção 'Vírgula/Comma' para CSV

Dim aba As Worksheet

For Each aba In ThisWorkbook.Sheets
    
    aba.Activate
    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=True, Space:=False, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1)), _
        TrailingMinusNumbers:=True
        
Next
        
End Sub
