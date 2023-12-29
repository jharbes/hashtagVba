Attribute VB_Name = "Module1"
Sub compila_vendas()

'aumentar a velocidade da macro desligando a visualizacao da macro rodando
'importante desligar apos fim do codigo
Application.ScreenUpdating = False

'aumentar a velocidade da macro desligando o calculo automatico na tabela do excel
'importante desligar apos fim do codigo
Application.Calculation = xlCalculationManual


Dim fso As Object
Dim pasta As Object
Dim arquivo As Object
Dim caminho_da_pasta As String
Dim mes As String


'Defina o caminho da pasta aqui
caminho_da_pasta = ThisWorkbook.Path

'Criar um novo FileSystemObject
'nesse caso usamos o Set pois estamos declarando um OBJETO e nao simplesmente o valor
Set fso = CreateObject("Scripting.FileSystemObject")

'Definir a pasta
Set pasta = fso.GetFolder(caminho_da_pasta)


'Percorrer cada arquivo na pasta
For Each arquivo In pasta.Files

    If Right(arquivo.Name, 4) = "xlsx" Then
    
        Set wb = Workbooks.Open(ThisWorkbook.Path & "\" & arquivo.Name)
        
        Range("A2").Select
        Range(Selection, Selection.End(xlToRight)).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy
        
        ThisWorkbook.Activate
        
        linha_disponivel = Range("A1048576").End(xlUp).Row + 1
        Range("A" & linha_disponivel).Select
        ActiveSheet.Paste
        
        Application.CutCopyMode = False
        wb.Close
        
    
    End If

Next arquivo

ultima_linha = Range("A1").End(xlDown).Row

Range("D1").Select
    ActiveWorkbook.Worksheets("Compilado").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Compilado").Sort.SortFields.Add2 Key:=Range( _
        "D2:D" & ultima_linha), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Compilado").Sort
        .SetRange Range("A1:D" & ultima_linha)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
Range("A1").Select


'Limpar
Set arquivo = Nothing
Set pasta = Nothing
Set fso = Nothing


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic


End Sub

