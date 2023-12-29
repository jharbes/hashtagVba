Attribute VB_Name = "Module1"
Sub compila_vendas()

Dim fso As Object
Dim pasta As Object
Dim arquivo As Object
Dim caminho_da_pasta As String
Dim mes As Integer


'Defina o caminho da pasta aqui
caminho_da_pasta = ThisWorkbook.Path

'Criar um novo FileSystemObject
'nesse caso usamos o Set pois estamos declarando um OBJETO e nao simplesmente o valor
Set fso = CreateObject("Scripting.FileSystemObject")

'Definir a pasta
Set pasta = fso.GetFolder(caminho_da_pasta)


'Percorrer cada arquivo na pasta
For Each arquivo In pasta.Files

    If InStr(arquivo.Name, "xlsx") > 0 Then
    
        'Aqui você pode fazer algo com cada arquivo
        Debug.Print arquivo.Name
    
    End If

Next arquivo



'Limpar
Set arquivo = Nothing
Set pasta = Nothing
Set fso = Nothing

End Sub
