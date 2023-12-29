Attribute VB_Name = "Module1"
Sub compila_vendas()

Dim fso As Object
Dim pasta As Object
Dim arquivo As Object
Dim caminho_da_pasta As String
Dim mes As Integer
Dim linha_compilacao As Integer


'Defina o caminho da pasta aqui
caminho_da_pasta = ThisWorkbook.Path

'Criar um novo FileSystemObject
'nesse caso usamos o Set pois estamos declarando um OBJETO e nao simplesmente o valor
Set fso = CreateObject("Scripting.FileSystemObject")

'Definir a pasta
Set pasta = fso.GetFolder(caminho_da_pasta)


'Percorrer cada arquivo na pasta
For Each arquivo In pasta.Files
    
    'InStr(arquivo.Name, "xlsx") retorna o numero do caractere onde se encontra a string
    'procurada, caso contrario retorna o valor zero (0)
    If InStr(arquivo.Name, "xlsx") > 0 Then
        
        'descobrindo a primeira linha disponivel no arquivo "compilado"
        linha_compilacao = Range("A1048576").End(xlUp).Row + 1
        
        'abrindo o arquivo de cada mes
        Workbooks.Open (caminho_da_pasta & "\" & arquivo.Name)
        
        'descobrindo a ultima linha preenchida do arquivo de cada mes
        linha_copia = Range("A1").End(xlDown).Row
        
        'copiando todos os dados dos arquivos de cada mes
        Range("A2:D" & linha_copia).Copy
        
        'alterando o arquivo ativo para o arquivo do script
        ThisWorkbook.Activate
        
        'copiando os dados do arquivo do mes para o arquivo "compilado"
        Range("A" & linha_compilacao).PasteSpecial
        
        'limpando a area de transferencia para conseguir fechar o arquivo sem msgbox de confirmacao
        'tem que estar antes do fechamento do arquivo de mes para impedir que apareça a msgbox
        Application.CutCopyMode = False
        
        'fechando o arquivo do mes
        Workbooks(arquivo.Name).Close
        
        
        
        'Aqui você pode fazer algo com cada arquivo
        Debug.Print arquivo.Name
    
    End If

Next arquivo



'Limpar
Set arquivo = Nothing
Set pasta = Nothing
Set fso = Nothing

End Sub
