Attribute VB_Name = "Module1"
Sub inputa()

'tratando o erro por GoTo Label
'quando der erro ele pula para o tratamento (tratar:)
On Error GoTo tratar

funcionario = InputBox("Digite o nome do Funcion�rio")
idade = InputBox("Digite a idade do funcion�rio")
cargo = InputBox("Digite o cargo do novo funcion�rio")
salario = InputBox("Digite o sal�rio do novo funcion�rio")

linha_disponivel = Range("A1").End(xlDown).Row + 1

codigo:

Cells(linha_disponivel, 1).Value = funcionario
Cells(linha_disponivel, 2).Value = idade
Cells(linha_disponivel, 3).Value = cargo
Cells(linha_disponivel, 4).Value = salario

'bloqueia novamente a planilha ap�s rodar a macro
ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True

Exit Sub

tratar:
'os codigos que vamos usar para tratar o erro

'nesse caso teremos que desbloquear a planilha
ActiveSheet.Unprotect

GoTo codigo

End Sub
