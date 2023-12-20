Attribute VB_Name = "Module1"
Sub inputa()

'tratando o erro por GoTo Label
'quando der erro ele pula para o tratamento (tratar:)
On Error GoTo tratar

funcionario = InputBox("Digite o nome do Funcionário")
idade = InputBox("Digite a idade do funcionário")
cargo = InputBox("Digite o cargo do novo funcionário")
salario = InputBox("Digite o salário do novo funcionário")

linha_disponivel = Range("A1").End(xlDown).Row + 1

codigo:

Cells(linha_disponivel, 1).Value = funcionario
Cells(linha_disponivel, 2).Value = idade
Cells(linha_disponivel, 3).Value = cargo
Cells(linha_disponivel, 4).Value = salario

'bloqueia novamente a planilha após rodar a macro
ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True

Exit Sub

tratar:
'os codigos que vamos usar para tratar o erro

'nesse caso teremos que desbloquear a planilha
ActiveSheet.Unprotect

GoTo codigo

End Sub
