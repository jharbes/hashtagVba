Attribute VB_Name = "Module1"
Sub inputa()

funcionario = InputBox("Digite o nome do Funcionário")
idade = InputBox("Digite a idade do funcionário")
cargo = InputBox("Digite o cargo do novo funcionário")
salario = InputBox("Digite o salário do novo funcionário")

linha_disponivel = Range("A1").End(xlDown).Row + 1

Cells(linha_disponivel, 1).Value = funcionario
Cells(linha_disponivel, 2).Value = idade
Cells(linha_disponivel, 3).Value = cargo
Cells(linha_disponivel, 4).Value = salario


End Sub
