Attribute VB_Name = "Module1"
Sub inputa()

funcionario = InputBox("Digite o nome do Funcion�rio")
idade = InputBox("Digite a idade do funcion�rio")
cargo = InputBox("Digite o cargo do novo funcion�rio")
salario = InputBox("Digite o sal�rio do novo funcion�rio")

linha_disponivel = Range("A1").End(xlDown).Row + 1

Cells(linha_disponivel, 1).Value = funcionario
Cells(linha_disponivel, 2).Value = idade
Cells(linha_disponivel, 3).Value = cargo
Cells(linha_disponivel, 4).Value = salario


End Sub
