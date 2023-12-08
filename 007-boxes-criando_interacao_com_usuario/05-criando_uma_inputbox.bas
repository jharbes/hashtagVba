Attribute VB_Name = "Module4"
Sub inputbox_cpf()

Dim cpf As String


cpf = InputBox("Preencha o CPF do Cliente:", "Informe os dados:", "Escreva aqui o cpf")

Range("C15").Value = cpf

MsgBox ("Macro finalizada com sucesso!")

'ou

'Range("C15").Value = InputBox("Preencha o CPF do Cliente:", "Informe os dados:", "Escreva aqui o cpf")


End Sub
