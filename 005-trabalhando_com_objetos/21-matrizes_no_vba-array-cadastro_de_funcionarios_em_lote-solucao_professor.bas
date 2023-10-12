Attribute VB_Name = "Módulo2"
Sub cadastra_em_lote()

Dim array_funcionario(33, 7) As String

For lin = 2 To 34
    For col = 1 To 7
        array_funcionario(lin - 2, col - 1) = Sheets("Lote de Funcionários").Cells(lin, col)
    Next
Next

Sheets("Cadastro").Activate

lin_vazia = Range("A1000").End(xlUp).Row + 1

For lin = lin_vazia To lin_vazia + 33
    Cells(lin, 1) = array_funcionario(lin - lin_vazia, 0)
    Cells(lin, 2) = array_funcionario(lin - lin_vazia, 6)
    Cells(lin, 3) = array_funcionario(lin - lin_vazia, 2)
    Cells(lin, 4) = array_funcionario(lin - lin_vazia, 5)
    Cells(lin, 5) = array_funcionario(lin - lin_vazia, 4)
Next

End Sub

