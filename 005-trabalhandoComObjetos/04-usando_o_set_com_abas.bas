Attribute VB_Name = "Module1"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    Sheets("Exemplo Aba Verde").Select
    Sheets("Exemplo Aba Verde").Move Before:=Sheets(1)
    With ActiveWorkbook.Sheets("Exemplo Aba Verde").Tab
        .Color = 5296274
        .TintAndShade = 0
    End With
End Sub


Sub exercicio()

'declarando aba como uma variavel do tipo "Worksheet" (aba)
Dim aba As Worksheet

'setando a aba "aba" como a aba "Exemplo Aba Verde"
Set aba = Sheets("Exemplo Aba Verde")

aba.Select
aba.Move Before:=Sheets(1)

'Temos que tirar o "ActiveWorkbook" nesse caso pois o VBA já está com essa
'aba como ativa
With aba.Tab
    .Color = 5296274
    .TintAndShade = 0
End With


End Sub
