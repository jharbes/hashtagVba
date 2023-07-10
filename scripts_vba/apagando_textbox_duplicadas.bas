Attribute VB_Name = "Module1"
Sub DeleteTextBoxes()
    Dim sh As Worksheet
    Dim obj As Object
    Set sh = ThisWorkbook.Sheets("Sheet1") ' Substitua "Sheet1" pelo nome da sua planilha

    For Each obj In sh.Shapes
        If obj.Type = msoTextBox Then
            If obj.Name = "TextBox 164" Or obj.Name = "TextBox 176" Then
                obj.Delete
            End If
        End If
    Next obj
End Sub
