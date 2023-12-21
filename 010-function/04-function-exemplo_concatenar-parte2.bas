Attribute VB_Name = "Módulo1"
Function minhasoma(num1 As Double, num2 As Double) As Double

minhasoma = num1 + num2



End Function


Function meu_concatenar(intervalo_celulas As Range) As String
    
    Dim texto As String
    texto = ""
    
    'percorre todas as celulas do intervalo
    For Each celula In intervalo_celulas
        
        If texto <> "" Then
            texto = texto & " " & celula.Value
        Else
            texto = texto & celula.Value
        End If
    
    Next
    
    meu_concatenar = texto

End Function
