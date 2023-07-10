#If VBA7 Then
  Private Declare PtrSafe Function getFrequency Lib "kernel32" _
  Alias "QueryPerformanceFrequency" (cyFrequency As Currency) As Long
  Private Declare PtrSafe Function getTickCount Lib "kernel32" _
  Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long
#Else
  Private Declare Function getFrequency Lib "kernel32" _
  Alias "QueryPerformanceFrequency" (cyFrequency As Currency) As Long
  Private Declare Function getTickCount Lib "kernel32" _
  Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long
#End If
Function MicroTimer() As Double

' Segundos
    Dim cyTicks1 As Currency
    Static cyFrequency As Currency
    MicroTimer = 0

' Busca frequência
    If cyFrequency = 0 Then getFrequency cyFrequency

' Busca os ticks
    getTickCount cyTicks1

' Segundos
    If cyFrequency Then MicroTimer = cyTicks1 / cyFrequency
End Function

Sub TempoArea()
    DoCalcTimer 1
End Sub
Sub TempoAba()
    DoCalcTimer 2
End Sub
Sub TempoPlanilha()
    DoCalcTimer 3
End Sub
Sub TempoExcel()
    DoCalcTimer 4
End Sub
Sub TempoTodasAbas()
    DoCalcTimer 5
End Sub

Sub DoCalcTimer(jMethod As Long)
    Dim dTime As Double
    Dim dOvhd As Double
    Dim oRng As Range
    Dim oCell As Range
    Dim oArrRange As Range
    Dim sCalcType As String
    Dim lCalcSave As Long
    Dim bIterSave As Boolean

    On Error GoTo Errhandl

' Inicia
    dTime = MicroTimer

    ' Salva as configurações de cálculo
    lCalcSave = Application.Calculation
    bIterSave = Application.Iteration
    If Application.Calculation <> xlCalculationManual Then
        Application.Calculation = xlCalculationManual
    End If
    Select Case jMethod
    Case 1

        ' Desliga interações.
        If Application.Iteration <> False Then
            Application.Iteration = False
        End If
        
        ' Max range usado.
        If Selection.Count > 1000 Then
            Set oRng = Intersect(Selection, Selection.Parent.UsedRange)
        Else
            Set oRng = Selection
        End If

        ' Incluir matrizes fora da área selecionada.
        For Each oCell In oRng
            If oCell.HasArray Then
                If oArrRange Is Nothing Then
                    Set oArrRange = oCell.CurrentArray
                End If
                If Intersect(oCell, oArrRange) Is Nothing Then
                    Set oArrRange = oCell.CurrentArray
                    Set oRng = Union(oRng, oArrRange)
                End If
            End If
        Next oCell

        sCalcType = "Cálculo de " & CStr(oRng.Count) & _
            " célula(s) na área selecionada: "
    Case 2
        sCalcType = "Cálculo da aba " & ActiveSheet.Name & ": "
    Case 3
        sCalcType = "Cálculo da planilha atual: "
    Case 4
        sCalcType = "Cálculo completo das planilhas abertas: "
    Case 5
        sCalcType = "Cálculo de cada aba da planilha atual: "
    End Select

' Busca tempo de início
    dTime = MicroTimer
    Select Case jMethod
    Case 1
        If Val(Application.Version) >= 12 Then
            oRng.CalculateRowMajorOrder
        Else
            oRng.Calculate
        End If
    Case 2
        ActiveSheet.Calculate
    Case 3
        Application.Calculate
    Case 4
        Application.CalculateFull
    Case 5
        Dim WS As Worksheet
        
        For Each WS In ThisWorkbook.Worksheets
            WS.Calculate
            dTime = MicroTimer - dTime
            dTime = Round(dTime, 5)
            MsgBox sCalcType & vbNewLine & WS.Name & ": " & CStr(dTime) & " segundos", vbOKOnly + vbInformation, "TempoCalc"
            dTime = MicroTimer
        Next WS
        
        ' Reestabelece métodos de cálculo
        If Application.Calculation <> lCalcSave Then
             Application.Calculation = lCalcSave
        End If
        If Application.Iteration <> bIterSave Then
             Application.Calculation = bIterSave
        End If
        Exit Sub
        
    End Select

' Duração do cálculo
    dTime = MicroTimer - dTime
    On Error GoTo 0

    dTime = Round(dTime, 5)
    MsgBox sCalcType & " " & vbNewLine & CStr(dTime) & " segundos", _
        vbOKOnly + vbInformation, "TempoCalc"

Finish:

    ' Reestabelece métodos de cálculo
    If Application.Calculation <> lCalcSave Then
         Application.Calculation = lCalcSave
    End If
    If Application.Iteration <> bIterSave Then
         Application.Calculation = bIterSave
    End If
    Exit Sub
Errhandl:
    On Error GoTo 0
    MsgBox "Incapaz de calcular " & sCalcType, _
        vbOKOnly + vbCritical, "TempoCalc"
    GoTo Finish
End Sub