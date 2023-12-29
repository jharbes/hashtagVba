Sub loop_scrooldown_sap()

'Loop pegando cada item da tabela do SAP, fazendo scroll down de página em página e alimentando a tabela do excel
For count = 0 To Division - 1
  For j = 0 To 13
   
   session.findById("wnd[0]/usr/tblSAPML02BD0103").verticalScrollbar.Position = count * 14
   If Val(session.findById("wnd[0]/usr/tblSAPML02BD0103/txtLTBP-TBPOS[0," & j & "]").Text) = 0 Then
    Sheet1.Activate
    lblOrdem.Caption = "Ordem: " & Sheet3.Range("K8").Text
    
    GoTo Line2
   End If
   Cells(j + 5 + (count * 14), 1).Value = session.findById("wnd[0]/usr/tblSAPML02BD0103/txtLTBP-TBPOS[0," & j & "]").Text
   Cells(j + 5 + (count * 14), 2).Value = session.findById("wnd[0]/usr/tblSAPML02BD0103/ctxtLTBP-MATNR[1," & j & "]").Text
   Cells(j + 5 + (count * 14), 3).Value = session.findById("wnd[0]/usr/tblSAPML02BD0103/txtMLVS-MAKTX[2," & j & "]").Text
   Cells(j + 5 + (count * 14), 4).Value = session.findById("wnd[0]/usr/tblSAPML02BD0103/txtLTBP-MENGA[5," & j & "]").Text
   Cells(j + 5 + (count * 14), 5).Value = Val(Replace(Replace(session.findById("wnd[0]/usr/tblSAPML02BD0103/txtLTBP-BRGEW[31," & j & "]").Text, ".", "", 1), ",", ".", 1))
   Cells(j + 5 + (count * 14), 6).Value = session.findById("wnd[0]/usr/tblSAPML02BD0103/txtLTBP-GEWEI[32," & j & "]").Text
   
   If rst.RecordCount > 0 Then
    rst.Find "Item = '" & session.findById("wnd[0]/usr/tblSAPML02BD0103/txtLTBP-TBPOS[0," & j & "]").Text & "'"
    Cells(j + 5 + (count * 14), 8) = "Já solicitado (" & rst.Fields("Solicitação") & ")"
    Range("H" & j + 5 + (count * 14)).Interior.ColorIndex = 6
   End If
   
  Next
Next

End Sub