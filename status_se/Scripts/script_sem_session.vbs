If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject() Then
   Set     = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject ,     "on"
   WScript.ConnectObject application, "on"
End If
.findById("wnd[0]").maximize
.findById("wnd[0]/tbar[0]/okcd").text = "coois"
.findById("wnd[0]").sendVKey 0
.findById("wnd[0]/tbar[1]/btn[17]").press
.findById("wnd[1]/usr/txtV-LOW").text = "status se"
.findById("wnd[1]/usr/txtENAME-LOW").text = ""
.findById("wnd[1]/usr/txtENAME-LOW").setFocus
.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 0
.findById("wnd[1]/tbar[0]/btn[8]").press
.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/btn%_S_FEVOR_%_APP_%-VALU_PUSH").press
.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "051"
.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "052"
.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/btnRSCSEL_255-SOP_I[0,3]").setFocus
.findById("wnd[1]/tbar[0]/btn[8]").press
.findById("wnd[0]/tbar[1]/btn[8]").press
.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").firstVisibleRow = 41
.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").firstVisibleRow = 246
.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").firstVisibleRow = 328
.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").firstVisibleRow = 913
.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").setCurrentCell -1,""
.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectAll
.findById("wnd[1]/tbar[0]/btn[0]").press
.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").firstVisibleRow = 80
.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").firstVisibleRow = 162
.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").firstVisibleRow = 367
.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").firstVisibleRow = 408
.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").firstVisibleRow = 449
.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").firstVisibleRow = 490
.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").firstVisibleRow = 572
.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").firstVisibleRow = 613
.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").firstVisibleRow = 654
.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").firstVisibleRow = 695
.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").firstVisibleRow = 736
.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").firstVisibleRow = 818
.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").firstVisibleRow = 913
