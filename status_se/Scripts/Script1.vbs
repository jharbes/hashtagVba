If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "coois"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/txtV-LOW").text = "status se"
session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
session.findById("wnd[1]/usr/txtENAME-LOW").setFocus
session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 0
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/btn%_S_FEVOR_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "051"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "052"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/btnRSCSEL_255-SOP_I[0,3]").setFocus
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").firstVisibleRow = 41
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").firstVisibleRow = 246
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").firstVisibleRow = 328
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").firstVisibleRow = 913
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").setCurrentCell -1,""
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectAll
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").firstVisibleRow = 80
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").firstVisibleRow = 162
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").firstVisibleRow = 367
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").firstVisibleRow = 408
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").firstVisibleRow = 449
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").firstVisibleRow = 490
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").firstVisibleRow = 572
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").firstVisibleRow = 613
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").firstVisibleRow = 654
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").firstVisibleRow = 695
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").firstVisibleRow = 736
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").firstVisibleRow = 818
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").firstVisibleRow = 913
