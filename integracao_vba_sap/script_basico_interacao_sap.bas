Sub Macro_PrimariaComando()

Call abrir_sap
Application.Wait Now + TimeValue("00:00:05")
Call executar_sap
        
     MsgBox ("PROCESSAMENTO FINALIZADO")

End Sub

____________________________________________________________________________________

Sub abrir_sap() 
'**IMPORTANTE**: Este scrip é padrão para abertura da tela de login do SAP740 PRD desktop

Dim sapConn As Object

Set objshell = CreateObject("WScript.Shell")
Set objapp = objshell.Exec("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplgpad.exe")
Application.Wait Now + TimeValue("00:00:07")
AppActivate "SAP logon Pad 740"
Application.Wait Now + TimeValue("00:00:05")

Application.SendKeys "PRD", True
Application.Wait Now + TimeValue("00:00:03")
Application.SendKeys "~", True
Application.Wait Now + TimeValue("00:00:07")

End Sub

____________________________________________________________________________________

Sub executar_sap()

Dim Application, SapGuiAuto, Connection, session, WScrip

'**IMPORTANTE**: Abaixo daqui, basta colar o scrip gerado pela gravação do SAP, sem retirar nada:


End Sub