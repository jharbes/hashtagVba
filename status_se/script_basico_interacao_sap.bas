Sub abrir_sap()

Dim sapConn As Object

'**IMPORTANTE**: Este scrip é padrão para abertura da tela de login do SAP740 PRD desktop

Set objshell = CreateObject("WScript.Shell")
Set objapp = objshell.Exec("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplgpad.exe")
Application.Wait Now + TimeValue("00:00:04")
AppActivate "SAP logon Pad 770"
Application.Wait Now + TimeValue("00:00:03")

Application.SendKeys "PRD", True
Application.Wait Now + TimeValue("00:00:02")
Application.SendKeys "~", True
Application.Wait Now + TimeValue("00:00:04")

End Sub



Sub executar_sap()

Dim Application, SapGuiAuto, Connection, session, WScrip

'**IMPORTANTE**: Abaixo daqui, basta colar o scrip gerado pela gravação do SAP, sem retirar nada:


End Sub



Sub Macro_PrimariaComando()

Call abrir_sap
Application.Wait Now + TimeValue("00:00:03")
Call executar_sap
        
     MsgBox ("PROCESSAMENTO FINALIZADO")

End Sub