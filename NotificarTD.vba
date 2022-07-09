'Script Creado por Zamir Pineda
Sub Notificacion_TD_11()

'Declarar variables
Dim Application
Dim Connection
Dim Filas As Long
Dim i As Long
Filas = ThisWorkbook.Sheets("TG").Range("G2").CurrentRegion.Rows.Count
For i = 2 To Filas

'Conectar con SAP
If Not IsObject(Application) Then
   Set SapGuiAuto = GetObject("SAPGUI")
   Set Application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(Connection) Then
   Set Connection = Application.Children(0)
End If
If Not IsObject(session) Then
   Set session = Connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session, "on"
   WScript.ConnectObject Application, "on"
End If

'Limpiar SAP
session.findById("wnd[0]").resizeWorkingPane 198, 32, False
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]").resizeWorkingPane 198, 32, False
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
'Notificación TD 11
session.findById("wnd[0]/tbar[0]/okcd").text = "COR6N"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ssubSUB01:SAPLCORU_S:0010/subSLOT_HDR:SAPLCORU_S:5117/ctxtAFRUD-AUFNR").text = Sheets("TG").Cells(i, 7).Value
session.findById("wnd[0]/usr/ssubSUB01:SAPLCORU_S:0010/subSLOT_HDR:SAPLCORU_S:5117/ctxtAFRUD-VORNR").text = "11"
session.findById("wnd[0]/usr/ssubSUB01:SAPLCORU_S:0010/subSLOT_HDR:SAPLCORU_S:5117/ctxtAFRUD-VORNR").SetFocus
session.findById("wnd[0]/usr/ssubSUB01:SAPLCORU_S:0010/subSLOT_HDR:SAPLCORU_S:5117/ctxtAFRUD-VORNR").caretPosition = 2
session.findById("wnd[0]").sendVKey 11
Next i

MsgBox "Datos enviados correctamente"

End Sub

'Script Creado por Zamir Pineda
Sub Notificacion_TD_21()

'Declarar variables
Dim Application
Dim Connection

If Not IsObject(Application) Then
   Set SapGuiAuto = GetObject("SAPGUI")
   Set Application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(Connection) Then
   Set Connection = Application.Children(0)
End If
If Not IsObject(session) Then
   Set session = Connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session, "on"
   WScript.ConnectObject Application, "on"
End If

session.findById("wnd[0]").resizeWorkingPane 198, 32, False
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/okcd").text = "ZPP_POM_2057_1"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/btn%_ORD_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]").sendVKey 24
session.findById("wnd[1]").sendVKey 8
session.findById("wnd[0]").sendVKey 8
session.findById("wnd[0]").sendVKey 3

Call Abrir_Revision_TD

End Sub

'Script Creado por Zamir Pineda
Sub Abrir_Revision_TD()

'Conectar con SAP
Dim Application
Dim Connection

If Not IsObject(Application) Then
   Set SapGuiAuto = GetObject("SAPGUI")
   Set Application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(Connection) Then
   Set Connection = Application.Children(0)
End If
If Not IsObject(session) Then
   Set session = Connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session, "on"
   WScript.ConnectObject Application, "on"
End If

session.findById("wnd[0]").resizeWorkingPane 198, 32, False
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/okcd").text = "coid"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/radREP_COMP").Select
session.findById("wnd[0]/usr/radREP_COMP").SetFocus
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtP_PROFID").text = "000001"
session.findById("wnd[0]/usr/ctxtP_LAYOUT").text = "//CONSUMOS"
session.findById("wnd[0]/usr/ctxtS_AUFNR-HIGH").SetFocus
session.findById("wnd[0]/usr/ctxtS_AUFNR-HIGH").caretPosition = 0
session.findById("wnd[0]/usr/btn%_S_AUFNR_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]").sendVKey 24
session.findById("wnd[1]").sendVKey 8
session.findById("wnd[0]").sendVKey 8

End Sub

'Script Creado por Zamir Pineda
Sub Cierre_Tecnico_TD()

'Conectar con SAP
Dim Application
Dim Connection

If Not IsObject(Application) Then
   Set SapGuiAuto = GetObject("SAPGUI")
   Set Application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(Connection) Then
   Set Connection = Application.Children(0)
End If
If Not IsObject(session) Then
   Set session = Connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session, "on"
   WScript.ConnectObject Application, "on"
End If

session.findById("wnd[0]").resizeWorkingPane 198, 32, False
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[0]/okcd").text = "COID"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtP_PROFID").text = "000001"
session.findById("wnd[0]/usr/ctxtP_PROFID").SetFocus
session.findById("wnd[0]/usr/ctxtP_PROFID").caretPosition = 6
session.findById("wnd[0]/usr/btn%_S_AUFNR_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]").sendVKey 24
session.findById("wnd[1]").sendVKey 24
session.findById("wnd[1]").sendVKey 8
session.findById("wnd[0]").sendVKey 8
session.findById("wnd[0]/usr/cntlGRID_0100/shellcont/shell").setCurrentCell -1, ""
session.findById("wnd[0]/usr/cntlGRID_0100/shellcont/shell").SelectAll
session.findById("wnd[0]").sendVKey 32
session.findById("wnd[1]/usr/subFUNCTION_SETUP:SAPLCOWORK:0200/cmbCOWORK_FCT_SETUP-FUNCT").Key = "220"
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]").sendVKey 8


End Sub
