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
session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode "F00015"
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = "20345050"
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").caretPosition = 8
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[24]").press
