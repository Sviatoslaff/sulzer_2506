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
session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode "F00019"
session.findById("wnd[0]/usr/ctxtSO_VKORG-LOW").text = "2080"
session.findById("wnd[0]/usr/ctxtSO_SPART-LOW").text = "21"
session.findById("wnd[0]/usr/ctxtSO_WERKS-LOW").text = "2001"
session.findById("wnd[0]/usr/tabsTABSTRIP_SELTAB/tabpPUSH1/ssub%_SUBSCREEN_SELTAB:ZSD_CSS_DOC_FLOW:0100/ctxtSO_VBELN-LOW").text = "20345050"
session.findById("wnd[0]/usr/tabsTABSTRIP_SELTAB/tabpPUSH1/ssub%_SUBSCREEN_SELTAB:ZSD_CSS_DOC_FLOW:0100/ctxtSO_AUDAT-LOW").text = "05.06.2025"
session.findById("wnd[0]/usr/tabsTABSTRIP_SELTAB/tabpPUSH1/ssub%_SUBSCREEN_SELTAB:ZSD_CSS_DOC_FLOW:0100/ctxtSO_AUDAT-LOW").setFocus
session.findById("wnd[0]/usr/tabsTABSTRIP_SELTAB/tabpPUSH1/ssub%_SUBSCREEN_SELTAB:ZSD_CSS_DOC_FLOW:0100/ctxtSO_AUDAT-LOW").caretPosition = 10
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[1]/btn[43]").press
