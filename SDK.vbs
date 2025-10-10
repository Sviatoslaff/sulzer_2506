Dim WshShell, SapGuiAuto, application, session, Wnd0, Menubar, UserArea, Statusbar, UserName 

' Create an object WScript.Shell
Set WshShell = WScript.CreateObject("WScript.Shell")

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Connect to SAP
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Create an object
Set SapGuiAuto = GetObject("SAPGUI")

' Create an object type GuiApplication (COM-interface)
Set application = SapGuiAuto.GetScriptingEngine()

' Create an object type GuiSession - session, active SAP window
' i.e starting WSF script itself runs in SAP in the same window
Set session = application.ActiveSession()

WScript.ConnectObject session,     "on"
WScript.ConnectObject application, "on"

' Create an object type GuiMainWindow
Set Wnd0 = session.findById("wnd[0]")

' Create an object type GuiMenubar
Set Menubar = Wnd0.findById("mbar")

' Create an object type GuiUserArea
Set UserArea = Wnd0.findById("usr")

' Create an object type GuiStatusbar
Set Statusbar = Wnd0.findById("sbar")

' Username identification
UserName = session.Info.User

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Auxilary functions
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Transaction start
Sub startTransaction(transaction_name)
    session.StartTransaction transaction_name
End Sub

' Press the "Enter"
Sub pressEnter()
    Wnd0.sendVKey 0
End Sub

' Press the F2
Sub pressF2()
    Wnd0.sendVKey 2
End Sub

' Press the F3
Sub pressF3()
    Wnd0.sendVKey 3
End Sub

' Press the F5
Sub pressF5()
    Wnd0.sendVKey 5
End Sub

' Press the F8
Sub pressF8()
    Wnd0.sendVKey 8
End Sub

' Choosing files
Function selectFile(createOuputFile)
    Set objDialog = CreateObject("UserAccounts.CommonDialog")
    ' Filling parameters and opening
    With objDialog
        .InitialDir = WshShell.SpecialFolders("Desktop") ' Start folder - Desktop
        .Filter = "Text files (*.csv;*.txt)|*.csv;*.txt"
        result = .ShowOpen
    End With
    ' If file is not chosen - exit
    If (result = 0) Then WScript.Quit
    inputFile = objDialog.FileName ' The full path
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set inputStream = fso.OpenTextFile(inputFile)
    ' Create an output file?
    If (createOuputFile) Then
        outputFile = Left(inputFile, Len(inputFile) - 3) & "out" & Right(inputFile, 4)
        Set outputStream = fso.CreateTextFile(outputFile, True)
 ' Returning the array from the read stream from the file and the write stream to the file
        selectFile = Array(inputStream, outputStream)
    Else
        ' Returning the reading stream from the file
        Set selectFile = inputStream
    End If
End Function

' Fill in one row in the table (for ME51N)
Sub fill_row(row, material, kolvo, zavod, zatreboval)
    Set grid = session.findById(UserArea.findByName("GRIDCONTROL", "GuiCustomControl").Id & "/shellcont/shell")
    grid.modifyCell row, "KNTTP", "K"        
    grid.modifyCell row, "MATNR", material   
    grid.modifyCell row, "MENGE", kolvo      
    grid.modifyCell row, "NAME1", zavod      
    grid.modifyCell row, "LGOBE", "0001"     
    grid.modifyCell row, "AFNAM", zatreboval 
End Sub

Function BrowseForFile()
    Const BIF_BROWSEINCLUDEFILES = &H4000 ' Includes files in the dialog
    Const BIF_RETURNONLYFSDIRS = &H1 ' Only allows selection of file system directories

    Set objFSO=CreateObject("Scripting.FileSystemObject") 
    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.BrowseForFolder(0, "Select a file or folder:", BIF_BROWSEINCLUDEFILES)

    If Not objFolder Is Nothing Then
        strTempPath = objFolder.Self.Path
    Else
        strTempPath = ""
    End If

    BrowseForFile = strTempPath

    Set objFolder = Nothing
    Set objShell = Nothing
End Function