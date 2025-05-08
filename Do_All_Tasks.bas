Attribute VB_Name = "Do_All_Tasks"
Sub AllTasks()
' This Macro will import coid, filter, print, save and email coid sheet for mats. Apps used- SAP,Outlook,Excel,Windows System.
' Created by Antonio Lassalle on 09/22/2024.

' Variable declarations.
    Dim DateEntry As Variant
    
    DateEntry = Range("DateEntry").Value

' Check to proceed with macro.
    If DateEntry = "" Then
        MsgBox Prompt:="Please enter the date, and try again.", Buttons:=vbOKCancel, Title:="Please Enter Date"
        Exit Sub
    End If

' Enable error trapping.
    On Error GoTo Errhandler:

' Establish SAP Connection.
    Set SapGuiAuto = GetObject("SAPGUI")
    Set SAPApp = SapGuiAuto.GetScriptingEngine
    Set Connection = SAPApp.Children(0)
    Set session = Connection.Children(0)

    If IsObject(WScript) Then
    WScript.ConnectObject session, "on"
    WScript.ConnectObject Application, "on"
    End If

' SAP t code selection. "COID"
    session.findById("wnd[0]").resizeWorkingPane 94, 28, False
    session.StartTransaction "COID"
    
    session.findById("wnd[0]/usr/radREP_HEADER").Select
    session.findById("wnd[0]/tbar[0]/btn[0]").press
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/tbar[1]/btn[17]").press
    session.findById("wnd[1]/usr/txtENAME-LOW").Text = "mittonr"
    session.findById("wnd[1]/usr/txtENAME-LOW").SetFocus
    session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 7
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellRow = 5
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "5"
    session.findById("wnd[1]/tbar[0]/btn[2]").press
    session.findById("wnd[0]/usr/ctxtS_ECKST-LOW").Text = DateEntry
    session.findById("wnd[0]/usr/ctxtS_ECKST-LOW").SetFocus
    session.findById("wnd[0]/usr/ctxtS_ECKST-LOW").caretPosition = 10
    session.findById("wnd[0]").sendVKey 8
    session.findById("wnd[0]/usr/cntlGRID_0100/shellcont/shell").setCurrentCell -1, ""
    session.findById("wnd[0]/usr/cntlGRID_0100/shellcont/shell").SelectAll
' Export data to excel sheet.
    session.findById("wnd[0]/usr/cntlGRID_0100/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
    session.findById("wnd[0]/usr/cntlGRID_0100/shellcont/shell").selectContextMenuItem "&PC"
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").SetFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[0]/tbar[0]/btn[3]").press
    session.findById("wnd[0]/tbar[0]/btn[3]").press
    session.findById("wnd[0]/tbar[0]/btn[3]").press
    
' Diable screen refresh.
    Application.ScreenUpdating = False

' Format COID in excel sheet and unhide/hide.
    With ThisWorkbook.Sheets("COID")
        .Visible = True
        .Activate
        Columns("A:O").Select
        Selection.Clear
        Range("A1").Select
        ActiveSheet.Paste
        Columns("A:A").Select
' Format text to columns.
        Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
            :="|", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
            1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12 _
            , 1), Array(13, 1), Array(14, 1), Array(15, 1)), TrailingMinusNumbers:=True
' Hide COID sheet.
        .Visible = False
    End With
    
' Maximize window screen for user.
    ThisWorkbook.Sheets("MATS").Activate
    ThisWorkbook.Application.WindowState = xlMaximized
        
' Turn on screen update.
    Application.ScreenUpdating = True
    
' Save sheet, print and email.
    Call SaveCoid
    
' Clean exit.
    Exit Sub
    
Errhandler:
' User update of failure.
    MsgBox Prompt:="Import of data failed! Verify a seesion of SAP is open and try again.", Buttons:=vbCritical, Title:="Import Failed"
' Turn on screen update.
    Application.ScreenUpdating = True

End Sub
