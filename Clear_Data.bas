Attribute VB_Name = "Clear_Data"
Sub ClearData()
Attribute ClearData.VB_Description = "This macro will clear data if the workbook is accidently saved over."
Attribute ClearData.VB_ProcData.VB_Invoke_Func = " \n14"
' This macro will clear data if the workbook is accidently saved over.

    
' Enable error trapping.
    On Error GoTo Errhandler:
    
' Turn off screen updating.
    Application.ScreenUpdating = False
    
'Hide and clear COID data.
    With ThisWorkbook.Sheets("COID")
        .Visible = True
        .Activate
        Cells.ClearContents
        Range("A1").Select
        .Visible = False
    End With
    
' Select main worksheet.
    ThisWorkbook.Sheets("MATS").Activate
    
' Clear any formatting from the "MATS" Sheet if any was applied.
    With Range("B5:G100")
        .Borders.LineStyle = xlNone
    End With
   
' Apply line border.
    With Range("CookieHeader")
        .BorderAround xlContinuous, xlThin
        .Borders(xlInsideVertical).LineStyle = xlContinuous
    End With
    
' Apply line border.
    With Range("CrackerHeader")
        .BorderAround xlContinuous, xlThin
        .Borders(xlInsideVertical).LineStyle = xlContinuous
    End With
    
' Select cell A1 at end of macro.
    Application.GoTo Reference:=Range("A1"), Scroll:=True
    
' Turn on screen updating.
    Application.ScreenUpdating = True

' Clean exit
    Exit Sub

Errhandler:
' User update of failure.
    MsgBox Prompt:="Clearing of data failed!", Buttons:=vbCritical, Title:="Date Wipe Failed"
' Turn on screen update.
    Application.ScreenUpdating = True
End Sub
