Attribute VB_Name = "Save_Coid"
Sub SaveCoid()
Attribute SaveCoid.VB_ProcData.VB_Invoke_Func = " \n14"
' This macro will save the coid sheet in the designated directory as an Xlsm type, then as a PDF type, then send the PDF to outlook.
' Created by Antonio Lassalle on 9/21/2024.

' Declare variables.
    Dim DateEntry As Variant
    Dim FilePath As String
    Dim WbName As String
    
    DateEntry = Format(Range("DateEntry"), "MM-DD-YYYY")
    FilePath = "G:\SAP\Inventory Coordinators\IC Log\COID\"
    WbName = "COID" & " " & DateEntry
    
' Enable error trap if workbook exists and user does not want to overwrite.
    On Error Resume Next
    
' Save the file type as xlsm in the g drive.
    With ThisWorkbook
        .SaveAs _
        Filename:=FilePath & WbName, _
        FileFormat:=xlOpenXMLWorkbookMacroEnabled, _
        CreateBackup:=False
    End With
    
    '// Remove formulas
    Sheet1.Cells.Copy
    Sheet1.Range("A1").PasteSpecial (xlPasteValues)
    Application.CutCopyMode = False
    Sheet1.Range("A1").Select
    ThisWorkbook.Save
    
' Reset error trpping behaviour.
    On Error GoTo 0
    
' Save the file type as PDF in the g drive.
    With Sheet1
        .ExportAsFixedFormat Type:=xlTypePDF, _
        Filename:=FilePath & WbName, _
        Quality:=0
    End With
    
' Print the coid sheet.
    With ThisWorkbook.Sheets("MATS")
        .PrintOut
    End With
    
' Declare email variables.
    Dim EmailApp As Object
    Dim EmailItem As Object
    
    Set EmailApp = CreateObject("Outlook.Application")
    Set EmailItem = EmailApp.CreateItem(0)
    
' Set up email in outlook.
    With EmailItem
        .To = ""
        .Subject = "COID" & " " & Format(Range("DateEntry"), "MM/DD/YYYY")
        .Body = ""
        .Attachments.Add FilePath & WbName & ".pdf"
        .Display
    End With

' Declare variables to delete the created PDF.
    Dim FSO
    Dim WbDelete As String
        
    Set FSO = CreateObject("Scripting.FileSystemObject")
    WbDelete = FilePath & WbName & ".pdf"

' Enable error trapping.
    On Error Resume Next
    
' Delete PDF workbook.
    FSO.Deletefile WbDelete, True
    
' Disable error trapping.
    On Error GoTo 0
    
End Sub
