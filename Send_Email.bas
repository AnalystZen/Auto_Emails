Attribute VB_Name = "Send_Email"
Sub SendEmail()
Attribute SendEmail.VB_ProcData.VB_Invoke_Func = " \n14"
' This macro will send the designated file to outlook via email.
' Created by Antonio Lassalle on 09/21/2024.

' Declare variables.
    Dim EmailApp As Object
    Dim EmailItem As Object
    Dim DateEntry As Variant
    Dim FilePath As String
    Dim WbName As String

    Set EmailApp = CreateObject("Outlook.Application")
    Set EmailItem = EmailApp.CreateItem(0)
    DateEntry = Format(Range("DateEntry"), "MM-DD-YYYY")
    FilePath = "G:\SAP\Inventory Coordinators\IC Log\"
    WbName = "COID" & " " & DateEntry

' Set up email in outlook.
    With EmailItem
        .To = ""
        .Subject = "COID" & " " & Format(Range("DateEntry"), "MM/DD/YYYY")
        .Body = ""
        .Attachments.Add FilePath & WbName & ".pdf"
        .Display
    End With

End Sub
