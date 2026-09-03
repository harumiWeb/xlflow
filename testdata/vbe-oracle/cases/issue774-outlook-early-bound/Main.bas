Attribute VB_Name = "Main"
Option Explicit

Public Sub ProbeIssue774Outlook()
    Dim outlookObj As Outlook.Application
    Dim mailItemObj As Outlook.MailItem
    Dim namespaceObj As Outlook.Namespace
    Dim itemId As String
    Dim storeId As String
    Set outlookObj = CreateObject("Outlook.Application")
    Set namespaceObj = outlookObj.GetNamespace("MAPI")
    Set mailItemObj = namespaceObj.GetItemFromID(itemId, storeId)
    mailItemObj.Subject = "xlflow VBA229 reproduction"
End Sub
