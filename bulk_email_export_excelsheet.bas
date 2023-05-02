Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1

Sub bulk_mail_excel_ab()

'Setting up the Excel variables.
Dim olApp As Object
Dim olMailItm As Object
  
Dim recipient As String
Dim attach As String
Dim bodyMail As String
Dim msgt As String

Sheets("list").Select
Cells(1, 1).Select

Do Until IsEmpty(ActiveCell.Offset(1, 0).Value) 'Find row = "END"
ActiveCell.Offset(1, 0).Select      'Select 1 row more
Loop

Dim lastTableRow As Double      'STORE VALUE WHERE ADDRESS.ROW = "END"
lastTableRow = ActiveCell.Row   'GET LAST ROW ADDRESS WHERE = "END" STRING

Cells(1, 1).Select


For i = 2 To lastTableRow

recipient = Sheets(1).Cells(i, 1).Text
attach = Sheets(1).Cells(i, 2).Text
'msgt = Sheets(1).Cells(i, 3).Value

'Create the Outlook application and the empty email.
Set olApp = CreateObject("Outlook.Application")
Set olMailItm = olApp.CreateItem(0)

'Set oNewMail = Application.CreateItem(olMailItem)
With olMailItm
.Subject = "Test Envoi massif emails APLIANONYME, avec pièce jointe nominative"
.HTMLBody = "Bonjour, " & Sheets(1).Cells(i, 5).Text & " " & Sheets(1).Cells(i, 4).Text & " <br> " & "J’envois un publipostage d’essais dans des buts d’amélioration " & "<br>" & "et (ussi pour répondre à Thierry sur les envois en lot d’emails et les envois de pièces jointes nominatives." & "<table border=2><tr><th>" & "Cest Un Tableau" & "</th><th>" & "Cest Un Tableau" & "</th><th>" & "Cest Un Tableau" & "</th><th>" & "Cest Un Tableau" & "</th></tr><tr><td>" & "Cest Un Tableau" & "</td><td>" & "Cest Un Tableau" & "</td><td></td><td></td></table>"
'.HTMLBody = GetHTML("C:\Users\AH\3.htm")
.To = recipient
.Display
.Attachments.Add attach
.Send
End With

Next i
   
'Clean up the Outlook application.
Set olMailItm = Nothing
Set olApp = Nothing
End Sub


Function GetHTML(URL As String) As String
    Dim HTML As String
    With CreateObject("MSXML2.XMLHTTP")
        .Open "GET", URL, False
        .Send
        GetHTML = .ResponseText
    End With
End Function



Sub createmail()
Dim oNewMail As Outlook.MailItem

'Workbooks("email_list.xlsx").Activate
'Sheets("list").Select
'Cells(1, 1).Select

email_list = Array("toto@consodo.es", "ext.marta.birgitz@gmx.du")

For Each email_dest In email_list
Set oNewMail = Application.CreateItem(olMailItem)
With oNewMail
.Subject = "Rambo des bacs a sable"
.body = "Rambo des bacs a sable"
.To = email_dest
.Display
.Attachments.Add "C:\Users\tata\AHC.docx"
'.Send
End With

Next email_dest

End Sub


