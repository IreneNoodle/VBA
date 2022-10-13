Attribute VB_Name = "Module1"
Sub PrepareEmails()


Dim ol As Outlook.Application
Dim olmail As Outlook.MailItem

Set ol = New Outlook.Application

For i = 9 To Sheet1.Cells(Rows.Count, 1).End(xlUp).Row

    'Create emial item for each row
    Set olmail = ol.CreateItem(olMailItem)
    With olmail
        .SentOnBehalfOfName = "xxx.xxx@mail.com"
        .To = Sheet1.Cells(i, 1).Value
        .Subject = Sheet1.Cells(i, 2).Value
        .CC = Sheet1.Cells(i, 3).Value
        .Body = Sheet1.Cells(i, 4).Value
        .Display
    End With


Next


End Sub



