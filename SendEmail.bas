Attribute VB_Name = "SendEmail"
Sub SendingEmail()

'finds the last spreadsheet's row
        Dim LastRow As Long
            With activesheet
                LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
            End With



For x = 2 To LastRow

    Dim rng As range
    Set rng = range("L" & x)
    
        If InStr(rng.Value, "n/a") > 0 Then
            GoTo NextIteration
        End If
        
            mailTo = range("L" & x)
            SON = range("A" & x)
            ownerName = range("M" & x)
            contactEmail = range("F" & x)
            contactEmail2 = range("G" & x)
            ContactName = range("N" & x)
            contactAccount = range("O" & x)
    
            Dim outlookApp As Outlook.Application
            Dim myMail As Outlook.MailItem
    
            Set outlookApp = New Outlook.Application
            Set myMail = outlookApp.CreateItem(olMailItem)
    
                myMail.To = mailTo
                'myMail.CC = ""
                myMail.Subject = "Undeliverable e-mail address; SON: " & SON
                'myMail.SentOnBehalfOfName = ""
    
                myMail.HTMLBody = "Hi " & ownerName & ", <br><p>" _
                & "I am reaching out to notify you that a contact on one of your accounts was used on an order and the e-mail address is incorrect. " _
                & "During the fulfillment process, we found that this e-mail address returned as an undeliverable e-mail address. " _
                & "Can you please review the contact info below and confirm if this is correct or please let us know what the correct e-mail address is? " _
                & "This is holding up distribution of orders and potential revenue implications. <br> <p>" _
                & "Contact Name: " & ContactName & " <br>" _
                & "Undeliverable E-mail(s): " & contactEmail & " " & contactEmail2 & " <br>" _
                & "Contact Account: " & contactAccount & " <br><p>" _
                & " <br>" _
                & "Best regards, <br><p>" _
                & "Gabriel Pintoiu <br>" _
                & "Global Business Operations"
    
    
                myMail.Display
                myMail.send

NextIteration:
Next x

End Sub



