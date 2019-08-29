Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)

Sub SFDC()

Dim sString As String
Dim sArray() As String
Dim StartTime As Double
Dim MinutesElapsed As String
Dim ie As InternetExplorer
Dim html As HTMLDocument
StartTime = Timer

'Call ESDsheet 'for customizing and generating a new spreadsheet

'finds the last spreadsheet's row
Dim LastRow As Long
With activesheet
    LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
End With
            
Set ie = New InternetExplorer
ie.Visible = True
ie.Navigate "https://login.salesforce.com/"     'goes to the SFDC's login page

Do While ie.READYSTATE <> READYSTATE_COMPLETE
    DoEvents
Loop

Call ieBusy(ie)
ie.Document.getElementById("Login").Click   'clicks on the Login SFDC button
Call ieBusy(ie)
Sleep 5000
Application.Wait (Now() + TimeValue("00:00:05"))


    For y = 2 To LastRow
    
        'x = range("F" & y)
        x = Range("C" & y)
        
        Dim rng As Range
        Set rng = Range("L" & y)
        
        If IsEmpty(rng.Value) = False Then
            GoTo NextIteration  'if the cell is not empty, skip it and go to the next row
        End If
        
        ie.Document.querySelector(".searchBoxClearContainer input:first-child").Value = x   'enter the wrong email address in the search bar
        Sleep 1000
        ie.Document.querySelector("#phSearchForm .headerSearchContainer .headerSearchLeftRoundedCorner .headerSearchRightRoundedCorner input:first-child").Click    'clicks on the search button
        Call ieBusy(ie)
        Sleep 1000
        Range("P" & y) = ie.Document.querySelector(".searchEntityList .itemLink .item .linkSelector .resultCount").innerHTML    'how many contacts were found in SFDC
        Range("P" & y) = onlyDigits(Range("P" & y))     'returns the contacts count (only the digits)
        
            If Range("P" & y) = 0 Then  'if the contact person was not found
                Range("N" & y) = "n/a"
                Range("O" & y) = "n/a"
                Range("M" & y) = "n/a"
                Range("L" & y) = "n/a"
                GoTo NextIteration
            End If
            
        ie.Document.querySelector(".dataRow .dataCell a:first-child").Click     'clicks on the SFDC result
        Call ieBusy(ie)
        Sleep 1000
        Range("N" & y) = ie.Document.querySelector(".textBlock h2:first-child").innerHTML   'assign the contact person's name to the N row's cell
        ie.Document.querySelector(".dataCol a:first-child").Click   'clicks on the contact's account hyperlink
        Sleep 3000
        Range("O" & y) = ie.Document.querySelector(".textBlock h2:first-child").innerHTML   'assign the contact person's account to the O row's cell
        Sleep 500
        ie.Document.querySelector(".oRight .bPageBlock .pbBody .pbSubsection table:first-child tbody:first-child .dataCol div:first-child span:first-child a:first-child").Click   'clicks on the account's owner hyperlink
        Sleep 3000
        Range("M" & y) = ie.Document.getElementById("tailBreadcrumbNode").innerHTML     'assigns the owner's name to the M row's cell
        Range("L" & y) = ie.Document.querySelector(".contactInfo .profileSectionBody .profileSectionData a:first-child").innerHTML      'assigns the owner's email to the L row's cell
        
        sString = Range("M" & y)
        sArray = Split(sString, "&")
        Range("M" & y) = sArray(0) 'remove extra characters from the html code
        
NextIteration:
    Next y

Sleep 500

MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")     'Determine the runtime
MsgBox ("Script completed in: " & MinutesElapsed)

'Call SendingEmail '>> if I want to send the emails automatically afterwards

End Sub





Sub ieBusy(ie As Object)
    Do While ie.Busy Or ie.READYSTATE <> 4
        DoEvents
    Loop
End Sub

Function onlyDigits(s As String) As String 'returns only the digits from a string with multiple characters
    Dim retval As String
    Dim i As Integer
    retval = ""
    For i = 1 To Len(s)
        If Mid(s, i, 1) >= "0" And Mid(s, i, 1) <= "9" Then
            retval = retval + Mid(s, i, 1)
        End If
    Next
    onlyDigits = retval
End Function
