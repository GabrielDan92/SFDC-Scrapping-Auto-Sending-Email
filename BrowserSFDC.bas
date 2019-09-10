Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)

Sub SFDC()

Dim sString As String
Dim sArray() As String
Dim StartTime As Double
Dim MinutesElapsed As String
Dim ie As InternetExplorer
Dim htmlDoc As Object
StartTime = Timer
Dim secondsCounter As Integer


Dim LastRow As Long
    With activesheet
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row                                                                                            'finds the last spreadsheet's row
End With
            
Set ie = New InternetExplorer
ie.Visible = True
ie.Navigate "https://login.salesforce.com/"                                                                                                         'goes to the SFDC's login page

On Error Resume Next
        Set htmlDoc = ie.document
        TabTitle = htmlDoc.Title                                                                                                                    'get the browser tab title
        
    Do While TabTitle = ""                                                                                                                          'wait for the page to load
        DoEvents
        Sleep 1000
        secondsCounter = secondsCounter + 1
        If secondsCounter = 20 Then
            Exit Do
        End If
        Set htmlDoc = ie.document
        TabTitle = htmlDoc.Title                                                                                                                    'get the browser tab title
    Loop
On Error GoTo 0

secondsCounter = 0                                                                                                                                  'reinitialize the secondsCounter variable to 0
ie.document.getElementById("Login").Click                                                                                                           'clicks on the Login SFDC button

For y = 2 To LastRow

    x = Range("C" & y)
    Dim rng As Range
    Set rng = Range("L" & y)
    If IsEmpty(rng.Value) = False Then
        GoTo NextIteration                                                                                                                          'if the cell is not empty, skip it and go to the next row
    End If
    
    On Error Resume Next
        Do While ie.document.querySelector(".searchBoxClearContainer input:first-child") Is Nothing                                                 'wait for the search bar to appear
            DoEvents
            Sleep 1000
            secondsCounter = secondsCounter + 1
            If secondsCounter = 20 Then
                Exit Do
            End If
        Loop
    On Error GoTo 0
    
    secondsCounter = 0                                                                                                                              'reinitialize the secondsCounter variable to 0
    
    ie.document.querySelector(".searchBoxClearContainer input:first-child").Value = x                                                               'enter the wrong email address in the search bar
    
                On Error Resume Next
                Do While ie.document.querySelector _
                ("#phSearchForm .headerSearchContainer .headerSearchLeftRoundedCorner .headerSearchRightRoundedCorner input:first-child") _
                Is Nothing                                                                                                                          'wait for the search button to appear
                    DoEvents
                    Sleep 1000
                    secondsCounter = secondsCounter + 1
                    If secondsCounter = 20 Then
                        Exit Do
                    End If
                Loop
                On Error GoTo 0
    
                secondsCounter = 0                                                                                                                  'reinitialize the secondsCounter variable to 0
    
    Sleep 1500
    ie.document.querySelector _
    ("#phSearchForm .headerSearchContainer .headerSearchLeftRoundedCorner .headerSearchRightRoundedCorner input:first-child").Click                 'click on the search button
    
                On Error Resume Next
                Do While ie.document.querySelector(".searchEntityList .itemLink .item .linkSelector .resultCount") Is Nothing                       'wait for the results count to appear
                    DoEvents
                    Sleep 1000
                    secondsCounter = secondsCounter + 1
                    If secondsCounter = 20 Then
                        Exit Do
                    End If
                Loop
                On Error GoTo 0
    
                secondsCounter = 0                                                                                                                  'reinitialize the secondsCounter variable to 0
                
    Sleep 1000
    Call ieBusy(ie)
    Range("P" & y) = ie.document.querySelector(".searchEntityList .itemLink .item .linkSelector .resultCount").innerHTML                            'how many contacts were found in SFDC
    Range("P" & y) = onlyDigits(Range("P" & y))                                                                                                     'returns the contacts count (only the digits)
    
    If Range("P" & y) = 0 Then                                                                                                                      'if the contact person was not found
        Range("N" & y) = "n/a"
        Range("O" & y) = "n/a"
        Range("M" & y) = "n/a"
        Range("L" & y) = "n/a"
        GoTo NextIteration
    End If
        
    ie.document.querySelector(".dataRow .dataCell a:first-child").Click                                                                             'clicks on the SFDC result
    
    
                On Error Resume Next
                Do While ie.document.querySelector(".dataCol a:first-child") Is Nothing                                                             'wait for the contact person's account to appear
                    DoEvents
                    Sleep 1000
                    secondsCounter = secondsCounter + 1
                    If secondsCounter = 20 Then
                        Exit Do
                    End If
                Loop
                On Error GoTo 0
            
                secondsCounter = 0                                                                                                                  'reinitialize the secondsCounter variable to 0
    
    Range("N" & y) = ie.document.querySelector(".dataCol a:first-child").innerHTML                                                                  'assign the contact person's account to the N row
    Range("O" & y) = ie.document.querySelector(".textBlock h2:first-child").innerHTML                                                               'assign the contact person's name to the O row
    ie.document.querySelector(".dataCol a:first-child").Click                                                                                       'clicks on the contact's account hyperlink
    
                On Error Resume Next
                Do While ie.document.querySelector _
                (".oRight .bPageBlock .pbBody .pbSubsection table:first-child tbody:first-child .dataCol div:first-child span:first-child a:first-child") _
                Is Nothing                                                                                                                          'wait for the account's owner hyperlink to appear
                    DoEvents
                    Sleep 1000
                    secondsCounter = secondsCounter + 1
                    If secondsCounter = 20 Then
                        Exit Do
                    End If
                Loop
                On Error GoTo 0
            
                secondsCounter = 0                                                                                                                  'reinitialize the secondsCounter variable to 0
    
        
    ie.document.querySelector _
    (".oRight .bPageBlock .pbBody .pbSubsection table:first-child tbody:first-child .dataCol div:first-child span:first-child a:first-child").Click 'clicks on the account's owner hyperlink
    
                On Error Resume Next
                Do While ie.document.getElementById("tailBreadcrumbNode") Is Nothing                                                                'wait for the owner's name to appear
                    DoEvents
                    Sleep 1000
                    secondsCounter = secondsCounter + 1
                    If secondsCounter = 20 Then
                        Exit Do
                    End If
                Loop
                On Error GoTo 0
            
                secondsCounter = 0                                                                                                                  'reinitialize the secondsCounter variable to 0
    
    Range("M" & y) = ie.document.getElementById("tailBreadcrumbNode").innerHTML                                                                     'assigns the owner's name to the M row's cell
    Range("L" & y) = ie.document.querySelector(".contactInfo .profileSectionBody .profileSectionData a:first-child").innerHTML                      'assigns the owner's email to the L row's cell
    
    sString = Range("M" & y)
    sArray = Split(sString, "&")
    Range("M" & y) = sArray(0)                                                                                                                      'remove extra characters from the html code
    
NextIteration:
Next y


ie.Quit                                                                                                                                             'close the browser
Set ie = Nothing

MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")                                                                                    'Determine the runtime
MsgBox ("Script completed in: " & MinutesElapsed)

'Call SendingEmail                                                                                                                                  'if I want to send the emails automatically afterwards

End Sub

Sub ieBusy(ie As Object)
    Do While ie.Busy Or ie.READYSTATE <> 4
        DoEvents
    Loop
End Sub

Function onlyDigits(s As String) As String                                                                                                          'returns only the digits from a string with multiple characters
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
