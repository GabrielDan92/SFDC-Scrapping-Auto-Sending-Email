Attribute VB_Name = "BrowserSFDC"
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)

Sub SFDC()

Call ESDsheet

'finds the last spreadsheet's row
        Dim LastRow As Long
            With activesheet
                LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
            End With
            

Dim ie As InternetExplorer
Dim html As HTMLDocument
'Dim activesheet As Worksheet
Set ie = New InternetExplorer

ie.Visible = True
ie.Navigate "https://login.salesforce.com/"

Do While ie.READYSTATE <> READYSTATE_COMPLETE
    DoEvents
Loop
Call ieBusy(ie)
ie.Document.getElementById("Login").Click
Call ieBusy(ie)
Sleep 5000
Application.Wait (Now() + TimeValue("00:00:016"))

    For y = 3 To LastRow
    
        x = range("F" & y)
        ie.Document.querySelector(".searchBoxClearContainer input:first-child").Value = x
        Sleep 5000
        ie.Document.querySelector("#phSearchForm .headerSearchContainer .headerSearchLeftRoundedCorner .headerSearchRightRoundedCorner input:first-child").Click
        Call ieBusy(ie)
        Sleep 5000
        ie.Document.querySelector(".dataRow .dataCell a:first-child").Click
        Call ieBusy(ie)
        Sleep 5000
        range("N" & y) = ie.Document.querySelector(".textBlock h2:first-child").innerHTML
        ie.Document.querySelector(".dataCol a:first-child").Click
        Sleep 5000
        range("O" & y) = ie.Document.querySelector(".textBlock h2:first-child").innerHTML
        ie.Document.querySelector(".dataCol a:first-child").Click
        Sleep 5000
        'range("K" & y) = ie.Document.querySelector(".headerContent .chatterBreadcrumbs").getElementsByTagName("span")(2).innerHTML
        range("M" & y) = ie.Document.getElementById("tailBreadcrumbNode").innerHTML
        range("L" & y) = ie.Document.querySelector(".contactInfo .profileSectionBody .profileSectionData a:first-child").innerHTML
        
    Next y


MsgBox ("Charlie Oscar Mike")

End Sub


Sub ieBusy(ie As Object)
    Do While ie.Busy Or ie.READYSTATE <> 4
        DoEvents
    Loop
End Sub
