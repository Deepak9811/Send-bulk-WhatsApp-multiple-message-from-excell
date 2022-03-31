# Getting Started with Create Excell file

first go to Developer option -> Visual Basic(left side of tool bar) -> then Paste whole code 

## Code 

Sub WhatsAppMsg()

Dim LastRow As Long
Dim I As Integer
Dim strip As String
Dim strPhoneNumber As String
Dim StrMessage As String
Dim strPostData As String
Dim IE As Object

LastRow = Range("A" & Rows.Count).End(xlUp).Row

For I = 2 To LastRow

    strPhoneNumber = Sheets("Data").Cells(I, 1).Value
    StrMessage = Sheets("Data").Cells(I, 2).Value
    
'IE.navigate "whatsapp://send?phone=phone_number&text=your_message"

    strPostData = "whatsapp://send?phone=" & strPhoneNumber & "&text=" & StrMessage
    Set IE = CreateObject("InternetExplorer.Application")
    IE.navigate strPostData
    Application.Wait Now() + TimeSerial(0, 0, 5)
    SendKeys "~"
    
    
Next I

End Sub





######


#### Finish



The Developer tab isn't displayed by default, but you can add it to the ribbon.

1. On the File tab, go to Options > Customize Ribbon.

2. Under Customize the Ribbon and under Main Tabs, select the Developer check box.

