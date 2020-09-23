<div align="center">

## Broadcast Over MSN


</div>

### Description

This code will Send a message to all open MSN Conversation Windows. Please comment and vote!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[ï¿½e7eN](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/e7en.md)
**Level**          |Beginner
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/e7en-broadcast-over-msn__1-34570/archive/master.zip)





### Source Code

```
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Const XP = "RichEdit20W"
Const Win98 = "RichEdit20A"
'If you cant Get it to work useing these two, then get an API Spyer
'and click on the IM's chat box then replace that value with that one
Sub SendText(Text As String)
Dim IMWindow, RichTB, RichTB2, SendButton As Long
IMWindow = FindWindow("IMWindowClass", vbNullString) 'Get IM's Hwnd
If IMWindow = 0 Then Exit Sub 'if no Im's open then exit
RichTB = FindWindowEx(IMWindow, 0, XP, vbNullString) ' Get Chat Rooms Hwnd
RichTB2 = FindWindowEx(IMWindow, RichTB, XP, vbNullString) 'Get Chat Box Hwnd
SendButton = FindWindowEx(IMWindow, 0, "Button", "&Send") 'Get Send Button Hwnd
SendMessageByString RichTB2, &HC, 0, Text 'Get Send Buttons Hwnd
Call SendMessage(SendButton, &H100, &H20, 0&) 'Click the Button
Call SendMessage(SendButton&, &H101, &H20, 0&)
End Sub
```

