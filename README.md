<div align="center">

## Subclassing without using the AdressOf\-Operator


</div>

### Description

This code simulates subclassing without the AdressOf-Operator. It's much safer than the "SetWindowLong-Method". The code shows a MessageBox when you click on the form (it's only a simple example!)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Stephan Kirchmaier](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/stephan-kirchmaier.md)
**Level**          |Advanced
**User Rating**    |4.8 (24 globes from 5 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/stephan-kirchmaier-subclassing-without-using-the-adressof-operator__1-9620/archive/master.zip)

### API Declarations

```
see code, PS: You must press a button to end the programm and the app must start with Sub Main!
PPS: Vote 4 me , pls.!
```


### Source Code

```
'*****Form1*****'
Option Explicit
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  PostQuitMessage 0&
End Sub
'*****Module1*****'
Option Explicit
Public Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Public Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As msg) As Long
Public Declare Function TranslateMessage Lib "user32" (lpMsg As msg) As Long
Public Declare Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)
Public Type POINTAPI
  x As Long
  y As Long
End Type
Public Type msg
  hwnd As Long
  message As Long
  wParam As Long
  lParam As Long
  time As Long
  pt As POINTAPI
End Type
Public Const PM_REMOVE = &H1
Public Const WM_QUIT = &H12
Public Const WM_RBUTTONDOWN = &H204
Private Sub Main()
  Dim tMsg As msg
  Load Form1
  Form1.Show
  Do
    If PeekMessage(tMsg, 0, 0, 0, PM_REMOVE) Then
      If tMsg.message = WM_QUIT Then Exit Do
      If tMsg.message = WM_RBUTTONDOWN Then
        MsgBox "You clicked the right mousebutton!" & vbCr & "Press a key to end the app"
      End If
      TranslateMessage tMsg
      DispatchMessage tMsg
    Else
      'There's nothing to do for your App!
      'In a game you could draw a new frame,
      'this is much faster than using the Timer!
    End If
  Loop Until False
  Unload Form1
End Sub
```

