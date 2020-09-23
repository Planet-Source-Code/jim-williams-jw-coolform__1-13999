Attribute VB_Name = "CoolformBas"
Option Explicit

Declare Sub ReleaseCapture Lib "user32" ()
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Type RECT
   left As Integer
   top As Integer
   right As Integer
   bottom As Integer
End Type

Type POINT
   X As Long
   Y As Long
End Type

Declare Sub ClipCursor Lib "user32" (lpRect As Any)
Declare Sub GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT)
Declare Sub ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINT)
Declare Sub OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long)


