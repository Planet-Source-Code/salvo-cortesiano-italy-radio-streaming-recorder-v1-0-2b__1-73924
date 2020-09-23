Attribute VB_Name = "modOnTop"
Option Explicit

'Dichiarazioni API x AlwaysOnTop
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOREDRAW = &H8
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long


Function AlwaysWindowsOnTop(hwnd As Long, AlwaysOnTop As Boolean) As Boolean
    On Local Error Resume Next
    If AlwaysOnTop Then
        AlwaysWindowsOnTop = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOREDRAW Or SWP_NOSIZE Or SWP_NOMOVE)
    Else
        AlwaysWindowsOnTop = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOREDRAW Or SWP_NOSIZE Or SWP_NOMOVE)
    End If
End Function
