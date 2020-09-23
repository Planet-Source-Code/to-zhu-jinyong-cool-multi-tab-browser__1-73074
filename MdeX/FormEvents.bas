Attribute VB_Name = "FormEvents"
Option Explicit
Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal lngHandle As Long, ByVal lngMsg As Long, ByVal lngFirstParam As Long, ByVal lngLastParam As Long) As Long
Declare Function SetWindowLongA Lib "user32" (ByVal lngHandle As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Sub InitCommonControls Lib "comctl32" ()

Public lngOldProc As Long
Public Sub SetProc(ByVal lngHandle As Long)
lngOldProc = SetWindowLongA(lngHandle, -4, AddressOf WinProc)
End Sub
Private Function WinProc(ByVal lngHandle As Long, ByVal lngMsg As Long, ByVal lngFirstParam As Long, ByVal lngLastParam As Long) As Long
WinProc = CallWindowProcA(lngOldProc, lngHandle, lngMsg, lngFirstParam, lngLastParam)
End Function
