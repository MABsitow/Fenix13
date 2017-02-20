Attribute VB_Name = "modMouseWheel"
Option Explicit

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Public Const GWL_WNDPROC = -4

Private Const WM_MOUSEWHEEL = &H20A

Private Const WHEEL_DELTA = 120
Private Const WHEEL_PAGESCROLL = &HFFFFFFFF

Public Const SPI_GETWHEELSCROLLLINES = 104

' store a pointer to the form object
' which is set via ObjPtr
Public lpFormObj As Long

Public Function WndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Dim objForm As frmCrearPersonaje
On Error GoTo errorHandler

If uMsg = WM_MOUSEWHEEL Then
    'If the flexGrid is the active control then
     'If TypeOf frmCrearPersonaje.ActiveControl Is PictureBox Then
        ' ##### Scroll direction #####
          If (HiWord(wParam) / WHEEL_DELTA) < 0 Then
            'Scrolling down
            'Debug.Print "Down"
            ' instantiate the pointer we have to the form
            Set objForm = PtrToForm(lpFormObj)
            ' call the method
            objForm.ScrollDown
            ' destroy the reference
            Set objForm = Nothing
        Else
            'Scrolling up
            'Debug.Print "UP"
            ' instantiate the pointer we have to the form
            Set objForm = PtrToForm(lpFormObj)
            ' call the method
            objForm.ScrollUp
            ' destroy the reference
            Set objForm = Nothing
        End If
    'End If
    ' ##### Paging = suggested number of lines to scroll (e.g. in a textbox) #####
    ' Windows 95: Not supported
    Dim r As Long

    SystemParametersInfo SPI_GETWHEELSCROLLLINES, 0, r, 0
        
    If r = WHEEL_PAGESCROLL Then
        'Wheel roll should be interpreted as clicking
        'once in the page down or page up regions of
        'the scroll bar
    Else
        'Scroll 3 lines (3 is the default value)
    End If
    
    'Pass the message to default window procedure and then onto the parent
    DefWindowProc hwnd, uMsg, wParam, lParam
Else
    'No messages handled, call original window procedure
    WndProc = CallWindowProc(GetProp(frmCrearPersonaje.hwnd, "PrevWndProc"), hwnd, uMsg, wParam, lParam)
End If

Exit Function
errorHandler:
Debug.Print Err.number & " " & Err.Description

End Function

Public Function HiWord(dw As Long) As Integer

If dw And &H80000000 Then
    HiWord = (dw \ 65535) - 1
Else
    HiWord = dw \ 65535
End If

End Function

Public Function PtrToForm(ByVal lPtr As Long) As frmCrearPersonaje
'//--[PtrToForm]--------------------------------//
'
'  Creates a dummy object from an ObjPtr

Dim Obj As frmCrearPersonaje

' instantiate the illegal referece
CopyMemory Obj, lPtr, 4
Set PtrToForm = Obj
CopyMemory Obj, 0&, 4

End Function

