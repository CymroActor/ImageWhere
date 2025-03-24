Attribute VB_Name = "MouseWheel"
Option Explicit
'************************************************************
'API
'************************************************************

Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" ( _
    ByVal lpPrevWndFunc As Long, _
    ByVal hwnd As Long, _
    ByVal MSG As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long

'************************************************************
'Constants
'************************************************************

Public Const MK_CONTROL = &H8
Public Const MK_LBUTTON = &H1
Public Const MK_RBUTTON = &H2
Public Const MK_MBUTTON = &H10
Public Const MK_SHIFT = &H4
Private Const GWL_WNDPROC = -4
Private Const WM_MOUSEWHEEL = &H20A

'************************************************************
'Variables
'************************************************************

Private hControl As Long
Private lPrevWndProc As Long

'*************************************************************
'WindowProc
'*************************************************************

Private Function WindowProc(ByVal lWnd As Long, ByVal lMsg As Long, _
ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim MouseKeys As Long
    Dim Rotation As Long
    Dim Xpos As Long
    Dim Ypos As Long

    'Test if the message is WM_MOUSEWHEEL
    If lMsg = WM_MOUSEWHEEL Then
        'Add event handling code here
        'this will be universal to all forms that are 'hooked' to this code
        MouseKeys = wParam And 65535
        Rotation = wParam / 65536
        Xpos = lParam And 65535
        Ypos = lParam / 65536
        Screen.ActiveForm.MouseWheelRolled MouseKeys, Rotation, Xpos, Ypos
    End If
    'Sends message to previous procedure if not MOUSEWHEEL
    'This is VERY IMPORTANT!!!
    If lMsg <> WM_MOUSEWHEEL Then
        WindowProc = CallWindowProc(lPrevWndProc, lWnd, lMsg, wParam, lParam)
    End If
End Function

'*************************************************************
'Hook
'All forms that call this procedure must implement this procedure in their module:
'   Public Sub MouseWheelRolled()
'        <your code>
'   End Sub
'*************************************************************
Public Sub Hook(ByVal hControl_ As Long)
    hControl = hControl_
    lPrevWndProc = SetWindowLong(hControl, GWL_WNDPROC, AddressOf WindowProc)
End Sub
'*************************************************************
'UnHook
'*************************************************************
Public Sub UnHook()
    Call SetWindowLong(hControl, GWL_WNDPROC, lPrevWndProc)
End Sub


