Attribute VB_Name = "Resize"
Option Explicit

Private colHWnd As New Collection
    
Private Const GWL_WNDPROC = -4
Private Const WM_GETMINMAXINFO = &H24

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type

Private Declare Function DefWindowProc Lib "user32" Alias _
   "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias _
   "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
    ByVal hWnd As Long, ByVal msg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias _
   "SetWindowLongA" (ByVal hWnd As Long, _
    ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Sub CopyMemoryToMinMaxInfo Lib "kernel32" Alias _
   "RtlMoveMemory" (hpvDest As MINMAXINFO, ByVal hpvSource As Long, _
    ByVal cbCopy As Long)
Private Declare Sub CopyMemoryFromMinMaxInfo Lib "kernel32" Alias _
   "RtlMoveMemory" (ByVal hpvDest As Long, hpvSource As MINMAXINFO, _
    ByVal cbCopy As Long)

Public Sub Hook(hWnd As Long, _
                lngWidth As Long, _
                lngHeight As Long)
    
    Dim objResize As New ResizeClass
    If gblnInhibitSubClassing Then Exit Sub
    '
    '   Add the forms references to the collection.
    '
    With objResize
        .hWnd = hWnd
        .Width = lngWidth
        .Height = lngHeight
        .PrevWndProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProc)
''        .WidthPixels = lngWidthPixels
''        .HeightPixels = lngHeightPixels
    End With
    colHWnd.Add objResize, CStr(hWnd)
    
    'Start subclassing.
    Set objResize = Nothing
End Sub

Public Sub Unhook(hWnd As Long)
    Dim temp As Long

    If gblnInhibitSubClassing Then Exit Sub
    'Cease subclassing.
    temp = SetWindowLong(hWnd, GWL_WNDPROC, colHWnd.Item(CStr(hWnd)).PrevWndProc)
    colHWnd.Remove CStr(hWnd)
End Sub

Function WindowProc(ByVal hw As Long, _
                    ByVal uMsg As Long, _
                    ByVal wParam As Long, _
                    ByVal lParam As Long) As Long
    
    Dim MinMax As MINMAXINFO
    Dim nWidthPixels&, nHeightPixels&
    
''    nWidthPixels = MDIForm1.Width \ Screen.TwipsPerPixelX
''    nHeightPixels = MDIForm1.Height \ Screen.TwipsPerPixelY
''    nWidthPixels = colHWnd.Item(CStr(hw)).WidthPixels
''    nHeightPixels = colHWnd.Item(CStr(hw)).HeightPixels
    
    'Check for request for min/max window sizes.
    If uMsg = WM_GETMINMAXINFO Then
        'Retrieve default MinMax settings
        CopyMemoryToMinMaxInfo MinMax, lParam, Len(MinMax)

       ' set the width of the form when it's maximized
'        MinMax.ptMaxSize.x = MDIForm1.Width + 100
        ' set the height of the form when it's maximized
'        If colHWnd.Item(1).hWnd <> hw Then
'            MinMax.ptMaxSize.Y = mdi_npls.Height + 200
'        End If
        ' set the Left of the form when it's maximized
'        MinMax.ptMaxPosition.x = 0
        ' set the Top of the form when it's maximized
'        If colHWnd.Item(1).hWnd <> hw Then
'            MinMax.ptMaxPosition.Y = -30
'        End If
        
        'Specify new minimum size for window.
        MinMax.ptMinTrackSize.X = colHWnd.Item(CStr(hw)).Width
        MinMax.ptMinTrackSize.Y = colHWnd.Item(CStr(hw)).Height

        'Specify new maximum size for window.
        MinMax.ptMaxTrackSize.X = 100000000
        MinMax.ptMaxTrackSize.Y = 100000000

        'Copy local structure back.
        CopyMemoryFromMinMaxInfo lParam, MinMax, Len(MinMax)

        WindowProc = DefWindowProc(hw, uMsg, wParam, lParam) '
    Else
        WindowProc = CallWindowProc(colHWnd.Item(CStr(hw)).PrevWndProc, hw, uMsg, _
           wParam, lParam)
    End If
End Function
