Attribute VB_Name = "mod_ToolTip"
Option Explicit

' Location coordinate type.

  Type POINTAPI
    X As Long
    Y As Long
  End Type
  
  Declare Function GetActiveWindow Lib "user32" () As Long
  Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
  Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
  Declare Function WindowFromPoint Lib "user32" (ByVal lpPointX As Long, ByVal lpPointY As Long) As Long

' ShowWindow() Command
Const SW_SHOWNOACTIVATE = 4

' Define the intervals used for the timer for both when the label is visible and not visible.
' The time unit is milliseconds.
Public Const Time_When_ToolTip_Displayed = 5
Public Const Time_When_ToolTip_Not_Displayed = 1500

Public current_control_details As Control
