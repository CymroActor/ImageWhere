VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ResizeClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mvarhWnd As Long
Private mvarWidth As Long
Private mvarHeight As Long
Private mvarWidthPixels As Long
Private mvarHeightPixels As Long
Private mvarPrevWndProc As Long

Public Property Get hWnd() As Long
    hWnd = mvarhWnd
End Property

Public Property Let hWnd(vData As Long)
    mvarhWnd = vData
End Property

Public Property Get Width() As Long
    Width = mvarWidth
End Property

Public Property Let Width(vData As Long)
    mvarWidth = vData
End Property

Public Property Get Height() As Long
    Height = mvarHeight
End Property

Public Property Let Height(vData As Long)
    mvarHeight = vData
End Property

Public Property Get WidthPixels() As Long
    WidthPixels = mvarWidthPixels
End Property

Public Property Let WidthPixels(vData As Long)
    mvarWidthPixels = vData
End Property

Public Property Get HeightPixels() As Long
    HeightPixels = mvarHeightPixels
End Property

Public Property Let HeightPixels(vData As Long)
    mvarHeightPixels = vData
End Property

Public Property Get PrevWndProc() As Long
    PrevWndProc = mvarPrevWndProc
End Property

Public Property Let PrevWndProc(vData As Long)
    mvarPrevWndProc = vData
End Property


