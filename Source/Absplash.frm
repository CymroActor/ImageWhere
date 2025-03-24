VERSION 5.00
Begin VB.Form AbSplash 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "application title"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6360
   Icon            =   "Absplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox txtWarning 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   13
      Text            =   "Absplash.frx":000C
      Top             =   3720
      Width           =   4695
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   480
      Top             =   480
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      Height          =   375
      Left            =   4920
      TabIndex        =   12
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4920
      TabIndex        =   11
      Top             =   3840
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Height          =   540
      Left            =   120
      ScaleHeight     =   480
      ScaleWidth      =   6075
      TabIndex        =   7
      Top             =   2640
      Width           =   6135
      Begin VB.Label lblUserName 
         BackStyle       =   0  'Transparent
         Caption         =   "user name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   0
         Width           =   5895
      End
      Begin VB.Label lblUserInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "user information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   5895
      End
   End
   Begin VB.Image Image1 
      Height          =   4935
      Left            =   0
      Top             =   0
      Width           =   6375
   End
   Begin VB.Label lblCompanyName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   840
      Width           =   4815
   End
   Begin VB.Label lblMisc 
      BackStyle       =   0  'Transparent
      Caption         =   "This product is licensed to:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   6135
   End
   Begin VB.Label lblPathEXE 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "path and exe information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3240
      Width           =   6135
   End
   Begin VB.Line linDivide 
      Index           =   1
      X1              =   120
      X2              =   6240
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "version information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   1080
      Width           =   4815
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "application title"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
   Begin VB.Image imgIcon 
      Height          =   1215
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblTrademark 
      BackStyle       =   0  'Transparent
      Caption         =   "trademark information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   6135
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      Caption         =   "copyright information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   6135
   End
   Begin VB.Line linDivide 
      Index           =   0
      X1              =   120
      X2              =   6240
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label lblFileDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "file description"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   6135
   End
End
Attribute VB_Name = "AbSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ------------------------------------------------------------------------
'      Copyright © 1997 Microsoft Corporation.  All rights reserved.
'
' You have a royalty-free right to use, modify, reproduce and distribute
' the Sample Application Files (and/or any modified version) in any way
' you find useful, provided that you agree that Microsoft has no warranty,
' obligations or liability for any Sample Application Files.
' ------------------------------------------------------------------------

Option Explicit

' API declarations
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
        (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
        (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" _
        (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" _
        (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

' API Constants
Private Const GWL_STYLE         As Long = (-16)
Private Const WS_CAPTION        As Long = &HC00000
Private Const WS_CAPTION_NOT    As Long = &HFFFFFFFF - WS_CAPTION

Private Const gREGKEYSYSINFOLOC As String = "SOFTWARE\Microsoft\Shared Tools Location"
Private Const gREGKEYSYSINFO    As String = "SOFTWARE\Microsoft\Shared Tools\MSINFO"

Private Const gREGVALSYSINFOLOC As String = "MSINFO"
Private Const gREGVALSYSINFO    As String = "PATH"

' NT location of user name and company
Private Const gNTREGKEYINFO     As String = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
Private Const gNTREGVALUSER     As String = "RegisteredOwner"
Private Const gNTREGVALCOMPANY  As String = "RegisteredOrganization"

' Win95 locataion of user name and company
Private Const g95REGKEYINFO     As String = "Software\Microsoft\MS Setup (ACME)\User Info"
Private Const g95REGVALUSER     As String = "DefName"
Private Const g95REGVALCOMPANY  As String = "DefCompany"

' Change these to what you want the default name and user info to be
Private Const DEFAULT_USER_NAME As String = "USER INFORMATION NOT AVAILABLE"
Private Const DEFAULT_USER_INFO As String = vbNullString

' Information for warning information at bottom of form
Private Const gWarningInfo      As String = "Warning:This computer program is protected by copyright law. " _
                                          & "Unauthorized reproduction or distribution of this program, or any " _
                                          & "portion of it, may result in severe civil and criminal penalties, and " _
                                          & "will be prosecuted to the maximum extent possible under law."

Private mBoxHeight              As Integer
Private mStyle                  As StyleType
Private mTitleBarHidden         As Boolean

' Type declarations
Private Type StyleType
    OldStyle As Long
    NewStyle As Long
End Type 'StyleType

Private Sub Form_Click()
10        Timer1.Enabled = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
10        Timer1.Enabled = False
End Sub

Private Sub Form_Load()
10        txtWarning.Text = gWarningInfo

      '   Fill in all of the information that comes from the App object
20        With App
30            Caption = "About " & .Title
40            lblTitle.Caption = .Title
  
50            If .CompanyName <> "" Then
60                lblCompanyName.Caption = "A product of " & .CompanyName
70            Else
80                lblCompanyName.Caption = ""
90            End If
  
100           lblVersion.Caption = "Version " & .Major & "." & .Minor & "." & _
                                   .Revision
110           lblCopyright.Caption = .LegalCopyright
120           lblTrademark.Caption = .LegalTrademarks
130           lblPathEXE.Caption = .Path & "\" & .EXEName & " "
140           lblFileDescription.Caption = .FileDescription
150       End With 'App
    
      '   Get "default" height of About Box
160       mBoxHeight = Height
End Sub

Private Sub cmdOK_Click()
10        Hide ' If you want to unload the form, change this to Unload Me
End Sub

Private Sub cmdSysInfo_Click()
10        Call StartSysInfo
End Sub

Public Sub About(frmParent As Form, Optional lUserName As String, _
                 Optional lUserInfo As String)
10        imgIcon.Picture = frmParent.Icon
20        cmdOK.Enabled = True
30        cmdSysInfo.Enabled = True
    
      '   Add user information to form
40        If lUserName <> "" Then
50            lblUserName.Caption = lUserName
60            lblUserInfo.Caption = lUserInfo
70        Else
80            lblUserName.Caption = GetUserName
90            lblUserInfo.Caption = GetUserCompany
100       End If
    
      '   Modify the form style to show the title bar
110       ShowTitleBar
    
      '   A resize event is needed in order to apply the changes to the form style.  Setting
      '   the height to 0 should do it.
120       If Height = mBoxHeight Then
130           Height = 0
140       End If
    
      '   Set height of About Box to "default" height
150       Height = mBoxHeight
    
160       Show vbModal, frmParent
End Sub

Public Sub SplashOn(frmParent As Form, Optional MinDisplay As Long, _
                    Optional lUserName As String, Optional lUserInfo As String)
10        If Not Visible Then
              Dim lHeight As Integer
  
20            imgIcon.Picture = frmParent.Icon
30            cmdOK.Enabled = False
40            cmdSysInfo.Enabled = False
    
      '       If a delay is specified, set up the Timer
50            If MinDisplay > 0 Then
60                Timer1.Interval = MinDisplay
70                Timer1.Enabled = True
80            End If
  
      '       Add user information to form
90            If lUserName <> "" Then
100               lblUserName.Caption = lUserName
110               lblUserInfo.Caption = lUserInfo
120           Else
130               lblUserName.Caption = GetUserName
140               lblUserInfo.Caption = GetUserCompany
150           End If
  
      '       Modify the form style to hide the title bar
160           HideTitleBar
  
      '       Need to cause a form resize in order to get updated ScaleHeight value
170           lHeight = Height
180           Height = 0
190           Height = lHeight
  
      '       Set height to hide the "About Box Only" information
200           Height = linDivide(1).Y1 + (Height - ScaleHeight)
  
      '       Show the form
210           Show vbModeless, frmParent

      '       For some reason, need a Refresh to make sure Splash Screen gets painted
220           Refresh
230       End If
End Sub

Public Sub SplashOff()
10        If Visible Then
      '       Wait until any minimum display time elapses
20            Do While Timer1.Enabled
30                DoEvents
40            Loop
  
50            Hide ' If you want to unload the form, change this to Unload Me

      '       Modify the form style to show the title bar
60            ShowTitleBar
  
      '       Set height of About Box to "default" height
70            Height = mBoxHeight
80        End If
End Sub

Private Sub Image1_Click()
10        Timer1.Enabled = False
End Sub

Private Sub lblUserInfo_Click()
10        Timer1.Enabled = False
End Sub

Private Sub lblUserName_Click()
10        Timer1.Enabled = False
End Sub

Private Sub Picture1_Click()
10        Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
10        Timer1.Enabled = False
End Sub

Private Sub HideTitleBar()
      '   Change the style of the form to not show a title bar
10        If mTitleBarHidden Then Exit Sub
    
20        mTitleBarHidden = True
    
30        With mStyle
40            .OldStyle = GetWindowLong(hWnd, GWL_STYLE)
50            .NewStyle = .OldStyle And WS_CAPTION_NOT
60            SetWindowLong hWnd, GWL_STYLE, .NewStyle
70        End With 'mStyle
End Sub

Private Sub ShowTitleBar()
      '   Change the style of the form to show a title bar
10        If Not mTitleBarHidden Then Exit Sub
20        mTitleBarHidden = False
30        SetWindowLong hWnd, GWL_STYLE, mStyle.OldStyle
End Sub

Private Sub StartSysInfo()
10        On Error GoTo SysInfoErr
  
          Dim rc As Long
          Dim SysInfoPath As String
    
          ' Try To Get System Info Program Path\Name From Registry...
20        If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
          ' Try To Get System Info Program Path Only From Registry...
30        ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
              ' Validate Existence Of Known 32 Bit File Version
40            If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
50                SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
      
              ' Error - File Can Not Be Found...
60            Else
70                GoTo SysInfoErr
80            End If
          ' Error - Registry Entry Can Not Be Found...
90        Else
100           GoTo SysInfoErr
110       End If
    
120       Call Shell(SysInfoPath, vbNormalFocus)
    
130       Exit Sub
SysInfoErr:
140       MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub

Private Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
          Dim i As Long                                           ' Loop Counter
          Dim rc As Long                                          ' Return Code
          Dim hKey As Long                                        ' Handle To An Open Registry Key
          Dim KeyValType As Long                                  ' Data Type Of A Registry Key
          Dim tmpVal As String                                    ' Temporary Storage For A Registry Key Value
          Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
          '------------------------------------------------------------
          ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
          '------------------------------------------------------------
10        rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
20        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
30        tmpVal = String$(1024, 0)                             ' Allocate Variable Space
40        KeyValSize = 1024                                       ' Mark Variable Size
    
          '------------------------------------------------------------
          ' Retrieve Registry Key Value...
          '------------------------------------------------------------
50        rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                               KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                  
60        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
70        If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
80            tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
90        Else                                                    ' WinNT Does NOT Null Terminate String...
100           tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
110       End If
          '------------------------------------------------------------
          ' Determine Key Value Type For Conversion...
          '------------------------------------------------------------
120       Select Case KeyValType                                  ' Search Data Types...
          Case REG_SZ                                             ' String Registry Key Data Type
130           KeyVal = tmpVal                                     ' Copy String Value
140       Case REG_DWORD                                          ' Double Word Registry Key Data Type
150           For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
160               KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
170           Next
180           KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
190       End Select
    
200       GetKeyValue = True                                      ' Return Success
210       rc = RegCloseKey(hKey)                                  ' Close Registry Key
220       Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occurred...
230       KeyVal = ""                                             ' Set Return Val To Empty String
240       GetKeyValue = False                                     ' Return Failure
250       rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

Private Function GetUserName() As String
          Dim KeyVal As String
      
      '   For WindowsNT
10        If (GetKeyValue(HKEY_LOCAL_MACHINE, gNTREGKEYINFO, gNTREGVALUSER, KeyVal)) Then
20            GetUserName = KeyVal
      '   For Windows95
30        ElseIf (GetKeyValue(HKEY_CURRENT_USER, g95REGKEYINFO, g95REGVALUSER, KeyVal)) Then
40            GetUserName = KeyVal
      '   None of the above
50        Else
60            GetUserName = DEFAULT_USER_NAME
70        End If
End Function

Private Function GetUserCompany() As String
          Dim KeyVal As String
    
      '   For WindowsNT
10        If (GetKeyValue(HKEY_LOCAL_MACHINE, gNTREGKEYINFO, gNTREGVALCOMPANY, KeyVal)) Then
20            GetUserCompany = KeyVal
      '   For Windows95
30        ElseIf (GetKeyValue(HKEY_CURRENT_USER, g95REGKEYINFO, g95REGVALCOMPANY, KeyVal)) Then
40            GetUserCompany = KeyVal
      '   None of the above
50        Else
60            GetUserCompany = DEFAULT_USER_INFO
70        End If
End Function
