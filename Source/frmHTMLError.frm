VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmHTMLError 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Error Display"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5745
   Icon            =   "frmHTMLError.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   5745
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timError 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   600
      Top             =   3480
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   495
      Left            =   4440
      TabIndex        =   1
      Top             =   3480
      Width           =   1215
   End
   Begin SHDocVwCtl.WebBrowser wbrError 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      ExtentX         =   9763
      ExtentY         =   5741
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "frmHTMLError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrHTML As String
Private mstrErrorText As String

Public Property Get ErrorText() As String
    ErrorText = mstrErrorText
End Property

Private Sub cmdClose_Click()
    Unload Me
End Sub

Public Sub Display(strHTML As String)
'***************************************
' Module/Form Name   : frmHTMLError
'
' Procedure Name     : Display
'
' Purpose            :
'
' Date Created       : 26/10/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : strHTML - String
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 26/10/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo Display_Error
'
'******** Code Starts Here *************
'
    timError.Enabled = True
    mstrHTML = strHTML
    Me.Show vbModal
    '
    '********* Code Ends Here **************
    '
    Exit Sub
    '
Display_Error:
    ErrorRaise "frmHTMLError.Display"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrErrorText = wbrError.Document.Body.innertext
End Sub

Private Sub timError_Timer()
'***************************************
' Module/Form Name   : frmHTMLError
'
' Procedure Name     : timError_Timer
'
' Purpose            :
'
' Date Created       : 26/10/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 26/10/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo timError_Timer_Error
'
'******** Code Starts Here *************
'
    timError.Enabled = False
    With wbrError
        .Navigate2 "about:blank"
        Do While .readyState <> READYSTATE_COMPLETE
          DoEvents
        Loop
        .Document.Body.Innerhtml = mstrHTML
    End With
    Screen.MousePointer = vbDefault
'
'********* Code Ends Here **************
'
    Exit Sub
    '
timError_Timer_Error:
    DisplayError , "frmHTMLError.timError_Timer", vbExclamation
End Sub
