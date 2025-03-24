VERSION 5.00
Begin VB.Form frmNewKeywordExclusion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Keyword Exclusion"
   ClientHeight    =   1230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5250
   Icon            =   "NewKeywordExclusion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1230
   ScaleWidth      =   5250
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtKeywordExcluded 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   555
      Left            =   3480
      Picture         =   "NewKeywordExclusion.frx":0742
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      UseMaskColor    =   -1  'True
      Width           =   795
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   555
      Left            =   4380
      Picture         =   "NewKeywordExclusion.frx":0CBC
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   600
      UseMaskColor    =   -1  'True
      Width           =   795
   End
   Begin VB.Label lblKeywordExcluded 
      Caption         =   "&Excluded Keyword:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmNewKeywordExclusion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnCancel              As Boolean
Private mvarKeywordExcluded     As String
Private moKeywordExclusions     As KeywordExclusions

Public Property Get KeywordExcluded() As String
    KeywordExcluded = mvarKeywordExcluded
End Property

Public Property Get Cancel() As Boolean
    Cancel = mblnCancel
End Property

Private Sub cmdCancel_Click()
    mblnCancel = True
    Unload Me
End Sub

Private Sub cmdOK_Click()
'***************************************
' Module/Form Name   : frmNewKeywordExclusion
'
' Procedure Name     : cmdOK_Click
'
' Purpose            :
'
' Date Created       : 19/07/2005
'
' Author             : Gareth Saunders
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                      19/07/2005 Gareth Saunders
'
'***************************************
'
    On Error GoTo ErrorHandler:
'
'******** Code Starts Here *************
'
    On Error Resume Next
    moKeywordExclusions.Validate txtKeywordExcluded.Text
    If Err.Number > vbObjectError And Err.Number < vbObjectError + 100 Then
        MsgBox Err.Description, vbExclamation
        txtKeywordExcluded.SetFocus
        Exit Sub
    End If
    '
    On Error GoTo ErrorHandler:
    mvarKeywordExcluded = txtKeywordExcluded.Text
    Unload Me
    DoEvents
'
'********* Code Ends Here **************
'
    On Error GoTo 0
    Exit Sub
'
ErrorHandler:
    DisplayError , "frmNewKeywordExclusion.cmdOK_Click", vbExclamation
End Sub

Private Sub Form_Initialize()
    mblnCancel = False
End Sub

Private Sub Form_Load()
    cmdOK.Default = True
End Sub

Public Sub Display(ByVal poKeywordExclusions As KeywordExclusions)
    Set moKeywordExclusions = poKeywordExclusions
    Me.Show vbModal
End Sub
