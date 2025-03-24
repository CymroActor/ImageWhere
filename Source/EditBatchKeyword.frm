VERSION 5.00
Begin VB.Form frmEditBatchKeyword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5835
   Icon            =   "EditBatchKeyword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   5835
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   555
      Left            =   4080
      Picture         =   "EditBatchKeyword.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      UseMaskColor    =   -1  'True
      Width           =   795
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   555
      Left            =   4980
      Picture         =   "EditBatchKeyword.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      UseMaskColor    =   -1  'True
      Width           =   795
   End
   Begin VB.TextBox txtKeyword 
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   4575
   End
   Begin VB.Label lblKeyword 
      Caption         =   "Keyword:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmEditBatchKeyword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrKeyword     As String
Private mblnCancel      As Boolean
Private moBatch         As Batch2
Private mAction         As UpdateMode

Public Property Get Keyword() As String
    Keyword = mstrKeyword
End Property

Public Property Get Cancel() As Boolean
    Cancel = mblnCancel
End Property

Public Sub Display(ByVal poBatch As Batch2, _
                   ByVal Action As UpdateMode, _
                   Optional ByVal pstrKeyword As String)
    
    mstrKeyword = pstrKeyword
    Set moBatch = poBatch
    mAction = Action
    
    If Action = Add Then
        Me.Caption = "Add Keyword"
        txtKeyword.Text = ""
    Else
        Me.Caption = "Edit Keyword"
        txtKeyword.Text = mstrKeyword
    End If
    '
    Me.Show vbModal
End Sub

Private Sub cmdCancel_Click()
    mblnCancel = True
    Unload Me
End Sub

Private Sub cmdOK_Click()
'***************************************
' Module/Form Name   : frmEditBatchKeyword
'
' Procedure Name     : cmdOK_Click
'
' Purpose            :
'
' Date Created       : 18/07/2005
'
' Author             : Gareth Saunders
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                      18/07/2005 Gareth Saunders
'
'***************************************
'
    On Error GoTo ErrorHandler:
'
'******** Code Starts Here *************
'
    If mAction = Edit Then
        If LCase(mstrKeyword) = LCase(txtKeyword.Text) Then
            MsgBox "No changes have been made", vbExclamation
            txtKeyword.SetFocus
            Exit Sub
        End If
    End If
    '
    On Error Resume Next
    moBatch.Keywords.Validate txtKeyword.Text, mAction
    If Err.Number > vbObjectError And Err.Number < vbObjectError + 100 Then
        MsgBox Err.Description, vbExclamation
        txtKeyword.SetFocus
        Exit Sub
    End If
    '
    On Error GoTo ErrorHandler:
    mstrKeyword = txtKeyword.Text
    Unload Me
'
'********* Code Ends Here **************
'
    On Error GoTo 0
    Exit Sub
'
ErrorHandler:
    DisplayError , "frmEditBatchKeyword.cmdOK_Click", vbExclamation
End Sub

Private Sub Form_Initialize()
    mblnCancel = False
End Sub

Private Sub txtKeyword_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
