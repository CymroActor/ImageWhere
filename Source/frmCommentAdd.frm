VERSION 5.00
Begin VB.Form frmCommentAdd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Comment"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmCommentAdd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtAuthor 
      Height          =   375
      Left            =   2280
      MaxLength       =   20
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
   Begin VB.TextBox txtComment 
      Height          =   2175
      Left            =   240
      MaxLength       =   200
      MultiLine       =   -1  'True
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   2  'Automatic
      TabIndex        =   2
      Top             =   840
      Width           =   4335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   555
      Left            =   3720
      Picture         =   "frmCommentAdd.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3240
      UseMaskColor    =   -1  'True
      Width           =   795
   End
   Begin VB.CommandButton cmbOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   555
      Left            =   2400
      Picture         =   "frmCommentAdd.frx":059E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3240
      UseMaskColor    =   -1  'True
      Width           =   795
   End
   Begin VB.Image imgComment 
      Height          =   480
      Left            =   240
      Picture         =   "frmCommentAdd.frx":0B18
      Top             =   120
      Width           =   480
   End
   Begin VB.Label labAuthor 
      Caption         =   "&Author:"
      Height          =   255
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "frmCommentAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event CommentsUpdated(ID As String)
Public mobjCustomer As Customer2
Public mbolEditComment As Boolean
Public mstrID As String

Private Sub cmbOK_Click()
'***************************************
' Module/Form Name   : frmCommentAdd
'
' Procedure Name     : cmbOK_Click
'
' Purpose            :
'
' Date Created       : 31/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo cmbOK_Click_Error
'
'******** Code Starts Here *************
'

    If Not ValidInput Then
        Exit Sub
    End If

    If mbolEditComment Then
        With mobjCustomer.Comments(mstrID)
            .Comment = txtComment.Text
            .DateWritten = Now
            On Error Resume Next
            .update
            If Err.Number = vbObjectError + 5 Then
                DisplayError
                txtComment.Text = mobjCustomer.Comments(mstrID).Comment
                txtAuthor.Text = mobjCustomer.Comments(mstrID).Author
                Exit Sub
            Else
                If Err.Number <> 0 Then
                    ErrorSave
                    On Error GoTo cmbOK_Click_Error
                    ErrorRestore
                Else
                    On Error GoTo cmbOK_Click_Error
                End If
            End If
        End With
    Else
        mstrID = CStr(mobjCustomer.Comments.Create(txtAuthor, _
                                                   txtComment, _
                                                   Now, _
                                                   False))
    End If

    RaiseEvent CommentsUpdated(mstrID)
    SaveSetting App.Title, "Add Comment", "Author", txtAuthor.Text
    Unload Me
'
'********* Code Ends Here **************
'
   Exit Sub
'
cmbOK_Click_Error:
    DisplayError , "frmCommentAdd.cmbOK_Click", vbExclamation
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
'***************************************
' Module/Form Name   : frmCommentAdd
'
' Procedure Name     : Form_Load
'
' Purpose            :
'
' Date Created       : 31/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo Form_Load_Error
'
'******** Code Starts Here *************
'
    If mbolEditComment Then
        Me.Caption = "Edit Comment"
        txtComment.Text = mobjCustomer.Comments(mstrID).Comment
        txtAuthor.Text = mobjCustomer.Comments(mstrID).Author
        txtAuthor.Enabled = False
        txtAuthor.BackColor = vbButtonFace
    Else
        Me.Caption = "Add Comment"
        txtAuthor.Text = GetSetting(App.Title, "Add Comment", "Author")
    End If

'
'********* Code Ends Here **************
'
   Exit Sub
'
Form_Load_Error:
    DisplayError , "frmCommentAdd.Form_Load", vbExclamation
End Sub

Private Sub Form_Paint()
    If Len(Trim(txtAuthor.Text)) <> 0 Then
        txtComment.SetFocus
    End If
End Sub

Private Sub txtAuthor_Change()
    If Len(Trim(txtAuthor.Text)) = 0 Then
        txtComment.Enabled = False
        txtComment.BackColor = vbButtonFace
    Else
        txtComment.Enabled = True
        txtComment.BackColor = vbWindowBackground
    End If
End Sub

Private Function ValidInput() As Boolean
'***************************************
' Module/Form Name   : frmCommentAdd
'
' Procedure Name     : ValidInput
'
' Purpose            :
'
' Date Created       : 31/12/2002
'
' Author             : GARETH SAUNDERS
'
' Returns            : Boolean
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo ValidInput_Error
'
'******** Code Starts Here *************
'
    ValidInput = False

    If Len(Trim(txtAuthor)) = 0 Then
        MsgBox "Please enter your Name", vbExclamation
        txtAuthor.SetFocus
        Exit Function
    End If

    If Len(Trim(txtComment)) = 0 Then
        MsgBox "Please enter a Comment", vbExclamation
        txtComment.SetFocus
        Exit Function
    End If

    ValidInput = True

'
'********* Code Ends Here **************
'
   Exit Function
'
ValidInput_Error:
    ErrorRaise "frmCommentAdd.ValidInput"
End Function
