VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmJobNoteAdd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Job Note"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   Icon            =   "frmJobNoteAdd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtRecordedBy 
      Height          =   375
      Left            =   2280
      MaxLength       =   20
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
   Begin VB.TextBox txtJobNote 
      Height          =   2175
      Left            =   240
      MaxLength       =   300
      MultiLine       =   -1  'True
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   2  'Automatic
      TabIndex        =   3
      Top             =   1200
      Width           =   4335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   555
      Left            =   3720
      Picture         =   "frmJobNoteAdd.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3480
      UseMaskColor    =   -1  'True
      Width           =   795
   End
   Begin VB.CommandButton cmbOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   555
      Left            =   2400
      Picture         =   "frmJobNoteAdd.frx":059E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3480
      UseMaskColor    =   -1  'True
      Width           =   795
   End
   Begin MSComctlLib.ImageCombo cmbContact 
      Height          =   330
      Left            =   2280
      TabIndex        =   2
      Top             =   720
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Locked          =   -1  'True
      ImageList       =   "imlRequestor"
   End
   Begin VB.Image imgComment 
      Height          =   480
      Left            =   240
      Picture         =   "frmJobNoteAdd.frx":0B18
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblRecordedBy 
      Caption         =   "&Recorded By:"
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmJobNoteAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public moJob                As Job2
Public moCustomer           As Customer2
Public mbolEditNote         As Boolean

Public Event NoteAdded(DateAdded As Date)

Public Sub Display(ByRef poJob As Job2, _
                   ByRef poCustomer As Customer2)
    Dim oContact As Contact
    
    Set moJob = poJob
    Set moCustomer = poCustomer
    '
    '   Display Contacts.
    '
    cmbContact.ComboItems.Clear
    For Each oContact In moCustomer.Contacts
        cmbContact.ComboItems.Add , , oContact.Name
    Next oContact
    
    Me.Show vbModal
End Sub

Private Sub cmbOK_Click()
    Dim oJobNote    As New Activity
    
    If Not ValidInput Then
        Exit Sub
    End If

    If mbolEditNote Then
    Else
        With oJobNote
            .ActivityType = "JNOT"
            .CustomerNo = moCustomer.CustomerNo
            .JobNo = moJob.JobNo
            .Status = 0
            .StartDate = Now()
            .EndDate = Now()
            .Description = txtJobNote.Text
            .UserField1 = cmbContact.Text
            .UserField2 = txtRecordedBy.Text
            .Create
        End With
    End If

    RaiseEvent NoteAdded(oJobNote.StartDate)
    SaveSetting App.Title, "Add Job Note", "RecordedBy", txtRecordedBy.Text
    Set oJobNote = Nothing
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

Private Sub Form_Initialize()
    mbolEditNote = False
End Sub

Private Sub Form_Load()
    If mbolEditNote Then
    Else
        Me.Caption = "Add Note"
        txtRecordedBy.Text = GetSetting(App.Title, "Add Job Note", "RecordedBy")
    End If
End Sub

Private Function ValidInput() As Boolean
    ValidInput = False

    If Len(Trim(txtRecordedBy.Text)) = 0 Then
        MsgBox "Please enter your Name", vbExclamation
        txtRecordedBy.SetFocus
        Exit Function
    End If

    If Len(Trim(txtJobNote.Text)) = 0 Then
        MsgBox "Please enter a Note", vbExclamation
        txtJobNote.SetFocus
        Exit Function
    End If

    ValidInput = True
End Function
