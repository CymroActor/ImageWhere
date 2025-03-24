VERSION 5.00
Begin VB.Form frmJobFilter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Job Filter"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7560
   Icon            =   "JobFilter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   7560
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraCustomer 
      Caption         =   "Customer"
      ForeColor       =   &H80000001&
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7335
      Begin VB.CommandButton cmdAllCustomers 
         Caption         =   "&All Customers"
         Height          =   435
         Left            =   5640
         TabIndex        =   6
         Top             =   720
         Width           =   1515
      End
      Begin VB.TextBox txt_address_line_1 
         BackColor       =   &H8000000F&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1380
         TabIndex        =   5
         Top             =   600
         Width           =   4095
      End
      Begin VB.TextBox txt_customer_name 
         BackColor       =   &H8000000F&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   4095
      End
      Begin VB.CommandButton cmdSelectCustomer 
         Caption         =   "Select C&ustomer..."
         Height          =   435
         Left            =   5640
         TabIndex        =   3
         Top             =   240
         Width           =   1515
      End
      Begin VB.Label lab_address_line_1_lab 
         Caption         =   "A&ddress Line 1:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1155
      End
      Begin VB.Label lab_customer_name_lab 
         Caption         =   "&Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   1215
      End
   End
   Begin VB.ComboBox cboFilter 
      Height          =   315
      Left            =   1560
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   1560
      Width           =   2295
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   555
      Left            =   6600
      Picture         =   "JobFilter.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1560
      UseMaskColor    =   -1  'True
      Width           =   795
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   555
      Left            =   5640
      Picture         =   "JobFilter.frx":09D4
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1560
      UseMaskColor    =   -1  'True
      Width           =   795
   End
   Begin VB.Label lblStatus 
      Caption         =   "&Status:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   975
   End
End
Attribute VB_Name = "frmJobFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mvarCustomer        As Customer2
Private mvarShowAll         As Boolean
Private mvarShowOpenOnly    As Boolean
Private mblnCancel          As Boolean

Public WithEvents fInputCustomer As frm_input_customer
Attribute fInputCustomer.VB_VarHelpID = -1
Public Event JobFiltered()

Public Property Get Customer() As Customer2
    Set Customer = mvarCustomer
End Property

Public Property Set Customer(vData As Customer2)
    Set mvarCustomer = vData
End Property

Public Property Get ShowAll() As Boolean
    ShowAll = mvarShowAll
End Property

Public Property Let ShowAll(vData As Boolean)
    mvarShowAll = vData
End Property

Public Property Get ShowOpenOnly() As Boolean
    ShowOpenOnly = mvarShowOpenOnly
End Property

Public Property Let ShowOpenOnly(vData As Boolean)
    mvarShowOpenOnly = vData
End Property

Public Property Get Cancel() As Boolean
    Cancel = mblnCancel
End Property

Public Property Let Cancel(vData As Boolean)
    mblnCancel = vData
End Property

Private Sub cboFilter_Click()
    If cboFilter.ListIndex = 0 Then
        mvarShowAll = True
        mvarShowOpenOnly = False
    Else
        mvarShowAll = False
        mvarShowOpenOnly = True
    End If
End Sub

Private Sub cmdAllCustomers_Click()
    txt_customer_name.Text = ""
    txt_address_line_1.Text = ""
    Set mvarCustomer = Nothing
    Set mvarCustomer = New Customer2
End Sub

Private Sub cmdCancel_Click()
    mblnCancel = True
    Unload Me
End Sub

Public Sub Display()
    txt_customer_name.Text = mvarCustomer.CustomerName
    txt_address_line_1.Text = mvarCustomer.Address1
    With cboFilter
        .AddItem "Show All"
        .AddItem "Show Open Only"
        If mvarShowAll Then
            .ListIndex = 0
        Else
            .ListIndex = 1
        End If
    End With
     Me.Show vbModal
End Sub

Private Sub cmdOK_Click()
    
'***************************************
' Module/Form Name   : frmJobFilter
'
' Procedure Name     : cmdOK_Click
'
' Purpose            :
'
' Date Created       : 19/04/2006 11:06
'
' Author             : Gareth
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
    On Error GoTo ErrorProc

    If Not ValidEntry Then Exit Sub

    mvarShowAll = (cboFilter.ListIndex = 0)
    mvarShowOpenOnly = (cboFilter.ListIndex = 1)
    RaiseEvent JobFiltered
    Unload Me
    DoEvents

    On Error GoTo 0
    Exit Sub
'
'********* Code Ends Here **************
'
ErrorProc:
    DisplayError , "frmJobFilter.cmdOK_Click", vbExclamation
End Sub

Private Function ValidEntry() As Boolean
'***************************************
' Module/Form Name   : frmJobFilter
'
' Procedure Name     : ValidEntry
'
' Purpose            :
'
' Date Created       : 19/04/2006 11:08
'
' Author             : Gareth
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
    On Error GoTo ErrorProc

    ValidEntry = False
    
    With frmJobFilter
        If .ShowAll = True And mvarCustomer.CustomerNo = 0 Then
            MsgBox "A customer must be selected when displaying all jobs", vbExclamation
            txt_customer_name.SetFocus
            Exit Function
        End If
    End With
    
    ValidEntry = True

    On Error GoTo 0
    Exit Function
'
'********* Code Ends Here **************
'
ErrorProc:
    ErrorRaise "frmJobFilter.ValidEntry"
End Function

Private Sub cmdSelectCustomer_Click()
    Set fInputCustomer = New frm_input_customer
    With fInputCustomer
        .Caption = "Select Customer"
        .Show 1
    End With
  
    If Me.Visible = True Then
        cboFilter.SetFocus
    End If
    Set fInputCustomer = Nothing
End Sub

Private Sub fInputCustomer_CustomerFound(oCustomer As Customer2, blnAbort As Boolean)
    If Not blnAbort Then
        txt_customer_name.Text = oCustomer.CustomerName
        txt_address_line_1.Text = oCustomer.Address1
        Set mvarCustomer = oCustomer
    End If
End Sub

Private Sub Form_Load()
    mblnCancel = False
End Sub
