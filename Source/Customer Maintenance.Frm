VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_cust_maint 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Maintenance"
   ClientHeight    =   1950
   ClientLeft      =   2820
   ClientTop       =   3360
   ClientWidth     =   5505
   ClipControls    =   0   'False
   Icon            =   "Customer Maintenance.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1950
   ScaleWidth      =   5505
   Begin VB.Frame fraCustomer 
      Height          =   795
      Left            =   120
      TabIndex        =   6
      Top             =   1020
      Width           =   5235
      Begin VB.TextBox txt_customer 
         Height          =   375
         Left            =   1560
         TabIndex        =   8
         Top             =   240
         Width           =   3555
      End
      Begin VB.Label lab_cust_name_lab 
         Caption         =   "&Customer Name:"
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   300
         Width           =   1275
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   675
      Left            =   4620
      Picture         =   "Customer Maintenance.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Label..."
      Height          =   675
      Left            =   3720
      Picture         =   "Customer Maintenance.frx":0B04
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "&View"
      Height          =   675
      Left            =   2820
      Picture         =   "Customer Maintenance.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   675
      Left            =   1920
      Picture         =   "Customer Maintenance.frx":1388
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   675
      Left            =   1020
      Picture         =   "Customer Maintenance.frx":1A4A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   675
      Left            =   120
      Picture         =   "Customer Maintenance.frx":210C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin Crystal.CrystalReport crs_delivery_note 
      Left            =   480
      Top             =   780
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowBorderStyle=   1
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin MSComDlg.CommonDialog cdg_print 
      Left            =   60
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Set printer for Customer Label"
   End
End
Attribute VB_Name = "frm_cust_maint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public snap_customer As DAO.Recordset
Public cancel_list As Boolean
Private mobjCustomer As Customer2

Public Sub load_and_delete_customer(delete_customer_name As String)
'***************************************
' Module/Form Name   : frm_cust_maint
'
' Procedure Name     : load_and_delete_customer
'
' Purpose            :
'
' Date Created       : 02/06/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : delete_customer_name - String
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 02/06/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo load_and_delete_customer_Error
'
'******** Code Starts Here *************
'
    Set mobjCustomer = New Customer2
    On Error Resume Next
    mobjCustomer.Locate delete_customer_name
    If Err.Number = vbObjectError + 7 Then  'Cancelled
        Exit Sub
    ElseIf Err.Number <> 0 Then
        ErrorSave
        On Error GoTo load_and_delete_customer_Error
        ErrorRestore
    Else
        On Error GoTo load_and_delete_customer_Error
    End If
    With mobjCustomer
        .Delete
    End With
'
'********* Code Ends Here **************
'
Exit Sub
    '
load_and_delete_customer_Error:
    ErrorRaise "frm_cust_maint.load_and_delete_customer"
End Sub

Public Sub tab_select(ByRef tab_customer As Object, selected_item As Integer, ByRef picture_box As Object, ByRef mod_pic_moved)
    '***************************************
    ' Module/Form Name   : frm_cust_maint
    '
    ' Procedure Name     : tab_select
    '
    ' Purpose            :
    '
    ' Date Created       : 02/06/2002
    '
    ' Author             : GARETH SAUNDERS
    '
    ' Parameters         : tab_customer - Object
    '                    : selected_item - Integer
    '                    : picture_box - Object
    '                    : mod_pic_moved -
    '
    ' Amendment History  : Date       Author    Description
    '                    : --------------------------------
    '                    : 02/06/2002 GARETH SAUNDERS
    '
    '***************************************
    '
On Error GoTo tab_select_Error
    '
    '******** Code Starts Here *************
    '
    '
    '   Used for positioning pictureboxes on the
    '   Customer screens
    '
If mod_pic_moved(selected_item) = "N" Then
  With picture_box(selected_item - 1)
      .Height = tab_customer.ClientHeight
      .Width = tab_customer.ClientWidth
      .Left = tab_customer.ClientLeft
      .Top = tab_customer.ClientTop
      .Visible = True
      mod_pic_moved(selected_item) = "Y"
  End With
End If
picture_box(selected_item - 1).ZOrder 0
    '
    '********* Code Ends Here **************
    '
Exit Sub
    '
tab_select_Error:
    ErrorRaise "frm_cust_maint.tab_select"
End Sub

Private Sub cmdDelete_Click()
'***************************************
' Module/Form Name   : frm_cust_maint
'
' Procedure Name     : cmdDelete_Click
'
' Purpose            :
'
' Date Created       : 02/06/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 02/06/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo cmdDelete_Click_Error
'
'******** Code Starts Here *************
'
    If Len(RTrim(txt_customer.Text)) = 0 Then
        MsgBox "Please enter at least part of a customer's name", vbExclamation, "Customer Maintenance"
        txt_customer.SetFocus
        Exit Sub
    End If
    
    txt_customer.SetFocus
    On Error Resume Next
    load_and_delete_customer txt_customer.Text
    Select Case Err.Number
        Case 0, vbObjectError + 7
            On Error GoTo cmdDelete_Click_Error
        Case vbObjectError + 6      '   No customer found
            MsgBox "No Customer found for deletion", vbExclamation, App.Title
        Case Else
            ErrorSave
            On Error GoTo cmdDelete_Click_Error
            ErrorRestore
    End Select
'
'********* Code Ends Here **************
'
    Exit Sub
'
cmdDelete_Click_Error:
    DisplayError , "frm_cust_maint.cmdDelete_Click", vbExclamation
End Sub

Private Sub cmdEdit_Click()
    '***************************************
    ' Module/Form Name   : frm_cust_maint
    '
    ' Procedure Name     : cmdEdit_Click
    '
    ' Purpose            :
    '
    ' Date Created       : 31/07/2001
    '
    ' Author             : GARETH
    '
    ' Amendment History  : Date       Author    Description
    '                    : --------------------------------
    '                    : 02/06/2002 GARETH SAUNDERS
    '
    '***************************************
    '
On Error GoTo cmdEdit_Click_Error
    '
    '******** Code Starts Here *************
    '

If Len(RTrim(txt_customer.Text)) = 0 Then
  MsgBox "Please enter at least part of a customer's name", vbExclamation, "Customer Maintenance"
  txt_customer.SetFocus
  Exit Sub
End If

txt_customer.SetFocus
With frmCustEdit
  .mMode = Edit
  .mstrCustomerName = txt_customer.Text
  On Error Resume Next
  Load frmCustEdit
  If is_form_loaded("frmcustedit") Then
      .SetFocus
  End If
End With
    '
    '********* Code Ends Here **************
    '
Exit Sub
    '
cmdEdit_Click_Error:
    DisplayError , "frm_cust_maint.com_edit_Click", vbExclamation
End Sub

Private Sub cmdExit_Click()
    Unload frm_cust_maint
    Set frm_cust_maint = Nothing
End Sub


Private Sub cmdNew_Click()
    '***************************************
    ' Module/Form Name   : frm_cust_maint
    '
    ' Procedure Name     : cmdNew_Click
    '
    ' Purpose            :
    '
    ' Date Created       : 02/06/2002
    '
    ' Author             : GARETH SAUNDERS
    '
    ' Amendment History  : Date       Author    Description
    '                    : --------------------------------
    '                    : 02/06/2002 GARETH SAUNDERS
    '
    '***************************************
    '
On Error GoTo cmdNew_Click_Error
    '
    '******** Code Starts Here *************
    '

txt_customer.SetFocus
With frm_cust_new
    '        .mMode = Add
  Load frm_cust_new
  .SetFocus
End With

    '
    '********* Code Ends Here **************
    '
Exit Sub
    '
cmdNew_Click_Error:
    DisplayError , "frm_cust_maint.cmdNew_Click", vbExclamation
End Sub

Private Sub cmdView_Click()
'***************************************
' Module/Form Name   : frm_cust_maint
'
' Procedure Name     : cmdView_Click
'
' Purpose            :
'
' Date Created       : 02/06/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 02/06/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo cmdView_Click_Error
'
'******** Code Starts Here *************
'
    If Len(RTrim(txt_customer.Text)) = 0 Then
        MsgBox "Please enter at least part of a customer's name", vbExclamation, "Customer Maintenance"
        txt_customer.SetFocus
        Exit Sub
    End If
    
    txt_customer.SetFocus
    On Error Resume Next
    frm_cust_view.load_customer txt_customer.Text
    Select Case Err.Number - vbObjectError
        Case Is = 6     '   No customer found
            MsgBox "No Customer found for viewing", vbExclamation, App.Title
            On Error GoTo cmdView_Click_Error
        Case Is = -vbObjectError, 7    '   Cancelled
            On Error GoTo cmdView_Click_Error
        Case Else
            ErrorSave
            On Error GoTo cmdView_Click_Error
            ErrorRestore
    End Select
    
    If is_form_loaded("frm_cust_view") Then
        frm_cust_view.SetFocus
    End If
'
'********* Code Ends Here **************
'
    Exit Sub
    '
cmdView_Click_Error:
    DisplayError , "frm_cust_maint.cmdView_Click", vbExclamation
End Sub

Private Sub Form_Load()
    'txt_customer.Text = ""
    txt_customer.Text = GetSetting(App.Title, "CustomerMaintenance", "Customer", "")
    HighLightText txt_customer
    Me.Show
    com_position_form Me
    txt_customer.SetFocus
    DoEvents
End Sub

Private Sub cmdPrint_Click()
    If Len(RTrim(txt_customer.Text)) = 0 Then
        MsgBox "Please enter at least part of a customer's name", vbExclamation, "Customer Maintenance"
        txt_customer.SetFocus
        Exit Sub
    End If
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting App.Title, "CustomerMaintenance", "Customer", txt_customer.Text
End Sub
