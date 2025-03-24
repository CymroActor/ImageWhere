VERSION 5.00
Begin VB.Form frm_cust_list 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Customer List"
   ClientHeight    =   3930
   ClientLeft      =   1875
   ClientTop       =   1920
   ClientWidth     =   6165
   ControlBox      =   0   'False
   Icon            =   "Customer List.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3930
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   Begin ImageWhere.SimpleGrid smgCustomers 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   5318
      Columns         =   1
      KeyCol          =   0
   End
   Begin VB.CommandButton cmdCancel 
      Height          =   600
      Left            =   4620
      Picture         =   "Customer List.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3300
      Width           =   600
   End
   Begin VB.CommandButton cmdSelect 
      Default         =   -1  'True
      Height          =   600
      Left            =   780
      Picture         =   "Customer List.frx":130C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3300
      Width           =   600
   End
End
Attribute VB_Name = "frm_cust_list"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mintCustomerNo As Integer
Private mblnCancel As Boolean

Public Property Get CustomerNo() As Integer
    CustomerNo = mintCustomerNo
End Property
Public Property Get Cancel() As Boolean
    Cancel = mblnCancel
End Property

''Public Sub ProcessCustomer(rstCustomers As ADODB.Recordset)
Public Sub ProcessCustomer(rstCustomers As DAO.Recordset)
'***************************************
' Module/Form Name   : frm_cust_list
'
' Procedure Name     : ProcessCustomer
'
' Purpose            :
'
' Date Created       : 02/06/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : rstCustomers - DAO.Recordset
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 02/06/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo ProcessCustomer_Error
'
'******** Code Starts Here *************
'
    Load Me
    Screen.MousePointer = vbHourglass
    With smgCustomers
        .Redraw = False
        .Clear
        rstCustomers.MoveFirst
        While Not rstCustomers.EOF
            .AddRow False, _
                    rstCustomers!customer_name, _
                    rstCustomers!address_line_1, _
                    rstCustomers!customer_no
            rstCustomers.MoveNext
        Wend
        .ResizeRows
        .Column(1).Sorted = smgAscending
        .CurrentRow = 1
        .Redraw = True
    End With
    Screen.MousePointer = vbDefault
    Me.Show 1
'
'********* Code Ends Here **************
'
    Exit Sub
    '
ProcessCustomer_Error:
    ErrorRaise "frm_cust_list.ProcessCustomer"
End Sub

Private Sub cmdCancel_Click()
    mblnCancel = True
    Unload Me
End Sub

Private Sub cmdSelect_Click()
    mintCustomerNo = smgCustomers.Column(smgCustomers.KeyCol).Value
    Unload Me
End Sub

Private Sub Form_Load()
    '***************************************
    ' Module/Form Name   : frm_cust_list
    '
    ' Procedure Name     : Form_Load
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
On Error GoTo Form_Load_Error
    '
    '******** Code Starts Here *************
    '
mblnCancel = False
With smgCustomers
  .Columns = 3
  .Column(1).Header = "Name"
  .Column(1).Align = flexAlignLeftTop
  .Column(1).Width = smgCustomers.Width / 2 - 182
  .Column(2).Header = "First Line of Address"
  .Column(2).Align = flexAlignLeftTop
  .Column(2).Width = .Column(1).Width
  .Column(3).Width = 0
  .KeyCol = 3
End With
    '
    '********* Code Ends Here **************
    '
Exit Sub
    '
Form_Load_Error:
    DisplayError , "frm_cust_list.Form_Load", vbExclamation
End Sub

Private Sub smgCustomers_DblClick()
    mintCustomerNo = smgCustomers.Column(smgCustomers.KeyCol).Value
    Unload Me
    DoEvents
End Sub



