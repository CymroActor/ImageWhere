VERSION 5.00
Begin VB.Form frm_input_customer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Customer"
   ClientHeight    =   1860
   ClientLeft      =   2760
   ClientTop       =   2580
   ClientWidth     =   5790
   ControlBox      =   0   'False
   Icon            =   "Input Customer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1860
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   Begin VB.Timer timCust 
      Interval        =   1
      Left            =   720
      Top             =   1320
   End
   Begin VB.Frame fraCustomer 
      Height          =   795
      Left            =   60
      TabIndex        =   3
      Top             =   240
      Width           =   5655
      Begin VB.TextBox txt_customer 
         Height          =   375
         Left            =   1680
         TabIndex        =   0
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label lab_cust_name_lab 
         Caption         =   "&Customer Name:"
         Height          =   315
         Left            =   360
         TabIndex        =   4
         Top             =   300
         Width           =   1275
      End
   End
   Begin VB.CommandButton cmdSelect 
      Default         =   -1  'True
      Height          =   600
      Left            =   4320
      Picture         =   "Input Customer.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   600
   End
   Begin VB.CommandButton cmdCancel 
      Height          =   600
      Left            =   5040
      Picture         =   "Input Customer.frx":0ED6
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   600
   End
End
Attribute VB_Name = "frm_input_customer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private retry As Boolean
Private mod_input_form As Form

Public customer_text As String
Public Cancel As Boolean
Public Event CustomerFound(oCustomer As Customer2, blnAbort As Boolean)

Public Sub select_customer(input_form As Form)
10        Set mod_input_form = input_form
20        Me.Icon = input_form.Icon
30        txt_customer.Text = input_form.mstrCustomerSelect
40        Me.Show 1
End Sub

Private Sub Form_Load()
10        txt_customer.Text = GetSetting(App.Title, "frm_input_customer", "Customer", "")
End Sub

Private Sub Form_Unload(Cancel As Integer)
10        SaveSetting App.Title, "frm_input_customer", "Customer", txt_customer.Text
End Sub

Private Sub cmdCancel_Click()
10        Cancel = True
20        Unload Me
      '    Set frm_input_customer = Nothing

End Sub

Private Sub cmdSelect_Click()
      '***************************************
      ' Module/Form Name   : frm_input_customer
      '
      ' Procedure Name     : cmdSelect_Click
      '
      ' Purpose            :
      '
      ' Date Created       : 07/12/2002
      '
      ' Author             : GARETH SAUNDERS
      '
      ' Amendment History  : Date       Author    Description
      '                    : --------------------------------
      '
      '***************************************
      '
10    On Error GoTo cmdSelect_Click_Error
      '
      '******** Code Starts Here *************
      '
          Dim oCustomer As Customer2
          Dim blnAbort As Boolean
    
20        If txt_customer.Text = "" Then
30            MsgBox "Enter a customer name or text that is contained within it", vbInformation
40            txt_customer.SetFocus
50            Exit Sub
60        End If
    
70        Screen.MousePointer = vbHourglass
    
80        Set oCustomer = New Customer2
90        On Error Resume Next
100       oCustomer.Locate txt_customer.Text
110       txt_customer.SetFocus
120       Select Case Err.Number
              Case Is = vbObjectError + 7     'Cancelled
130               txt_customer.SetFocus
140               Screen.MousePointer = vbDefault
150               Exit Sub
160           Case Is = vbObjectError + 6     'Customer not found
170               MsgBox "No Customer found for selection", vbExclamation, "Image Where?"
180               Screen.MousePointer = vbDefault
190               Exit Sub
200           Case Is <> 0
210               ErrorSave
220               On Error GoTo cmdSelect_Click_Error
230               ErrorRestore
240           Case Else
250               On Error GoTo cmdSelect_Click_Error
260       End Select
270       blnAbort = False
280       DoEvents
290       RaiseEvent CustomerFound(oCustomer, blnAbort)
    
300       Screen.MousePointer = vbDefault
    
310       If Not blnAbort Then
320           Unload Me
330           DoEvents
340       End If
      '
      '********* Code Ends Here **************
      '
350      Exit Sub
      '
cmdSelect_Click_Error:
360       DisplayError , "frm_input_customer.cmdSelect_Click", vbExclamation
End Sub

Private Sub timCust_Timer()
          '
          '   Select the text in the customer textbox once the form is shown modally.
          '
10        timCust.Enabled = False
20        txt_customer.SelStart = 0
30        txt_customer.SelLength = Len(txt_customer)
End Sub
