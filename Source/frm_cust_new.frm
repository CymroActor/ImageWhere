VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frm_cust_new 
   AutoRedraw      =   -1  'True
   Caption         =   "New Customer"
   ClientHeight    =   4980
   ClientLeft      =   2520
   ClientTop       =   2235
   ClientWidth     =   8295
   ClipControls    =   0   'False
   Icon            =   "frm_cust_new.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4980
   ScaleWidth      =   8295
   Begin VB.CommandButton ssc_done 
      Caption         =   "&OK"
      Height          =   555
      Left            =   1800
      Picture         =   "frm_cust_new.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4400
      UseMaskColor    =   -1  'True
      Width           =   795
   End
   Begin VB.CommandButton ssc_cancel 
      Caption         =   "&Cancel"
      Height          =   555
      Left            =   4680
      Picture         =   "frm_cust_new.frx":09BC
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4400
      UseMaskColor    =   -1  'True
      Width           =   795
   End
   Begin TabDlg.SSTab sstCustomer 
      Height          =   4215
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   7435
      _Version        =   393216
      Style           =   1
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "&General"
      TabPicture(0)   =   "frm_cust_new.frx":0F4E
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txt_county_or_state"
      Tab(0).Control(1)=   "txt_country"
      Tab(0).Control(2)=   "txt_fax_no"
      Tab(0).Control(3)=   "txt_telephone_no"
      Tab(0).Control(4)=   "txt_post_code"
      Tab(0).Control(5)=   "txt_address_line_3"
      Tab(0).Control(6)=   "txt_address_line_2"
      Tab(0).Control(7)=   "txt_address_line_1"
      Tab(0).Control(8)=   "txt_customer_name"
      Tab(0).Control(9)=   "lab_county_or_state_lab(1)"
      Tab(0).Control(10)=   "lab_country_lab(0)"
      Tab(0).Control(11)=   "lab_address_lab"
      Tab(0).Control(12)=   "lab_customername_lab"
      Tab(0).Control(13)=   "lab_telephoneno_lab(4)"
      Tab(0).Control(14)=   "lab_faxno_lab(5)"
      Tab(0).Control(15)=   "lab_post_code_lab(9)"
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "&Delivery Information"
      TabPicture(1)   =   "frm_cust_new.frx":0F6A
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lab_slash"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lab_preferred_delivery_method_lab"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lab_business_type_lab"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lab_information_lab"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lab_vat_no_country_code_lab"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txt_information"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txt_country_code"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txt_vat_no"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txt_preferred_delivery_method"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "cboBusinessType"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "&Billing"
      TabPicture(2)   =   "frm_cust_new.frx":0F86
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cboSageURN"
      Tab(2).Control(1)=   "lblSageURN"
      Tab(2).Control(2)=   "lblTermsAndConditionsApproved"
      Tab(2).Control(3)=   "lblOnHold"
      Tab(2).Control(4)=   "lblSageTermsAgreed"
      Tab(2).Control(5)=   "lblSageOnHold"
      Tab(2).ControlCount=   6
      Begin VB.ComboBox cboBusinessType 
         Height          =   315
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   3720
         Width           =   3435
      End
      Begin VB.ComboBox cboSageURN 
         Height          =   315
         Left            =   -72600
         TabIndex        =   29
         Text            =   "Combo1"
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txt_preferred_delivery_method 
         DataField       =   "preferred_delivery_method"
         DataSource      =   "dat_customer"
         Height          =   285
         Left            =   2640
         MaxLength       =   30
         TabIndex        =   22
         Top             =   3300
         Width           =   2595
      End
      Begin VB.TextBox txt_vat_no 
         DataField       =   "vat_no"
         DataSource      =   "dat_customer"
         Height          =   285
         Left            =   3240
         MaxLength       =   15
         TabIndex        =   21
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox txt_country_code 
         DataField       =   "country_code"
         DataSource      =   "dat_customer"
         Height          =   285
         Left            =   2640
         MaxLength       =   3
         TabIndex        =   20
         Top             =   1080
         Width           =   435
      End
      Begin VB.TextBox txt_information 
         DataField       =   "information"
         DataSource      =   "dat_customer"
         Height          =   495
         Left            =   2640
         MaxLength       =   300
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   480
         Width           =   5235
      End
      Begin VB.TextBox txt_county_or_state 
         DataField       =   "country"
         DataSource      =   "dat_customer"
         Height          =   285
         Left            =   -72780
         MaxLength       =   20
         TabIndex        =   7
         Top             =   2040
         Width           =   2295
      End
      Begin VB.TextBox txt_country 
         DataField       =   "country"
         DataSource      =   "dat_customer"
         Height          =   285
         Left            =   -72780
         MaxLength       =   20
         TabIndex        =   9
         Top             =   2400
         Width           =   2295
      End
      Begin VB.TextBox txt_fax_no 
         DataField       =   "fax_no"
         DataSource      =   "dat_customer"
         Height          =   285
         Left            =   -72780
         MaxLength       =   20
         TabIndex        =   15
         Top             =   3600
         Width           =   1935
      End
      Begin VB.TextBox txt_telephone_no 
         DataField       =   "telephone_no"
         DataSource      =   "dat_customer"
         Height          =   285
         Left            =   -72780
         MaxLength       =   20
         TabIndex        =   13
         Top             =   3240
         Width           =   1935
      End
      Begin VB.TextBox txt_post_code 
         DataField       =   "post_code"
         DataSource      =   "dat_customer"
         Height          =   285
         Left            =   -72780
         MaxLength       =   12
         TabIndex        =   11
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox txt_address_line_3 
         DataField       =   "address_line_3"
         DataSource      =   "dat_customer"
         Height          =   285
         Left            =   -72780
         MaxLength       =   40
         TabIndex        =   5
         Top             =   1680
         Width           =   4095
      End
      Begin VB.TextBox txt_address_line_2 
         DataField       =   "address_line_2"
         DataSource      =   "dat_customer"
         Height          =   285
         Left            =   -72780
         MaxLength       =   40
         TabIndex        =   4
         Top             =   1320
         Width           =   4095
      End
      Begin VB.TextBox txt_address_line_1 
         DataField       =   "address_line_1"
         DataSource      =   "dat_customer"
         Height          =   285
         Left            =   -72780
         MaxLength       =   40
         TabIndex        =   3
         Top             =   960
         Width           =   4095
      End
      Begin VB.TextBox txt_customer_name 
         DataField       =   "customer_name"
         DataSource      =   "dat_customer"
         Height          =   285
         Left            =   -72780
         MaxLength       =   40
         TabIndex        =   1
         Top             =   600
         Width           =   4095
      End
      Begin VB.Label lblSageURN 
         Caption         =   "Sage URN:"
         Height          =   255
         Left            =   -74400
         TabIndex        =   34
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblTermsAndConditionsApproved 
         Caption         =   "Terms Agreed:"
         Height          =   255
         Left            =   -74400
         TabIndex        =   33
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblOnHold 
         Caption         =   "On Hold:"
         Height          =   255
         Left            =   -74400
         TabIndex        =   32
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label lblSageTermsAgreed 
         Height          =   255
         Left            =   -72600
         TabIndex        =   31
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label lblSageOnHold 
         Height          =   255
         Left            =   -72600
         TabIndex        =   30
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label lab_vat_no_country_code_lab 
         Caption         =   "Country Code/VAT No.:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   1140
         Width           =   1755
      End
      Begin VB.Label lab_information_lab 
         Caption         =   "Information:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   480
         Width           =   1635
      End
      Begin VB.Label lab_business_type_lab 
         Caption         =   "Business Type:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   3780
         Width           =   1815
      End
      Begin VB.Label lab_preferred_delivery_method_lab 
         Caption         =   "Preferred Delivery Method:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   3360
         Width           =   2115
      End
      Begin VB.Label lab_slash 
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3120
         TabIndex        =   24
         Top             =   1080
         Width           =   135
      End
      Begin VB.Label lab_county_or_state_lab 
         Caption         =   "Count&y/State:"
         Height          =   255
         Index           =   1
         Left            =   -74760
         TabIndex        =   6
         Top             =   2100
         Width           =   1815
      End
      Begin VB.Label lab_country_lab 
         Caption         =   "Cou&ntry:"
         Height          =   255
         Index           =   0
         Left            =   -74760
         TabIndex        =   8
         Top             =   2460
         Width           =   1815
      End
      Begin VB.Label lab_address_lab 
         Caption         =   "&Address:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   -74760
         TabIndex        =   2
         Top             =   1020
         Width           =   1395
      End
      Begin VB.Label lab_customername_lab 
         Caption         =   "C&ustomer Name:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   -74760
         TabIndex        =   0
         Top             =   660
         Width           =   1395
      End
      Begin VB.Label lab_telephoneno_lab 
         Caption         =   "&Telephone No:"
         Height          =   255
         Index           =   4
         Left            =   -74760
         TabIndex        =   12
         Top             =   3240
         Width           =   1815
      End
      Begin VB.Label lab_faxno_lab 
         Caption         =   "Fa&x No:"
         Height          =   255
         Index           =   5
         Left            =   -74760
         TabIndex        =   14
         Top             =   3600
         Width           =   555
      End
      Begin VB.Label lab_post_code_lab 
         Caption         =   "&Post  Code:"
         Height          =   255
         Index           =   9
         Left            =   -74760
         TabIndex        =   10
         Top             =   2880
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frm_cust_new"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Public snap_customer As DAO.Recordset
Private dyna_customer As DAO.Recordset

Private retry_update As Boolean
Private SQL As String
Dim mod_pic_moved(3) As String
Public cancel_list As Boolean
Private current_tab As Integer
Private mbolDontResize As Boolean
Public mstrCurrentContact As String
Public ContactsDisplayed As Boolean
Private Const FORM_MIN_HEIGHT = 5385
Private Const FORM_MIN_WIDTH = 8415
'Public mMode As UpdateMode
Public mstrCustomerName As String
Private moCustomer As Customer2
Private moSageAccounts As SageAccounts

Private Function valid_input() As Boolean
      '***************************************
      ' Module/Form Name   : frm_cust_new
      '
      ' Procedure Name     : valid_input
      '
      ' Purpose            :
      '
      ' Date Created       : 18/05/2002
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
10    On Error GoTo valid_input_Error
      '
      '******** Code Starts Here *************
      '
20        valid_input = False

30        If Trim(txt_customer_name.Text) = "" Then
40            MsgBox "Please enter the Customer's Name", vbExclamation
50            sstCustomer.Tab = 0
60            txt_customer_name.SetFocus
70            Exit Function
80        End If
    
90        If cboBusinessType.ListIndex = -1 Then
100           MsgBox "Please select a Business Type", vbExclamation
110           cboBusinessType.SetFocus
120           sstCustomer.Tab = 1
130           Exit Function
140       End If
    
150       valid_input = True
      '
      '********* Code Ends Here **************
      '
160      Exit Function
      '
valid_input_Error:
170       ErrorRaise "frm_cust_new.valid_input"
End Function

Private Sub cboSageURN_Click()
      '***************************************
      ' Module/Form Name   : frm_cust_new
      '
      ' Procedure Name     : cboSageURN_Click
      '
      ' Purpose            :
      '
      ' Date Created       : 06/12/2002
      '
      ' Author             : GARETH SAUNDERS
      '
      ' Amendment History  : Date       Author    Description
      '                    : --------------------------------
      '
      '***************************************
      '
10    On Error GoTo cboSageURN_Click_Error
      '
      '******** Code Starts Here *************
      '
    
20        If cboSageURN.ListIndex = 0 Then
30            moCustomer.SageURN = ""
40        Else
50            moCustomer.SageURN = cboSageURN.Text
60        End If
70        RefreshSageFields
      '
      '********* Code Ends Here **************
      '
80       Exit Sub
      '
cboSageURN_Click_Error:
90        DisplayError , "frm_cust_new.cboSageURN_Click", vbExclamation
End Sub

Private Sub Form_Load()
      '***************************************
      ' Module/Form Name   : frm_cust_new
      '
      ' Procedure Name     : Form_Load
      '
      ' Purpose            :
      '
      ' Date Created       : 18/05/2002
      '
      ' Author             : GARETH SAUNDERS
      '
      ' Amendment History  : Date       Author    Description
      '                    : --------------------------------
      '
      '***************************************
      '
10    On Error GoTo Form_Load_Error
      '
      '******** Code Starts Here *************
      '
          Dim oBusinessTypes As BusinessTypes
          Dim oBusinessType As BusinessType
    
20        On Error GoTo ErrorProc
          '
          '   Create Customer Object
          '
30        Set moCustomer = New Customer2
          '
          '   If there is no Sage Link don't show the Billing Tab.
          '
40        If goSystemConfig.SageLink Then
50            sstCustomer.TabVisible(2) = True
60            InitialiseSageURNs
70        Else
80            sstCustomer.TabVisible(2) = False
90        End If
          '
          '   Display the Business Types.
          '
100       Set oBusinessTypes = New BusinessTypes
110       For Each oBusinessType In oBusinessTypes
120           cboBusinessType.AddItem oBusinessType.Name
130       Next oBusinessType
          '
          '   Set the width & height of the form.
          '
140       mbolDontResize = True
150       Me.Caption = "New Customer"
    
160       sstCustomer.Tab = 0
170       Me.Width = FORM_MIN_WIDTH
180       Me.Height = FORM_MIN_HEIGHT
190       com_position_form Me
200       DoEvents
210       mbolDontResize = False
          '
220       ContactsDisplayed = False
          '
230       Me.Show
240       txt_customer_name.SetFocus
250       Exit Sub
ErrorProc:
260       DisplayError
      '
      '********* Code Ends Here **************
      '
270      Exit Sub
      '
Form_Load_Error:
280       DisplayError , "frm_cust_new.Form_Load", vbExclamation
End Sub

Private Sub Form_Resize()
    Dim Ctl As Control, CtlCln As New Collection
    
'
'   Ignore the first resize as this is caused through setting the
'   initial dimensions of the form.
'
    If mbolDontResize Then
        Exit Sub
    End If
    
    If Not ResizeForm(Me) Then
        Exit Sub
    End If
    
    If WindowState = 1 Then
        Exit Sub
    End If
  
    If WindowState = 0 Then
        If Width < FORM_MIN_WIDTH Then
            Width = FORM_MIN_WIDTH
        End If
        If Height < FORM_MIN_HEIGHT Then
            Height = FORM_MIN_HEIGHT
        End If
    End If
    
    sstCustomer.Visible = False
    '
    '   Determine those controls with bad left properties!
    '
    On Error Resume Next
    For Each Ctl In Controls
       If Ctl.Left < 0 Then CtlCln.Add Ctl
    Next
    '
    '   Now add 75000 to all negative left properties.
    '
    For Each Ctl In Controls
       If Ctl.Left < 0 Then Ctl.Left = Ctl.Left + 75000
    Next
'
'   Reposition Edit & Cancel buttons.
'
    ssc_done.Top = Me.Height - ssc_done.Height - 600
    ssc_cancel.Top = ssc_done.Top
    ssc_done.Left = (Width - (ssc_done.Width + ssc_cancel.Width + 500)) / 2
    ssc_cancel.Left = ssc_done.Left + ssc_done.Width + 500
'
'   Resize tab.
'
    sstCustomer.Width = Width - sstCustomer.Left * 2 - 200
    sstCustomer.Height = ssc_done.Top - sstCustomer.Top - 100
'
'   Resize all other fields.
'
    txt_information.Left = txt_country_code.Left
    txt_information.Width = sstCustomer.Width - txt_information.Left - 100
    lab_preferred_delivery_method_lab.Top = lab_vat_no_country_code_lab.Top + 400
    txt_preferred_delivery_method.Top = lab_preferred_delivery_method_lab.Top
    lab_business_type_lab.Top = lab_preferred_delivery_method_lab.Top + 400
    cboBusinessType.Top = lab_business_type_lab.Top
    '
    '   Put bad left property controls back again.
    '
    On Error Resume Next
    For Each Ctl In CtlCln
       If Ctl.Left > 0 Then Ctl.Left = Ctl.Left - 75000
    Next
    While CtlCln.Count <> 0
        CtlCln.Remove 1
    Wend
    Set CtlCln = Nothing
    sstCustomer.Visible = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set moCustomer = Nothing
    On Error Resume Next
    gcolMaxedWindows.Remove CStr(Me.hWnd)
End Sub

Private Sub ssc_cancel_Click()
10        Unload Me
End Sub

Private Sub ssc_done_Click()
      '***************************************
      ' Module/Form Name   : frm_cust_new
      '
      ' Procedure Name     : ssc_done_Click
      '
      ' Purpose            :
      '
      ' Date Created       : 06/12/2002
      '
      ' Author             : GARETH SAUNDERS
      '
      ' Amendment History  : Date       Author    Description
      '                    : --------------------------------
      '
      '***************************************
      '
10    On Error GoTo ssc_done_Click_Error
      '
      '******** Code Starts Here *************
      '
          Dim objCustomer As Customer2
    
20        If Not valid_input Then
30            Exit Sub
40        End If
    
50        With moCustomer
60            .CustomerName = txt_customer_name.Text
70            .Address1 = txt_address_line_1
80            .Address2 = txt_address_line_2
90            .Address3 = txt_address_line_3
100           .BusinessType = cboBusinessType.Text
110           .Country = txt_country.Text
120           .CountryCode = txt_country_code.Text
130           .CountyOrState = txt_county_or_state.Text
140           .Fax = txt_fax_no.Text
150           .Information = txt_information.Text
160           .PostCode = txt_post_code.Text
170           .PreferredDelivery = txt_preferred_delivery_method.Text
180           .Telephone = txt_telephone_no.Text
190           .VATNo = txt_vat_no.Text
200           On Error Resume Next
210           .create
220           Select Case Err.Number - vbObjectError
                  Case Is = 1
230                   Beep
240                   MsgBox "There is already a customer with this name and" + vbCr + _
                             "first line of address.", vbExclamation
250                   Exit Sub
260               Case Is = 2
270                   msg = Err.Description & vbCr _
                            & "Do you wish to continue?"
280                   style = vbYesNo + vbInformation + vbDefaultButton2
  
290                   response = MsgBox(msg, style)

300                   If response = vbNo Then
310                       Exit Sub
320                   End If
330                   On Error GoTo ssc_done_Click_Error
340                   .create Force:=True
350               Case Else
360                   If Err.Number <> 0 Then
370                       ErrorSave
380                       On Error GoTo ssc_done_Click_Error
390                       ErrorRestore
400                   Else
410                       On Error GoTo ssc_done_Click_Error
420                   End If
430           End Select
440       End With
    
450       If is_form_loaded("frm_cust_maint") Then
460           With frm_cust_maint
470               .txt_customer = moCustomer.CustomerName
480               HighLightText .txt_customer
490           End With
500       End If
    
510       Unload Me
520       Set frm_cust_new = Nothing
530       Set moCustomer = Nothing
      '
      '********* Code Ends Here **************
      '
540      Exit Sub
      '
ssc_done_Click_Error:
550       DisplayError , "frm_cust_new.ssc_done_Click", vbExclamation
End Sub

Private Sub RefreshSageFields()
      '***************************************
      ' Module/Form Name   : frm_cust_new
      '
      ' Procedure Name     : RefreshSageFields
      '
      ' Purpose            :
      '
      ' Date Created       : 06/12/2002
      '
      ' Author             : GARETH SAUNDERS
      '
      ' Amendment History  : Date       Author    Description
      '                    : --------------------------------
      '
      '***************************************
      '
10    On Error GoTo RefreshSageFields_Error
      '
      '******** Code Starts Here *************
      '
20        With moCustomer
30            lblSageTermsAgreed = IIf(.TermsAndConditionsApproved, "Yes", "No")
40            lblSageOnHold = IIf(.OnHold, "Yes", "No")
50        End With
      '
      '********* Code Ends Here **************
      '
60       Exit Sub
      '
RefreshSageFields_Error:
70        ErrorRaise "frm_cust_new.RefreshSageFields"
End Sub

Private Sub InitialiseSageURNs()
      '***************************************
      ' Module/Form Name   : frm_cust_new
      '
      ' Procedure Name     : InitialiseSageURNs
      '
      ' Purpose            :
      '
      ' Date Created       : 06/12/2002
      '
      ' Author             : GARETH SAUNDERS
      '
      ' Amendment History  : Date       Author    Description
      '                    : --------------------------------
      '
      '***************************************
      '
10    On Error GoTo InitialiseSageURNs_Error
      '
      '******** Code Starts Here *************
      '
          Dim oSageAccount As SageAccount
  
20        cboSageURN.AddItem "<Not Set>"
30        Set moSageAccounts = New SageAccounts
40        For Each oSageAccount In moSageAccounts
50            cboSageURN.AddItem oSageAccount.Key
60        Next oSageAccount
70        cboSageURN.ListIndex = 0
80        lblSageTermsAgreed = "No"
90        lblSageOnHold = "No"
    
      '
      '********* Code Ends Here **************
      '
100      Exit Sub
      '
InitialiseSageURNs_Error:
110       ErrorRaise "frm_cust_new.InitialiseSageURNs"
End Sub

Public Sub ForceResize()
    gblnResizeMaxedWindows = False
    Form_Resize
    gblnResizeMaxedWindows = True
End Sub


