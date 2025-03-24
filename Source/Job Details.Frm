VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm_job_details 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Input Job Details"
   ClientHeight    =   6840
   ClientLeft      =   2355
   ClientTop       =   1800
   ClientWidth     =   7455
   Icon            =   "Job Details.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6840
   ScaleWidth      =   7455
   Begin MSComctlLib.ImageList imlRequestor 
      Left            =   840
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Job Details.frx":0442
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraDeliveryAddress 
      Caption         =   "Delivery Address"
      ForeColor       =   &H80000001&
      Height          =   2895
      Left            =   120
      TabIndex        =   24
      Top             =   3840
      Width           =   7335
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   555
         Left            =   6360
         Picture         =   "Job Details.frx":075C
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   2280
         UseMaskColor    =   -1  'True
         Width           =   795
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   555
         Left            =   5460
         Picture         =   "Job Details.frx":0CEE
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   2280
         UseMaskColor    =   -1  'True
         Width           =   795
      End
      Begin VB.TextBox txt_del_county_or_state 
         DataField       =   "country"
         DataSource      =   "dat_customer"
         Height          =   285
         Left            =   1380
         MaxLength       =   20
         TabIndex        =   34
         Top             =   1800
         Width           =   2295
      End
      Begin VB.TextBox txt_del_country 
         DataField       =   "country"
         DataSource      =   "dat_customer"
         Height          =   285
         Left            =   1380
         MaxLength       =   20
         TabIndex        =   36
         Top             =   2160
         Width           =   2295
      End
      Begin VB.TextBox txt_del_post_code 
         DataField       =   "post_code"
         DataSource      =   "dat_customer"
         Height          =   285
         Left            =   1380
         MaxLength       =   12
         TabIndex        =   38
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox txt_del_address_line_3 
         DataField       =   "address_line_3"
         DataSource      =   "dat_customer"
         Height          =   285
         Left            =   1380
         MaxLength       =   40
         TabIndex        =   32
         Top             =   1440
         Width           =   4095
      End
      Begin VB.TextBox txt_del_address_line_2 
         DataField       =   "address_line_2"
         DataSource      =   "dat_customer"
         Height          =   285
         Left            =   1380
         MaxLength       =   40
         TabIndex        =   30
         Top             =   1080
         Width           =   4095
      End
      Begin VB.TextBox txt_del_address_line_1 
         DataField       =   "address_line_1"
         DataSource      =   "dat_customer"
         Height          =   285
         Left            =   1380
         MaxLength       =   40
         TabIndex        =   29
         Top             =   720
         Width           =   4095
      End
      Begin VB.CommandButton com_select_delivery 
         Caption         =   "SELECT &DELIVERY..."
         Height          =   555
         Left            =   5760
         TabIndex        =   27
         Top             =   360
         Width           =   1395
      End
      Begin VB.TextBox txt_del_customer_name 
         DataField       =   "address_line_1"
         DataSource      =   "dat_customer"
         Height          =   285
         Left            =   1380
         MaxLength       =   40
         TabIndex        =   26
         Top             =   360
         Width           =   4095
      End
      Begin VB.CommandButton com_default 
         Caption         =   "DEF&AULT"
         Enabled         =   0   'False
         Height          =   555
         Left            =   5760
         TabIndex        =   31
         Top             =   960
         Width           =   1395
      End
      Begin VB.Label lab_county_or_state_lab 
         Caption         =   "County/State:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   33
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lab_country_lab 
         Caption         =   "Country:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   35
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label lab_address_lab 
         Caption         =   "Address:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lab_post_code_lab 
         Caption         =   "Post  Code:"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   37
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label lab_delivery_name 
         Caption         =   "Delivery Name:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame fraCustomer 
      Caption         =   "Customer"
      ForeColor       =   &H80000001&
      Height          =   975
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      Begin VB.CommandButton com_select_customer 
         Caption         =   "SELECT C&USTOMER ..."
         Height          =   555
         Left            =   5760
         TabIndex        =   3
         Top             =   240
         Width           =   1395
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
      Begin VB.TextBox txt_address_line_1 
         BackColor       =   &H8000000F&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1380
         TabIndex        =   5
         Top             =   600
         Width           =   4095
      End
      Begin VB.Label lab_customer_name_lab 
         Caption         =   "Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label lab_address_line_1_lab 
         Caption         =   "Address Line 1:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1155
      End
   End
   Begin VB.Frame fraRequestorsDetails 
      Caption         =   "Requestor's Details"
      ForeColor       =   &H80000001&
      Height          =   1335
      Left            =   60
      TabIndex        =   16
      Top             =   2400
      Width           =   7335
      Begin VB.ComboBox cboPreferredPhotoType 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   600
         Width           =   2655
      End
      Begin MSComctlLib.ImageCombo cmb_Requested_By 
         Height          =   330
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Locked          =   -1  'True
         ImageList       =   "imlRequestor"
      End
      Begin VB.CheckBox chkWebUser 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   255
         Left            =   2520
         TabIndex        =   19
         Top             =   960
         Width           =   200
      End
      Begin VB.TextBox txtRequestorsEmail 
         DataField       =   "post_code"
         DataSource      =   "dat_customer"
         Height          =   285
         Left            =   4440
         MaxLength       =   50
         TabIndex        =   23
         Top             =   600
         Width           =   2700
      End
      Begin VB.TextBox txt_requestors_telno 
         DataField       =   "post_code"
         DataSource      =   "dat_customer"
         Height          =   285
         Left            =   4440
         MaxLength       =   20
         TabIndex        =   21
         Top             =   240
         Width           =   2700
      End
      Begin VB.Label lblWebUser 
         Caption         =   "Web User:"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblRequestorsEmail 
         Caption         =   "Email:"
         Height          =   195
         Left            =   3180
         TabIndex        =   22
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lab_requestors_telno 
         Caption         =   "Telephone:"
         Height          =   195
         Left            =   3180
         TabIndex        =   20
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSComCtl2.DTPicker txt_date_taken 
      Height          =   285
      Left            =   1440
      TabIndex        =   13
      Top             =   2040
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   63635459
      CurrentDate     =   36844
   End
   Begin VB.TextBox txt_job_reference 
      DataField       =   "post_code"
      DataSource      =   "dat_customer"
      Height          =   285
      Left            =   1440
      MaxLength       =   30
      TabIndex        =   7
      Top             =   1020
      Width           =   2715
   End
   Begin VB.TextBox txt_taken_by 
      DataField       =   "post_code"
      DataSource      =   "dat_customer"
      Height          =   285
      Left            =   4500
      MaxLength       =   20
      TabIndex        =   15
      Top             =   2040
      Width           =   2700
   End
   Begin VB.TextBox txt_request_details 
      Height          =   495
      Left            =   1440
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   1440
      Width           =   5775
   End
   Begin MSComCtl2.DTPicker txt_date_required 
      Height          =   285
      Left            =   5700
      TabIndex        =   8
      Top             =   1035
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   63635459
      CurrentDate     =   36844
   End
   Begin VB.Label lab_date_required 
      Caption         =   "Date Required:"
      Height          =   255
      Left            =   4500
      TabIndex        =   9
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label lab_job_reference 
      Caption         =   "Reference:"
      Height          =   195
      Left            =   180
      TabIndex        =   6
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label lab_taken_by 
      Caption         =   "Taken By:"
      Height          =   195
      Left            =   3240
      TabIndex        =   14
      Top             =   2100
      Width           =   1095
   End
   Begin VB.Label lab_date_taken 
      Caption         =   "Date Taken:"
      Height          =   195
      Left            =   180
      TabIndex        =   12
      Top             =   2100
      Width           =   1095
   End
   Begin VB.Label lab_request_details 
      Caption         =   "Request Details:"
      Height          =   375
      Left            =   180
      TabIndex        =   10
      Top             =   1440
      Width           =   1155
   End
End
Attribute VB_Name = "frm_job_details"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mod_job_form As Form
Private mod_action As String

Private retry As Boolean
Private retry_update As Boolean
Private SQL As String

Public cancel_list As Boolean
Public mstrCustomerSelect As String
Private alt_key_pressed As Boolean
Private moJob As Job2
Private mobjContacts As Contacts
Private moCustomer As Customer2
Private moDeliveryCustomer As Customer2
Public WithEvents fInputCustomer As frm_input_customer
Attribute fInputCustomer.VB_VarHelpID = -1
Public WithEvents fDeliveryCustomer As frm_input_customer
Attribute fDeliveryCustomer.VB_VarHelpID = -1

Private Sub create_job()
'***************************************
' Module/Form Name   : frm_job_details
'
' Procedure Name     : create_job
'
' Purpose            :
'
' Date Created       : 09/04/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo create_job_Error
'
'******** Code Starts Here *************
'
    Dim oContact As Contact
    '
    On Error Resume Next
    Set oContact = moCustomer.Contacts.Item(cmb_Requested_By.SelectedItem.Tag)
    
    Set moJob = New Job2
    
    On Error Resume Next
    moJob.PreferredPhotoType = IIf(cboPreferredPhotoType.ListIndex = 0, "T", "D")
    moJob.create 0, _
                 txt_job_reference.Text, _
                 txt_date_taken.Value, _
                 txt_taken_by.Text, _
                 moCustomer.CustomerNo, _
                 oContact, _
                 txt_requestors_telno.Text, _
                 txt_del_address_line_1.Text, _
                 txt_del_address_line_2.Text, _
                 txt_del_address_line_3.Text, _
                 txt_del_county_or_state.Text, _
                 txt_del_country.Text, _
                 txt_del_post_code.Text, _
                 txt_date_required.Value, _
                 txt_request_details.Text, _
                 txt_del_customer_name.Text, _
                 txtRequestorsEmail.Text
    If Err.Number = vbObjectError + 1 Then
        MsgBox Err.Description, vbExclamation, App.Title
    ElseIf Err.Number <> 0 Then
        ErrorSave
        On Error GoTo create_job_Error
        ErrorRestore
    Else
        On Error GoTo create_job_Error
    End If
'
'********* Code Ends Here **************
'
   Exit Sub
'
create_job_Error:
    ErrorRaise "frm_job_details.create_job"
End Sub

Private Sub load_customer(customer_name As String, customer As Customer2, retry As Boolean)
'***************************************
' Module/Form Name   : frm_job_details
'
' Procedure Name     : load_customer
'
' Purpose            :
'
' Date Created       : 10/04/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : customer_name - String
'                    : customer - Customer2
'                    : retry - Boolean
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 10/04/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo load_customer_Error
'
'******** Code Starts Here *************
'
    retry = False
    Set customer = New Customer2
    On Error Resume Next
    customer.Locate customer_name
    Select Case Err.Number
        Case Is = vbObjectError + 7     'Cancelled
            Exit Sub
        Case Is = vbObjectError + 6     'Customer not found
            retry = True
            Exit Sub
        Case Is <> 0
            ErrorSave
            On Error GoTo load_customer_Error
            ErrorRestore
        Case Else
            On Error GoTo load_customer_Error
    End Select
    Exit Sub
'
'********* Code Ends Here **************
'
   Exit Sub
'
load_customer_Error:
    ErrorRaise "frm_job_details.load_customer"
End Sub


Public Sub display_customer()
'***************************************
' Module/Form Name   : frm_job_details
'
' Procedure Name     : display_customer
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
On Error GoTo display_customer_Error
'
'******** Code Starts Here *************
'
    With moCustomer
        txt_customer_name = .CustomerName
        txt_address_line_1 = .Address1
    End With
    
    set_up_contacts

'
'********* Code Ends Here **************
'
   Exit Sub
'
display_customer_Error:
    ErrorRaise "frm_job_details.display_customer"
End Sub

Private Sub display_job()
'***************************************
' Module/Form Name   : frm_job_details
'
' Procedure Name     : display_job
'
' Purpose            :
'
' Date Created       : 01/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 01/12/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo display_job_Error
    '
    '******** Code Starts Here *************
    '
    Dim counter As Integer
    Dim blnFound    As Boolean
    Dim oContact    As Contact
    Dim strRequestedBy  As String

    set_up_contacts
    
    With moCustomer
        txt_customer_name = .CustomerName
        txt_address_line_1 = .Address1
    End With
    
    With moJob
        txt_job_reference.Text = .reference
        txt_request_details.Text = .RequestDetails
        txt_date_taken.Value = Format(.DateTaken, "dd/mm/yyyy")
        txt_taken_by.Text = .TakenBy
        txt_date_required.Value = Format(.DateRequired, "dd/mm/yyyy")
        '
        '   ID may not be set up for historical records or if the contact no long exists.
        '
        On Error Resume Next
        Set oContact = moJob.Contact
        Select Case Err.Number
        Case Is = 0, vbObjectError + 5, vbObjectError + 9
            On Error GoTo display_job_Error
        Case Else
            ErrorSave
            On Error GoTo display_job_Error
            ErrorRestore
        End Select
        '
        If oContact Is Nothing Then
            strRequestedBy = .RequestedBy
        Else
            strRequestedBy = oContact.Name
            chkWebUser.Value = IIf(oContact.WebUser, vbChecked, vbUnchecked)
        End If
        
        blnFound = False
        For counter = 1 To cmb_Requested_By.ComboItems.Count
            If LCase(strRequestedBy) = LCase(cmb_Requested_By.ComboItems(counter)) Then
                Set cmb_Requested_By.SelectedItem = cmb_Requested_By.ComboItems(counter)
                blnFound = True
            End If
        Next counter
        If Not blnFound Then
            cmb_Requested_By.Text = ""
        End If
        '
        If cmb_Requested_By.ComboItems.Count <> 0 Then
            If mod_action <> "E" And cmb_Requested_By.Text <> "" Then
                SetDefaultRequestorsDetails
            End If
        End If
        txt_requestors_telno = .RequestorsTelNo
        txt_del_customer_name = .DeliveryCustomerName
        txt_del_address_line_1 = .DeliveryAddressLine1
        txt_del_address_line_2 = .DeliveryAddressLine2
        txt_del_address_line_3 = .DeliveryAddressLine3
        txt_del_county_or_state = .DeliveryCountyState
        txt_del_country.Text = .DeliveryCountry
        txt_del_post_code.Text = .DeliveryPostCode
        cboPreferredPhotoType.ListIndex = IIf(moJob.PreferredPhotoType = "T", 0, 1)
        txtRequestorsEmail.Text = .RequestorsEmail
    End With
    
    '
    '********* Code Ends Here **************
    '
    Exit Sub
    '
display_job_Error:
    ErrorRaise "frm_job_details.display_job"
End Sub

Private Sub set_up_contacts()
    '***************************************
    ' Module/Form Name   : frm_job_details
    '
    ' Procedure Name     : set_up_contacts
    '
    ' Purpose            :
    '
    ' Date Created       : 09/12/2001
    '
    ' Author             : GARETH
    '
    ' Amendment History  : Date       Author    Description
    '                    : --------------------------------
    '                    : 09/12/2001 GARETH
    '
    '***************************************
    '
    On Error GoTo set_up_contacts_Error
    '
    '******** Code Starts Here *************
    '
    Dim objContact As Contact
    Dim cboX As ComboItem

    cmb_Requested_By.ComboItems.Clear
    For Each objContact In moCustomer.Contacts
        Set cboX = cmb_Requested_By.ComboItems.Add(, , objContact.Name)
        cboX.Tag = CStr(objContact.ID)
        If objContact.WebUser Then
            cboX.Image = 1
        End If
    Next objContact
    If cmb_Requested_By.Text <> "" Then
        If cmb_Requested_By.ComboItems.Count > 0 Then
            Set cmb_Requested_By.SelectedItem = cmb_Requested_By.ComboItems.Item(1)
            SetDefaultRequestorsDetails
        Else
            txt_requestors_telno.Text = ""
            txtRequestorsEmail = ""
            chkWebUser.Value = vbUnchecked
        End If
    End If
    '
    '********* Code Ends Here **************
    '
    Exit Sub
    '
set_up_contacts_Error:
    ErrorRaise "frm_job_details.set_up_contacts"
End Sub

Private Sub SetDefaultRequestorsDetails()
'***************************************
' Module/Form Name   : frm_job_details
'
' Procedure Name     : SetDefaultRequestorsDetails
'
' Purpose            :
'
' Date Created       : 09/04/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 09/04/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo SetDefaultRequestorsDetails_Error
'
'******** Code Starts Here *************
'
    Dim oContact As Contact
    '
    '
    On Error Resume Next
    Set oContact = moCustomer.Contacts.Item(cmb_Requested_By.SelectedItem.Tag)
    '
    chkWebUser.Value = IIf(oContact.WebUser, vbChecked, vbUnchecked)
    If Trim(txt_requestors_telno.Text) <> "" Or _
       Trim(txtRequestorsEmail.Text) <> "" And _
       (txt_requestors_telno.Text <> oContact.Phone Or txtRequestorsEmail.Text <> oContact.Email) Then
        If MsgBox("Do you wish to overwrite existing contact details for the job?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            Set oContact = Nothing
            Exit Sub
        End If
    End If
    '
    '   Set up the Telephone Number, Email Address & Web User indicator.
    '
    txt_requestors_telno = oContact.Phone
    txtRequestorsEmail.Text = oContact.Email
    cboPreferredPhotoType.ListIndex = IIf(oContact.PreferredPhotoType = "T", 0, 1)
    
    On Error GoTo SetDefaultRequestorsDetails_Error
'
'********* Code Ends Here **************
'
    Exit Sub
    '
SetDefaultRequestorsDetails_Error:
    ErrorRaise "frm_job_details.SetDefaultRequestorsDetails"
End Sub

Private Sub update_job()
    '***************************************
    ' Module/Form Name   : frm_job_details
    '
    ' Procedure Name     : update_job
    '
    ' Purpose            :
    '
    ' Date Created       : 25/05/2002
    '
    ' Author             : GARETH SAUNDERS
    '
    ' Amendment History  : Date       Author    Description
    '                    : --------------------------------
    '                    : 25/05/2002 GARETH SAUNDERS
    '
    '***************************************
    '
On Error GoTo update_job_Error
    '
    '******** Code Starts Here *************
    '
retry_update = False
begin_trans
  With moJob
      '
      '   Update the Job record.
      '
      .reference = txt_job_reference.Text
      .DateTaken = txt_date_taken.Value
      .TakenBy = txt_taken_by.Text
      .CustomerNo = moCustomer.CustomerNo
      .RequestedBy = cmb_Requested_By.Text
      .RequestedByID = moCustomer.Contacts.Item(cmb_Requested_By.SelectedItem.Tag).ID
      .RequestorsTelNo = txt_requestors_telno.Text
      .DeliveryAddressLine1 = txt_del_address_line_1.Text
      .DeliveryAddressLine2 = txt_del_address_line_2.Text
      .DeliveryAddressLine3 = txt_del_address_line_3.Text
      .DeliveryCountyState = txt_del_county_or_state.Text
      .DeliveryCountry = txt_del_country.Text
      .DeliveryPostCode = txt_del_post_code.Text
      .DateRequired = txt_date_required.Value
      .RequestDetails = txt_request_details.Text
      .DeliveryCustomerName = txt_del_customer_name.Text
      .RequestorsEmail = txtRequestorsEmail.Text
      .PreferredPhotoType = IIf(cboPreferredPhotoType.ListIndex = 0, "T", "D")
      .update
  End With
commit_trans

    '
    '********* Code Ends Here **************
    '
Exit Sub
    '
update_job_Error:
ErrorRaise "frm_job_details.update_job"
End Sub

Private Function valid_input() As Boolean
'***************************************
' Module/Form Name   : frm_job_details
'
' Procedure Name     : valid_input
'
' Purpose            :
'
' Date Created       : 01/12/2002
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
On Error GoTo valid_input_Error
'
'******** Code Starts Here *************
'
    valid_input = True
    
    If Len(LTrim(txt_customer_name.Text)) = 0 Then
        MsgBox "Enter a Customer", vbExclamation
        com_select_customer.SetFocus
        valid_input = False
        Exit Function
    End If
    
    If Len(LTrim(txt_job_reference.Text)) = 0 Then
        MsgBox "Enter a Job Reference", vbExclamation
        txt_job_reference.SetFocus
        valid_input = False
        Exit Function
    End If
    
    If Not IsDate(txt_date_required) Then
        MsgBox "Enter a valid date", vbExclamation
        txt_date_required.SetFocus
        valid_input = False
        Exit Function
    End If
    
    If Not IsDate(txt_date_taken) Then
        MsgBox "Enter a valid date", vbExclamation
        txt_date_taken.SetFocus
        valid_input = False
        Exit Function
    End If
    
    If txt_taken_by.Text = "" Then
        MsgBox "Enter who took the request", vbExclamation
        txt_taken_by.SetFocus
        valid_input = False
        Exit Function
    End If
    
    If txt_del_customer_name.Text = "" Then
        MsgBox "The delivery customer name and first line of address must be entered." + vbCr + _
               "Alternatively, select DEFAULT for delivery to main customer.", vbExclamation
        txt_del_customer_name.SetFocus
        valid_input = False
        Exit Function
    End If
    
    If txt_del_address_line_1.Text = "" Then
        MsgBox "Enter at least the first line of the delivery address", vbExclamation
        txt_del_address_line_1.SetFocus
        valid_input = False
        Exit Function
    End If
    
    If cmb_Requested_By.SelectedItem Is Nothing Then
        MsgBox "Please select a contact", vbExclamation
        cmb_Requested_By.SetFocus
        valid_input = False
        Exit Function
    End If
    '
    '   Contact must be a web contact if job is a web job.
    '
    If Not moJob Is Nothing Then
        If moJob.WebStatus <> "" And _
           moCustomer.Contacts.Item(cmb_Requested_By.SelectedItem.Tag).WebUser = False Then
            MsgBox "Contact must be a Web User for this job.", vbExclamation, App.Title
            cmb_Requested_By.SetFocus
            valid_input = False
            Exit Function
        End If
    End If
'
'********* Code Ends Here **************
'
   Exit Function
'
valid_input_Error:
    ErrorRaise "frm_job_details.valid_input"
End Function
Public Sub input_details(ByVal Action As String, _
                         Optional ByVal oJob As Job2)
    '***************************************
    ' Module/Form Name   : frm_job_details
    '
    ' Procedure Name     : input_details
    '
    ' Purpose            :
    '
    ' Date Created       : 30/11/2002
    '
    ' Author             : GARETH SAUNDERS
    '
    ' Parameters         : Action - String
    '                    : oJob - Job2
    '
    ' Amendment History  : Date       Author    Description
    '                    : --------------------------------
    '                    : 01/12/2002 GARETH SAUNDERS
    '
    '***************************************
    '
On Error GoTo input_details_Error
    '
    '******** Code Starts Here *************
    '

mod_action = Action

If mod_action = "E" Then
  '
  '      On edit set up label for key description field
  '      and display current job details
  '
  Set moJob = Nothing
  Set moJob = New Job2
  moJob.Read oJob.JobNo
  '
  Set moCustomer = Nothing
  Set moCustomer = New Customer2
  moCustomer.Read moJob.CustomerNo
  '
  '   Do not allow changing of customer if:
  '   1.  photographs are pending.
  '   2.  job is a web job that has been posted.
  '
  If moJob.NoPhotos = 0 And _
     moJob.WebStatus = "" Then
      com_select_customer.Enabled = True
  Else
      com_select_customer.Enabled = False
  End If

  display_job
Else
    '
    '       On create set up default field values.
    '
  txt_date_taken.Value = Format(Date, "dd/mm/yyyy")
  txt_date_required.Value = Format(Date, "dd/mm/yyyy")
  com_select_customer.Value = True
End If
    If com_select_customer.Enabled Then
        SendKeys "{TAB}{TAB}{TAB}"
    Else
        SendKeys "{TAB}{TAB}"
    End If

    Me.Show
    Me.SetFocus
    DoEvents

    '
    '********* Code Ends Here **************
    '
Exit Sub
    '
input_details_Error:
    ErrorRaise "frm_job_details.input_details"
End Sub

Private Sub cmb_requested_by_Click()
    '***************************************
    ' Module/Form Name   : frm_job_details
    '
    ' Procedure Name     : cmb_requested_by_Click
    '
    ' Purpose            :
    '
    ' Date Created       : 09/04/2002
    '
    ' Author             : GARETH SAUNDERS
    '
    ' Amendment History  : Date       Author    Description
    '                    : --------------------------------
    '                    : 09/04/2002 GARETH SAUNDERS
    '
    '***************************************
    '
On Error GoTo cmb_requested_by_Click_Error
    '
    '******** Code Starts Here *************
    '
    SetDefaultRequestorsDetails
    '
    '********* Code Ends Here **************
    '
Exit Sub
    '
cmb_requested_by_Click_Error:
    DisplayError , "frm_job_details.cmb_requested_by_Click", vbExclamation
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    Set frm_job_details = Nothing
End Sub

Private Sub com_default_Click()
    With moCustomer
        txt_del_customer_name.Text = .CustomerName
        txt_del_address_line_1.Text = .Address1
        txt_del_address_line_2.Text = .Address2
        txt_del_address_line_3.Text = .Address3
        txt_del_county_or_state.Text = .CountyOrState
        txt_del_country.Text = .Country
        txt_del_post_code.Text = .PostCode
        cmdOK.SetFocus
    End With
End Sub

Private Sub com_select_customer_Click()
'***************************************
' Module/Form Name   : frm_job_details
'
' Procedure Name     : com_select_customer_Click
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
On Error GoTo com_select_customer_Click_Error
'
'******** Code Starts Here *************
'
    Set fInputCustomer = New frm_input_customer
    With fInputCustomer
        .Caption = "Select Customer"
        .Show 1
    End With
  
    If Me.Visible = True Then
        txt_job_reference.SetFocus
    End If
    Set fInputCustomer = Nothing
'
'********* Code Ends Here **************
'
   Exit Sub
'
com_select_customer_Click_Error:
    DisplayError , "frm_job_details.com_select_customer_Click", vbExclamation
End Sub

Private Sub com_select_delivery_Click()
'***************************************
' Module/Form Name   : frm_job_details
'
' Procedure Name     : com_select_delivery_Click
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
On Error GoTo com_select_delivery_Click_Error
'
'******** Code Starts Here *************
'
    
    Set fDeliveryCustomer = New frm_input_customer
    With fDeliveryCustomer
        .Caption = "Select Customer for Delivery"
        .Show 1
    End With
    
    If Me.Visible = True Then
        cmdOK.SetFocus
    End If
    
    Set fDeliveryCustomer = Nothing
  
'
'********* Code Ends Here **************
'
   Exit Sub
'
com_select_delivery_Click_Error:
    DisplayError , "frm_job_details.com_select_delivery_Click", vbExclamation
End Sub

Private Sub fDeliveryCustomer_CustomerFound(oCustomer As Customer2, blnAbort As Boolean)
    Set moDeliveryCustomer = oCustomer
    With moDeliveryCustomer
        txt_del_customer_name.Text = .CustomerName
        txt_del_address_line_1.Text = .Address1
        txt_del_address_line_2.Text = .Address2
        txt_del_address_line_3.Text = .Address3
        txt_del_county_or_state.Text = .CountyOrState
        txt_del_country.Text = .Country
        txt_del_post_code.Text = .PostCode
        cmdOK.Default = True
    End With

End Sub

Private Sub fInputCustomer_CustomerFound(oCustomer As Customer2, blnAbort As Boolean)
'***************************************
' Module/Form Name   : frm_job_details
'
' Procedure Name     : fInputCustomer_CustomerFound
'
' Purpose            :
'
' Date Created       : 07/12/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : oCustomer - Customer2
'                    : blnAbort - Boolean
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo fInputCustomer_CustomerFound_Error
'
'******** Code Starts Here *************
'
    blnAbort = True
    If goSystemConfig.SageLink Then
        If Not oCustomer.TermsAndConditionsApproved Then
            MsgBox "Terms & Conditions have not yet been agreed for '" & oCustomer.CustomerName & "'.", vbExclamation
            Exit Sub
        End If
        If oCustomer.OnHold Then
            MsgBox "'" & oCustomer.CustomerName & "' is on hold." & vbCrLf, vbExclamation
            Exit Sub
        End If
    End If

    Set moCustomer = oCustomer
    With moCustomer
        .Contacts.Refresh
        .Comments.Refresh
    End With
    display_customer
    blnAbort = False
'
'********* Code Ends Here **************
'
   Exit Sub
'
fInputCustomer_CustomerFound_Error:
    DisplayError , "frm_job_details.fInputCustomer_CustomerFound", vbExclamation
End Sub

Private Sub Form_Load()
    '***************************************
    ' Module/Form Name   : frm_job_details
    '
    ' Procedure Name     : Form_Load
    '
    ' Purpose            :
    '
    ' Date Created       : 01/12/2002
    '
    ' Author             : GARETH SAUNDERS
    '
    ' Amendment History  : Date       Author    Description
    '                    : --------------------------------
    '                    : 01/12/2002 GARETH SAUNDERS
    '
    '***************************************
    '
    On Error GoTo Form_Load_Error
    '
    '******** Code Starts Here *************
    '
    With cboPreferredPhotoType
        .AddItem goSystemConfig.PhotoTypeDescription("T")
        .AddItem goSystemConfig.PhotoTypeDescription("D")
        .ListIndex = 0
    End With
    
    Set mobjContacts = New Contacts
    com_position_form Me
    '
    '********* Code Ends Here **************
    '
    Exit Sub
    '
Form_Load_Error:
    DisplayError , "frm_job_details.Form_Load", vbExclamation
End Sub

Private Sub cmdOK_Click()
'***************************************
' Module/Form Name   : frm_job_details
'
' Procedure Name     : cmdOK_Click
'
' Purpose            :
'
' Date Created       : 09/04/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 09/04/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo cmdOK_Click_Error
'
'******** Code Starts Here *************
'
    Dim fSearch As frm_search
    Dim colFrmSearch As New Collection
    Dim strMessage As String

    If Not valid_input Then
        Exit Sub
    End If
    If mod_action = "C" Then
        create_job
    Else
        On Error Resume Next
        update_job
        If Err.Number = vbObjectError + 12 Then
            strMessage = Err.Description
            roll_back
            MsgBox strMessage, vbExclamation, App.Title
            On Error GoTo cmdOK_Click_Error
            input_details "E", moJob
            Exit Sub
        ElseIf Err.Number = 0 Then
            On Error GoTo cmdOK_Click_Error
        Else
            ErrorSave
            On Error GoTo cmdOK_Click_Error
            ErrorRestore
        End If
    End If
    
    If Not retry_update Then
    Me.Hide
    DoEvents
    Select Case mod_action
        Case Is = "C"
            frm_job_edit.display_job moCustomer.CustomerNo, moJob.JobNo, moJob
        Case Else
            If is_form_loaded("frm_job_edit") Then
                frm_job_edit.display_job moCustomer.CustomerNo, moJob.JobNo, moJob
            End If
    End Select
    If is_form_loaded("frm_job_maint") Then
        'frm_job_maint.current_job_no = mod_job_no
        frm_job_maint.redisplay plngJobNo:=moJob.JobNo
    End If
    
    If is_form_loaded("frm_search", Me, colFrmSearch) Then
        For Each fSearch In colFrmSearch
            fSearch.display_title
        Next fSearch
    End If
    
    Unload Me
    Set frm_job_details = Nothing
    DoEvents
    End If
    
      '
      '********* Code Ends Here **************
      '
    Exit Sub
      '
cmdOK_Click_Error:
    DisplayError , "frm_job_details.cmdOK_Click", vbExclamation
End Sub

Private Sub Form_Paint()
''    On Error Resume Next
''    txt_job_reference.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set moJob = Nothing
    Set moCustomer = Nothing
End Sub

Private Sub txt_customer_name_Change()
    If Len(RTrim(txt_customer_name.Text)) = 0 Then
        com_default.Enabled = False
    Else
        com_default.Enabled = True
    End If
End Sub

Public Sub RefreshContactDetails()
    '***************************************
    ' Module/Form Name   : frm_job_details
    '
    ' Procedure Name     : RefreshContactDetails
    '
    ' Purpose            :
    '
    ' Date Created       : 25/05/2002
    '
    ' Author             : GARETH SAUNDERS
    '
    ' Amendment History  : Date       Author    Description
    '                    : --------------------------------
    '                    : 25/05/2002 GARETH SAUNDERS
    '
    '***************************************
    '
On Error GoTo RefreshContactDetails_Error
    '
    '******** Code Starts Here *************
    '
If moCustomer Is Nothing Then Exit Sub

moCustomer.Contacts.Refresh
SetDefaultRequestorsDetails
    '
    '********* Code Ends Here **************
    '
Exit Sub
    '
RefreshContactDetails_Error:
    ErrorRaise "frm_job_details.RefreshContactDetails"
End Sub
