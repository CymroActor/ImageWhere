VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCustEdit 
   AutoRedraw      =   -1  'True
   Caption         =   "Edit Customer"
   ClientHeight    =   5595
   ClientLeft      =   2520
   ClientTop       =   2235
   ClientWidth     =   8295
   ClipControls    =   0   'False
   Icon            =   "Customer Edit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5595
   ScaleWidth      =   8295
   Begin MSComctlLib.ImageList imlContacts 
      Left            =   5880
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   10
      ImageHeight     =   9
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Customer Edit.frx":0442
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Data datContacts 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "E:\Image Where\Development\Database\Library7.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4920
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.CommandButton ssc_done 
      Caption         =   "&OK"
      Height          =   555
      Left            =   1800
      Picture         =   "Customer Edit.frx":051C
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   4755
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.CommandButton ssc_cancel 
      Caption         =   "&Cancel"
      Height          =   555
      Left            =   4680
      Picture         =   "Customer Edit.frx":0A96
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   4755
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   795
   End
   Begin TabDlg.SSTab sstCustomer 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   8070
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "&General"
      TabPicture(0)   =   "Customer Edit.frx":1028
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lab_post_code_lab(9)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lab_faxno_lab(5)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lab_telephoneno_lab(4)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lab_customername_lab"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lab_address_lab"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lab_country_lab(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lab_county_or_state_lab(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblDateAmended"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtDateAmended"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txt_customer_name"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txt_address_line_1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txt_address_line_2"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txt_address_line_3"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txt_post_code"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txt_telephone_no"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txt_fax_no"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txt_country"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txt_county_or_state"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "Deli&very Information"
      TabPicture(1)   =   "Customer Edit.frx":1044
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cboBusinessType"
      Tab(1).Control(1)=   "fraContacts"
      Tab(1).Control(2)=   "txt_preferred_delivery_method"
      Tab(1).Control(3)=   "txt_vat_no"
      Tab(1).Control(4)=   "txt_country_code"
      Tab(1).Control(5)=   "txt_information"
      Tab(1).Control(6)=   "lab_vat_no_country_code_lab"
      Tab(1).Control(7)=   "lab_information_lab"
      Tab(1).Control(8)=   "lab_business_type_lab"
      Tab(1).Control(9)=   "lab_preferred_delivery_method_lab"
      Tab(1).Control(10)=   "lab_slash"
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "Priva&te"
      TabPicture(2)   =   "Customer Edit.frx":1060
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "smgComments"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "cmbNewComment"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cmbEditComment"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cmbDeleteComment"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "&Billing"
      TabPicture(3)   =   "Customer Edit.frx":107C
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "chkServiceFeeRequired"
      Tab(3).Control(1)=   "txtServiceFee"
      Tab(3).Control(2)=   "cboSageURN"
      Tab(3).Control(3)=   "lblSageOnHold"
      Tab(3).Control(4)=   "lblSageTermsAgreed"
      Tab(3).Control(5)=   "lblOnHold"
      Tab(3).Control(6)=   "lblTermsAndConditionsApproved"
      Tab(3).Control(7)=   "lblSageURN"
      Tab(3).ControlCount=   8
      Begin VB.CheckBox chkServiceFeeRequired 
         Alignment       =   1  'Right Justify
         Caption         =   "Service Fee"
         Height          =   255
         Left            =   -74300
         TabIndex        =   43
         Top             =   1920
         Width           =   2030
      End
      Begin VB.TextBox txtServiceFee 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -71880
         MaxLength       =   6
         TabIndex        =   44
         Top             =   1920
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ComboBox cboSageURN 
         Height          =   315
         Left            =   -72480
         TabIndex        =   38
         Top             =   720
         Width           =   1695
      End
      Begin VB.ComboBox cboBusinessType 
         DataField       =   "business_type"
         DataSource      =   "dat_business_type"
         Height          =   315
         Left            =   -72360
         TabIndex        =   32
         Top             =   4020
         Width           =   3975
      End
      Begin VB.CommandButton cmbDeleteComment 
         Caption         =   "&De&lete"
         Height          =   375
         Left            =   -72240
         TabIndex        =   36
         Top             =   3900
         Width           =   975
      End
      Begin VB.CommandButton cmbEditComment 
         Caption         =   "Ed&it..."
         Height          =   375
         Left            =   -73440
         TabIndex        =   35
         Top             =   3900
         Width           =   975
      End
      Begin VB.CommandButton cmbNewComment 
         Caption         =   "&Add..."
         Height          =   375
         Left            =   -74640
         TabIndex        =   34
         Top             =   3900
         Width           =   975
      End
      Begin VB.Frame fraContacts 
         Caption         =   "Contacts"
         Height          =   1695
         Left            =   -74880
         TabIndex        =   24
         Top             =   1740
         Width           =   7815
         Begin VB.CommandButton comEdit 
            Caption         =   "&Edit..."
            Height          =   315
            Left            =   4140
            TabIndex        =   26
            Top             =   1320
            Width           =   1095
         End
         Begin VB.CommandButton comDelete 
            Caption         =   "&Delete"
            Height          =   315
            Left            =   6600
            TabIndex        =   28
            Top             =   1320
            Width           =   1095
         End
         Begin VB.CommandButton comNew 
            Caption         =   "&New..."
            Height          =   315
            Left            =   5340
            TabIndex        =   27
            Top             =   1320
            Width           =   1095
         End
         Begin MSFlexGridLib.MSFlexGrid msfContacts 
            Bindings        =   "Customer Edit.frx":1098
            Height          =   1035
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   7575
            _ExtentX        =   13361
            _ExtentY        =   1826
            _Version        =   393216
            FixedCols       =   0
            WordWrap        =   -1  'True
            AllowBigSelection=   0   'False
            ScrollTrack     =   -1  'True
            FocusRect       =   0
            SelectionMode   =   1
            AllowUserResizing=   1
         End
      End
      Begin VB.TextBox txt_preferred_delivery_method 
         DataField       =   "preferred_delivery_method"
         DataSource      =   "dat_customer"
         Height          =   285
         Left            =   -72360
         MaxLength       =   30
         TabIndex        =   30
         Top             =   3600
         Width           =   2595
      End
      Begin VB.TextBox txt_vat_no 
         DataField       =   "vat_no"
         DataSource      =   "dat_customer"
         Height          =   285
         Left            =   -71760
         MaxLength       =   15
         TabIndex        =   23
         Top             =   1380
         Width           =   1815
      End
      Begin VB.TextBox txt_country_code 
         DataField       =   "country_code"
         DataSource      =   "dat_customer"
         Height          =   285
         Left            =   -72360
         MaxLength       =   3
         TabIndex        =   22
         Top             =   1380
         Width           =   435
      End
      Begin VB.TextBox txt_information 
         DataField       =   "information"
         DataSource      =   "dat_customer"
         Height          =   495
         Left            =   -72360
         MaxLength       =   300
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Top             =   780
         Width           =   5235
      End
      Begin VB.TextBox txt_county_or_state 
         DataField       =   "country"
         DataSource      =   "dat_customer"
         Height          =   285
         Left            =   2220
         MaxLength       =   20
         TabIndex        =   8
         Top             =   2040
         Width           =   2295
      End
      Begin VB.TextBox txt_country 
         DataField       =   "country"
         DataSource      =   "dat_customer"
         Height          =   285
         Left            =   2220
         MaxLength       =   20
         TabIndex        =   10
         Top             =   2400
         Width           =   2295
      End
      Begin VB.TextBox txt_fax_no 
         DataField       =   "fax_no"
         DataSource      =   "dat_customer"
         Height          =   285
         Left            =   2220
         MaxLength       =   20
         TabIndex        =   16
         Top             =   3600
         Width           =   1935
      End
      Begin VB.TextBox txt_telephone_no 
         DataField       =   "telephone_no"
         DataSource      =   "dat_customer"
         Height          =   285
         Left            =   2220
         MaxLength       =   20
         TabIndex        =   14
         Top             =   3240
         Width           =   1935
      End
      Begin VB.TextBox txt_post_code 
         DataField       =   "post_code"
         DataSource      =   "dat_customer"
         Height          =   285
         Left            =   2220
         MaxLength       =   12
         TabIndex        =   12
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox txt_address_line_3 
         DataField       =   "address_line_3"
         DataSource      =   "dat_customer"
         Height          =   285
         Left            =   2220
         MaxLength       =   40
         TabIndex        =   6
         Top             =   1680
         Width           =   4095
      End
      Begin VB.TextBox txt_address_line_2 
         DataField       =   "address_line_2"
         DataSource      =   "dat_customer"
         Height          =   285
         Left            =   2220
         MaxLength       =   40
         TabIndex        =   5
         Top             =   1320
         Width           =   4095
      End
      Begin VB.TextBox txt_address_line_1 
         DataField       =   "address_line_1"
         DataSource      =   "dat_customer"
         Height          =   285
         Left            =   2220
         MaxLength       =   40
         TabIndex        =   4
         Top             =   960
         Width           =   4095
      End
      Begin VB.TextBox txt_customer_name 
         DataField       =   "customer_name"
         DataSource      =   "dat_customer"
         Height          =   285
         Left            =   2220
         MaxLength       =   40
         TabIndex        =   2
         Top             =   600
         Width           =   4095
      End
      Begin ImageWhere.SimpleGrid smgComments 
         Height          =   2000
         Left            =   -74600
         TabIndex        =   33
         Top             =   500
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   3519
         Columns         =   1
         KeyCol          =   0
      End
      Begin VB.Label txtDateAmended 
         DataField       =   "faxno"
         DataSource      =   "dat_customer"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2220
         TabIndex        =   18
         Top             =   4020
         Width           =   1815
      End
      Begin VB.Label lblDateAmended 
         Caption         =   "Date Amended:"
         DataField       =   "faxno"
         DataSource      =   "dat_customer"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   4020
         Width           =   1215
      End
      Begin VB.Label lblSageOnHold 
         Height          =   255
         Left            =   -72480
         TabIndex        =   42
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label lblSageTermsAgreed 
         Height          =   255
         Left            =   -72480
         TabIndex        =   40
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label lblOnHold 
         Caption         =   "On Hold:"
         Height          =   255
         Left            =   -74280
         TabIndex        =   41
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label lblTermsAndConditionsApproved 
         Caption         =   "Terms Agreed:"
         Height          =   255
         Left            =   -74280
         TabIndex        =   39
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblSageURN 
         Caption         =   "Sage URN:"
         Height          =   255
         Left            =   -74280
         TabIndex        =   37
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lab_vat_no_country_code_lab 
         Caption         =   "Country Code/VAT No.:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74760
         TabIndex        =   21
         Top             =   1440
         Width           =   1755
      End
      Begin VB.Label lab_information_lab 
         Caption         =   "Information:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74760
         TabIndex        =   19
         Top             =   780
         Width           =   1635
      End
      Begin VB.Label lab_business_type_lab 
         Caption         =   "Business Type:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74760
         TabIndex        =   31
         Top             =   4080
         Width           =   1815
      End
      Begin VB.Label lab_preferred_delivery_method_lab 
         Caption         =   "Preferred Delivery Method:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74760
         TabIndex        =   29
         Top             =   3660
         Width           =   2115
      End
      Begin VB.Label lab_slash 
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -71880
         TabIndex        =   45
         Top             =   1380
         Width           =   135
      End
      Begin VB.Label lab_county_or_state_lab 
         Caption         =   "County/State:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   2100
         Width           =   1815
      End
      Begin VB.Label lab_country_lab 
         Caption         =   "Country:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   2460
         Width           =   1815
      End
      Begin VB.Label lab_address_lab 
         Caption         =   "Address:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   1395
      End
      Begin VB.Label lab_customername_lab 
         Caption         =   "Customer Name:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   660
         Width           =   1395
      End
      Begin VB.Label lab_telephoneno_lab 
         Caption         =   "Telephone No:"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   13
         Top             =   3240
         Width           =   1815
      End
      Begin VB.Label lab_faxno_lab 
         Caption         =   "Fax No:"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   15
         Top             =   3600
         Width           =   555
      End
      Begin VB.Label lab_post_code_lab 
         Caption         =   "Post  Code:"
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   11
         Top             =   2880
         Width           =   1815
      End
   End
   Begin VB.Label labTest 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   255
      Left            =   7080
      TabIndex        =   47
      Top             =   4920
      Visible         =   0   'False
      Width           =   1095
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmCustEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private mblnCustomerDisplayed As Boolean
Public snap_customer As ADODB.Recordset
Private dyna_customer As DAO.Recordset

Private retry_update As Boolean
Public cancel_list As Boolean
Private mbolDontResize As Boolean
Public mlngCurrentContact As Long
Private mstrCurrentComment As String
Public ContactsDisplayed As Boolean
Private Const FORM_MIN_HEIGHT = 5800
Private Const FORM_MIN_WIDTH = 8800
Private WithEvents fNewContact As frmNewContact
Attribute fNewContact.VB_VarHelpID = -1
Public mMode As UpdateMode
Public mstrCustomerName As String
Private WithEvents fCommentAdd As frmCommentAdd
Attribute fCommentAdd.VB_VarHelpID = -1
Private mobjCustomer As Customer2
Private moSageAccounts As SageAccounts

Private Sub cboSageURN_Click()
'***************************************
' Module/Form Name   : frmCustEdit
'
' Procedure Name     : cboSageURN_Click
'
' Purpose            :
'
' Date Created       : 03/04/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 03/04/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo cboSageURN_Click_Error
'
'******** Code Starts Here *************
'
    If cboSageURN.Text = mobjCustomer.SageURN Then
        Exit Sub
    End If

    If mblnCustomerDisplayed Then
        If MsgBox("Are you sure you wish to change the Sage URN Link?", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
            mblnCustomerDisplayed = False
            If Not LocateComboItem(cboSageURN, Format(mobjCustomer.SageURN)) Then
                cboSageURN.ListIndex = 0
            End If
            RefreshSageFields
            mblnCustomerDisplayed = True
            Exit Sub
        End If
    End If

    If cboSageURN.ListIndex = 0 Then
        mobjCustomer.SageURN = ""
    Else
        mobjCustomer.SageURN = cboSageURN.Text
    End If
    RefreshSageFields
    Exit Sub
    '
    '********* Code Ends Here **************
    '
    Exit Sub
    '
cboSageURN_Click_Error:
    DisplayError , "frmCustEdit.cboSageURN_Click", vbExclamation
End Sub

Private Sub chkServiceFeeRequired_Click()
    txtServiceFee.Visible = chkServiceFeeRequired.Value
End Sub

Private Sub cmbDeleteComment_Click()
'***************************************
' Module/Form Name   : frmCustEdit
'
' Procedure Name     : cmbDeleteComment_Click
'
' Purpose            :
'
' Date Created       : 03/04/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 03/04/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo cmbDeleteComment_Click_Error
'
'******** Code Starts Here *************
'
    If smgComments.CurrentRow = 0 Then
        MsgBox "A Comment must be selected", vbInformation
        Exit Sub
    End If
    If MsgBox("Are you sure you wish to delete '" & smgComments.Column(2).Value & "'", vbQuestion + vbYesNo + vbDefaultButton2, "Delete Contact") = vbYes Then
        On Error Resume Next
        mobjCustomer.Comments(smgComments.Column(smgComments.KeyCol).Value).Delete
        If Err.Number = vbObjectError + 5 Then
            DisplayError
        ElseIf Err.Number <> 0 Then
            GoTo cmbDeleteComment_Click_Error
        Else
            On Error GoTo cmbDeleteComment_Click_Error
        End If
        mobjCustomer.Comments.Refresh
        sstCustomer.Visible = False
        RefreshComments
        sstCustomer.Visible = True
        Me.Refresh
    End If
    '
    '********* Code Ends Here **************
    '
    Exit Sub
    '
cmbDeleteComment_Click_Error:
    DisplayError , "frmCustEdit.cmbDeleteComment_Click", vbExclamation
End Sub

Private Sub cmbEditComment_Click()
'***************************************
' Module/Form Name   : frmCustEdit
'
' Procedure Name     : cmbEditComment_Click
'
' Purpose            :
'
' Date Created       : 03/04/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 03/04/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo cmbEditComment_Click_Error
'
'******** Code Starts Here *************
'
    If smgComments.CurrentRow = 0 Then
        MsgBox "A Comment must be selected", vbInformation
        Exit Sub
    End If
    
    Set fCommentAdd = Nothing
    Set fCommentAdd = New frmCommentAdd
    With fCommentAdd
        .mbolEditComment = True
        Set .mobjCustomer = mobjCustomer
        .mstrID = smgComments.Column(4).Value
        .Show 1
    End With
    Me.SetFocus
    sstCustomer.Visible = False
    sstCustomer.Visible = True
    Set fCommentAdd = Nothing
    '
    '********* Code Ends Here **************
    '
    Exit Sub
    '
cmbEditComment_Click_Error:
    DisplayError , "frmCustEdit.cmbEditComment_Click", vbExclamation
End Sub

Private Sub cmbNewComment_Click()
'***************************************
' Module/Form Name   : frmCustEdit
'
' Procedure Name     : cmbNewComment_Click
'
' Purpose            :
'
' Date Created       : 03/04/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 03/04/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo cmbNewComment_Click_Error
'
'******** Code Starts Here *************
'
    Set fCommentAdd = Nothing
    Set fCommentAdd = New frmCommentAdd
    With fCommentAdd
        .mbolEditComment = False
        Set .mobjCustomer = mobjCustomer
        .Show 1
    End With
    Set fCommentAdd = Nothing
    Me.SetFocus
    '
    '********* Code Ends Here **************
    '
    Exit Sub
    '
cmbNewComment_Click_Error:
    DisplayError , "frmCustEdit.cmbNewComment_Click", vbExclamation
End Sub

Private Sub comDelete_Click()
'***************************************
' Module/Form Name   : frmCustEdit
'
' Procedure Name     : comDelete_Click
'
' Purpose            :
'
' Date Created       : 03/04/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo comDelete_Click_Error
'
'******** Code Starts Here *************
'
    If msfContacts.row = 0 Then
        MsgBox "A Contact must be selected", vbInformation
        Exit Sub
    End If

    If MsgBox("Are you sure you wish to delete '" & msfContacts.TextArray(FlexGridIndex(msfContacts, msfContacts.row, 2)) & "'", vbQuestion + vbYesNo + vbDefaultButton2, "Delete Contact") = vbYes Then
        With msfContacts
            On Error Resume Next
            mobjCustomer.Contacts.Item(.TextMatrix(.row, 8)).Delete
            If Err.Number = vbObjectError + 5 Then
                DisplayError , "frmCustEdit.comDelete_Click", vbExclamation
            ElseIf Err.Number <> 0 Then
                GoTo comDelete_Click_Error
            Else
                On Error GoTo comDelete_Click_Error
            End If
        End With
        mobjCustomer.Contacts.Refresh
        DisplayContacts
    End If
    '
    RefreshScreens
    '
    '********* Code Ends Here **************
    '
       Exit Sub
    '
comDelete_Click_Error:
    DisplayError , "frmCustEdit.comDelete_Click", vbExclamation
End Sub

Private Sub RefreshScreens()
    Dim f As Form

    For Each f In Forms
        '
        '   Refresh the Edit Job screen.
        '
        If f.Name = "frm_job_edit" Then
            f.RefreshJob
        End If
    Next f
End Sub

Private Sub comEdit_Click()
'***************************************
' Module/Form Name   : frmCustEdit
'
' Procedure Name     : comEdit_Click
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
On Error GoTo comEdit_Click_Error
'
'******** Code Starts Here *************
'
    If msfContacts.row = 0 Then
        MsgBox "A contact must be selected", vbInformation
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Set fNewContact = New frmNewContact
    fNewContact.DisplayForm mobjCustomer.CustomerNo, mlngCurrentContact
    Set fNewContact = Nothing
    Me.SetFocus
    Screen.MousePointer = vbDefault
'
'********* Code Ends Here **************
'
   Exit Sub
'
comEdit_Click_Error:
    DisplayError , "frmCustEdit.comEdit_Click", vbExclamation
End Sub

Private Sub comNew_Click()
'***************************************
' Module/Form Name   : frmCustEdit
'
' Procedure Name     : comNew_Click
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
On Error GoTo comNew_Click_Error
'
'******** Code Starts Here *************
'
    Set fNewContact = New frmNewContact
    fNewContact.DisplayForm mobjCustomer.CustomerNo
    Set fNewContact = Nothing
    Me.SetFocus
'
'********* Code Ends Here **************
'
   Exit Sub
'
comNew_Click_Error:
    DisplayError , "frmCustEdit.comNew_Click", vbExclamation
End Sub

Private Sub display_customer()
'***************************************
' Module/Form Name   : frmCustEdit
'
' Procedure Name     : display_customer
'
' Purpose            :
'
' Date Created       : 03/04/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 03/04/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo display_customer_Error
'
'******** Code Starts Here *************
'
    Dim strSQL As String

    Screen.MousePointer = vbHourglass

    mblnCustomerDisplayed = False

        'Me.Visible = False
    If mMode = Edit Then
        With mobjCustomer
            txt_customer_name = .CustomerName
            txt_address_line_1 = .Address1
            txt_address_line_2 = .Address2
            txt_address_line_3 = .Address3
            txt_post_code = .PostCode
            txt_telephone_no = .Telephone
            txt_county_or_state = .CountyOrState
            txt_country = .Country
            txt_fax_no = .Fax
            txt_information = .Information
            txt_country_code = .CountryCode
            txt_vat_no = .VATNo
            If RTrim(txt_country_code) = "" Or RTrim(txt_vat_no) = "" Then
                lab_slash = ""
            Else
                lab_slash = "/"
            End If
            txt_preferred_delivery_method.Text = .PreferredDelivery
            If Not LocateComboItem(cboBusinessType, .BusinessType) Then
                MsgBox "Business Type for customer not found on database", vbExclamation
            End If
            If goSystemConfig.SageLink Then
                If Not LocateComboItem(cboSageURN, Format(.SageURN)) Then
                    cboSageURN.ListIndex = 0
                End If
                RefreshSageFields
            End If
            chkServiceFeeRequired.Value = IIf(.ServiceFeeRequired, vbChecked, vbUnchecked)
            txtServiceFee.Text = .ServiceFee
            If .ServiceFeeRequired Then
                txtServiceFee.Visible = True
            Else
                txtServiceFee.Visible = False
            End If
            txtDateAmended = IIf(.DateAmended = 0, "", Format(.DateAmended, "dd/mm/yyyy"))
        End With
    End If

        'Me.Visible = True
    mblnCustomerDisplayed = True
    sstCustomer_Click (0)
    Screen.MousePointer = vbDefault
'
'********* Code Ends Here **************
'
    Exit Sub
display_customer_Error:
    Me.Visible = True
    ErrorRaise "frmCustEdit.display_customer"
End Sub

Public Sub DisplayContacts()
'***************************************
' Module/Form Name   : frmCustEdit
'
' Procedure Name     : DisplayContacts
'
' Purpose            :
'
' Date Created       : 03/04/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 03/04/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo DisplayContacts_Error
'
'******** Code Starts Here *************
'
    Dim intCount As Integer
    Dim strSQL As String
    Dim intTopRow As Integer

    '
    '   Don't redraw the grid until changes have been made.
    '
    msfContacts.Redraw = False
    '
    If mMode = Edit Then
        strSQL = "SELECT WebUser as Web, MAILINGLIST as Mail, Name, Position, Phone, Fax, Email, Comments, ID " _
               & "FROM CONTACTS " _
               & "WHERE CUSTOMER_NO = " & mobjCustomer.CustomerNo & " ORDER BY NAME"
    Else
        strSQL = "SELECT WebUser as Web, MAILINGLIST as Mail, Name, Position, Phone, Comments, ID " _
               & "FROM CONTACTS " _
               & "WHERE CUSTOMER_NO = 0"
    End If

    With datContacts
        '.Connect
        .RecordSource = strSQL
        .Refresh
        .UpdateControls
    End With

    With msfContacts
        For intCount = 1 To .Rows - 1
            '
            '   Web User Graphic.
            '
            If .TextMatrix(intCount, 0) = True Then
                .col = 0
                .row = intCount
                Set .CellPicture = imlContacts.ListImages(1).Picture
                .CellPictureAlignment = flexAlignCenterTop
            End If
            '
            '   Mailing List Graphic.
            '
            If .TextMatrix(intCount, 1) = True Then
                .col = 1
                .row = intCount
                Set .CellPicture = imlContacts.ListImages(1).Picture
                .CellPictureAlignment = flexAlignCenterTop
            End If
            .TextMatrix(intCount, 0) = ""
            .TextMatrix(intCount, 1) = ""
        Next
    End With

    If mlngCurrentContact = 0 Then
        If msfContacts.row <> 0 Then
            HighLightRow msfContacts, 1
            mlngCurrentContact = CLng(IIf(msfContacts.TextMatrix(msfContacts.RowSel, 8) = "", 0, msfContacts.TextMatrix(msfContacts.RowSel, 8)))
        End If
    Else
        HighLightRow msfContacts, ContactGridPosition(msfContacts, mlngCurrentContact)
    End If
    '
    '   Resize rows if necessary.
    '
    intTopRow = msfContacts.TopRow
    ResizeContactRows
    If datContacts.Recordset.RecordCount <> 0 Then
        msfContacts.TopRow = intTopRow
    End If
    '
    msfContacts_SelChange
    '
    '   Now redraw.
    '
    msfContacts.Redraw = True
    DoEvents
    '
    '   Indicate that Contacts have now been displayed.
    '
    ContactsDisplayed = True
    '
    '
    '********* Code Ends Here **************
    '
    Exit Sub
    '
DisplayContacts_Error:
    ErrorRaise "frmCustEdit.DisplayContacts"
End Sub

Private Sub fCommentAdd_CommentsUpdated(ID As String)
'***************************************
' Module/Form Name   : frmCustEdit
'
' Procedure Name     : fCommentAdd_CommentsUpdated
'
' Purpose            :
'
' Date Created       : 31/12/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : ID - String
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo fCommentAdd_CommentsUpdated_Error
'
'******** Code Starts Here *************
'
    sstCustomer.Visible = False
    mstrCurrentComment = ID
    RefreshComments
    sstCustomer.Visible = True
'
'********* Code Ends Here **************
'
   Exit Sub
'
fCommentAdd_CommentsUpdated_Error:
    DisplayError , "frmCustEdit.fCommentAdd_CommentsUpdated", vbExclamation
End Sub

Private Sub fNewContact_ContactCreated(ContactID As Long)
'***************************************
' Module/Form Name   : frmCustEdit
'
' Procedure Name     : fNewContact_ContactCreated
'
' Purpose            :
'
' Date Created       : 31/12/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : ContactName - String
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 31/12/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo fNewContact_ContactCreated_Error
'
'******** Code Starts Here *************
'
    mlngCurrentContact = ContactID
    mobjCustomer.Contacts.Refresh
    DisplayContacts
'
'********* Code Ends Here **************
'
   Exit Sub
'
fNewContact_ContactCreated_Error:
    DisplayError , "frmCustEdit.fNewContact_ContactCreated", vbExclamation
End Sub

Private Sub Form_Load()
'***************************************
' Module/Form Name   : frmCustEdit
'
' Procedure Name     : Form_Load
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
On Error GoTo Form_Load_Error
'
'******** Code Starts Here *************
'
    Screen.MousePointer = vbHourglass
''    Resize.Hook Me.hWnd500, 400
'    DoEvents
    '
    '   Enable or otherwise SAGE fields.
    '
    If goSystemConfig.SageLink Then
        lblSageURN.Enabled = True
        cboSageURN.Enabled = True
        lblOnHold.Enabled = True
        lblTermsAndConditionsApproved.Enabled = True
        lblSageOnHold.Enabled = True
        lblSageTermsAgreed.Enabled = True
          Else
        lblSageURN.Enabled = False
        cboSageURN.Enabled = False
        lblOnHold.Enabled = False
        lblTermsAndConditionsApproved.Enabled = False
        lblSageOnHold.Enabled = False
        lblSageTermsAgreed.Enabled = False
    End If
    '
    '   Initialize the comments grid.
    '
    With smgComments
        .Columns = 6
        .KeyCol = 4
        .Column(1).Header = "Who"
        .Column(2).Header = "Comment"
        .Column(2).Align = flexAlignLeftTop
        .Column(3).Header = "Date"
        .Column(3).Width = 1800
        .ColumnProperty(3).PropertyType = smgPropTypeDateLong
        .Column(4).Width = 0
        .Column(5).Width = 0
        .Column(6).Width = 1000
        .Column(2).Width = .Width - .Column(1).Width - .Column(3).Width - 354
    End With
    '
    '   Initialize the Contacts Grid.
    '
    ContactsDisplayed = False
    '
    '   Set the width & height of the form.
    '
    sstCustomer.Tab = 0
    mbolDontResize = True

    Me.WindowState = GetSetting(App.Title, "frmCustEdit", "WindowState", 0)
    If GetSetting(App.Title, "frmCustEdit", "WindowState", 0) <> 2 Then
        Me.Width = GetSetting(App.Title, "frmCustEdit", "Width", FORM_MIN_WIDTH)
        Me.Height = GetSetting(App.Title, "frmCustEdit", "Height", FORM_MIN_HEIGHT)
    End If
    mbolDontResize = False
    com_position_form Me
    InitialiseBusinessTypes
    If goSystemConfig.SageLink Then
        InitialiseSageURNs
    End If
    If mMode = Edit Then
        If Not load_customer(mstrCustomerName) Then
            Screen.MousePointer = vbDefault
            Unload Me
            Exit Sub
        End If
        Me.Caption = "Edit Customer: " & mobjCustomer.CustomerName
    Else
        Me.Caption = "New Customer"
    End If

    InitialiseContacts
    DisplayContacts
    '
    RefreshComments
    '
    '   Show the form.
    '
    Me.Show
    DoEvents
    display_customer
    On Error Resume Next
    txt_customer_name.SetFocus
    Screen.MousePointer = vbDefault
'
'********* Code Ends Here **************
'
    Exit Sub
    '
Form_Load_Error:
    DisplayError , "frmCustEdit.Form_Load", vbExclamation
    Unload Me
End Sub

Private Sub Form_Resize()
'***************************************
' Module/Form Name   : frmCustEdit
'
' Procedure Name     : Form_Resize
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
On Error GoTo Form_Resize_Error
'
'******** Code Starts Here *************
'
    Dim strCurrentRow As String
    Dim Ctl As Control, CtlCln As New Collection
'
'   Ignore the first resize as this is caused through setting the
'   initial dimensions of the form.
'
    If Not ResizeForm(Me) Then
        Exit Sub
    End If
    '
    If mbolDontResize Then
        Exit Sub
    End If
    
    If WindowState = 1 Then
        Exit Sub
    End If
  
'    LockWindow Me.hWnd
    If WindowState = 0 Then
        If Width < FORM_MIN_WIDTH Then
            Width = FORM_MIN_WIDTH
            Exit Sub
        End If
        If Height < FORM_MIN_HEIGHT Then
            Height = FORM_MIN_HEIGHT
            Exit Sub
        End If
    End If
    sstCustomer.Visible = False
    ssc_done.Visible = False
    ssc_cancel.Visible = False
'    Me.Visible = False
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
    On Error GoTo Form_Resize_Error
'
'   Reposition Edit & Cancel buttons.
'
    ssc_done.Move (Width - (ssc_done.Width + ssc_cancel.Width + 500)) / 2, _
                  Me.Height - ssc_done.Height - 600
    ssc_cancel.Move ssc_done.Left + ssc_done.Width + 500, _
                    ssc_done.Top
'
'   Resize tab.
'
    With sstCustomer
        .Width = Me.Width - .Left * 2 - 200
        .Height = ssc_done.Top - .Top - 100
    End With
'
'   Resize Tab 1 fields.
'
    If mMode = Edit Then
        ResizeContacts
    End If
    txt_information.Left = txt_country_code.Left
    txt_information.Width = sstCustomer.Width - txt_information.Left - 100
    lab_preferred_delivery_method_lab.Top = fraContacts.Top + fraContacts.Height + 200
    txt_preferred_delivery_method.Top = lab_preferred_delivery_method_lab.Top
    lab_business_type_lab.Top = lab_preferred_delivery_method_lab.Top + 400
    cboBusinessType.Top = lab_business_type_lab.Top
    '
    '   Resize the Comments.
    '
    cmbNewComment.Top = sstCustomer.Height - cmbNewComment.Height - 200
    cmbEditComment.Top = cmbNewComment.Top
    cmbDeleteComment.Top = cmbNewComment.Top
    
    On Error Resume Next
    With smgComments
        .Left = 200
        .Width = sstCustomer.Width - (2 * smgComments.Left)
        .Height = cmbNewComment.Top - .Top - 200
        .Column(2).Width = .Width - .Column(1).Width - .Column(3).Width - 354
    End With
    On Error GoTo Form_Resize_Error
    '
    '   Refresh Comments.
    '
    With smgComments
        strCurrentRow = .Column(.KeyCol).Value
'        RefreshComments
'        .GetKeyRow (CStr(strCurrentRow))
        .ResizeRows
    End With
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
    ssc_done.Visible = True
    ssc_cancel.Visible = True
'    UnlockWindow
'
'********* Code Ends Here **************
'
   Exit Sub
'
Form_Resize_Error:
    DisplayError , "frmCustEdit.Form_Resize", vbExclamation
    sstCustomer.Visible = True
700       ssc_done.Visible = True
710       ssc_cancel.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting App.Title, "frmCustEdit", "Width", Me.Width
    SaveSetting App.Title, "frmCustEdit", "Height", Me.Height
    SaveSetting App.Title, "frmCustEdit", "WindowState", Me.WindowState
    On Error Resume Next
    gcolMaxedWindows.Remove CStr(Me.hWnd)
End Sub

Private Sub InitialiseBusinessTypes()
'***************************************
' Module/Form Name   : frmCustEdit
'
' Procedure Name     : InitialiseBusinessTypes
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
On Error GoTo InitialiseBusinessTypes_Error
'
'******** Code Starts Here *************
'
    Dim rstBusinessType As DAO.Recordset
    Dim strSQL As String

    strSQL = "SELECT * FROM BUSINESS_TYPE"
    Set rstBusinessType = db.OpenRecordset(strSQL, dbOpenSnapshot)
    With rstBusinessType
        If .RecordCount <> 0 Then
            .MoveFirst
        End If
        While Not .EOF
            cboBusinessType.AddItem !Business_Type
            .MoveNext
        Wend
    End With
'
'********* Code Ends Here **************
'
   Exit Sub
'
InitialiseBusinessTypes_Error:
    ErrorRaise "frmCustEdit.InitialiseBusinessTypes"
End Sub

Private Sub InitialiseContacts()
'***************************************
' Module/Form Name   : frmCustEdit
'
' Procedure Name     : InitialiseContacts
'
' Purpose            :
'
' Date Created       : 03/04/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 03/04/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo InitialiseContacts_Error
'
'******** Code Starts Here *************
'
    Dim intCount As Integer
    datContacts.DatabaseName = glo_dbname
    If mMode = Add Then
        fraContacts.Visible = False
        fraContacts.Height = 0
    Else
        With msfContacts
            .Visible = True
            .Cols = 8         ' Set number of columns.
        End With
    End If
    mlngCurrentContact = 0
'
'********* Code Ends Here **************
'
    Exit Sub
    '
InitialiseContacts_Error:
    ErrorRaise "frmCustEdit.InitialiseContacts"
End Sub

Private Sub InitialiseSageURNs()
'***************************************
' Module/Form Name   : frmCustEdit
'
' Procedure Name     : InitialiseSageURNs
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
On Error GoTo InitialiseSageURNs_Error
'
'******** Code Starts Here *************
'
    Dim oSageAccount As SageAccount

    cboSageURN.AddItem "<Not Set>"
    Set moSageAccounts = New SageAccounts
    For Each oSageAccount In moSageAccounts
        cboSageURN.AddItem oSageAccount.Key
    Next oSageAccount
'
'********* Code Ends Here **************
'
   Exit Sub
'
InitialiseSageURNs_Error:
    ErrorRaise "frmCustEdit.InitialiseSageURNs"
End Sub

Public Function load_customer(customer_name As String) As Boolean
'***************************************
' Module/Form Name   : frmCustEdit
'
' Procedure Name     : load_customer
'
' Purpose            :
'
' Date Created       : 31/07/2001
'
' Author             : GARETH
'
' Parameters         : customer_name - String
'
' Returns            : Boolean
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 09/12/2001 GARETH
'
'***************************************
'
On Error GoTo load_customer_Error
'
'******** Code Starts Here *************
'
    'Initialise load_customer as failed
    load_customer = False
    If mMode = Add Then
        display_customer
        Exit Function
    End If
    Set mobjCustomer = New Customer2
    On Error Resume Next
    mobjCustomer.Locate customer_name
    If Err.Number = vbObjectError + 7 Then  'Cancelled
        Exit Function
    ElseIf Err.Number = vbObjectError + 6 Then ' Customer not found
        MsgBox "No Customer found for editing", vbExclamation, "Image Where?"
        Exit Function
    ElseIf Err.Number <> 0 Then
        ErrorSave
        On Error GoTo load_customer_Error
        ErrorRestore
    Else
        On Error GoTo load_customer_Error
    End If
    load_customer = True
'
'********* Code Ends Here **************
'
    Exit Function
    '
load_customer_Error:
    ErrorRaise "frmCustEdit.load_customer"
End Function

Private Sub msfContacts_DblClick()
'***************************************
' Module/Form Name   : frmCustEdit
'
' Procedure Name     : msfContacts_DblClick
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
On Error GoTo msfContacts_DblClick_Error
'
'******** Code Starts Here *************
'
    Screen.MousePointer = vbHourglass
    Set fNewContact = New frmNewContact
    fNewContact.DisplayForm mobjCustomer.CustomerNo, mlngCurrentContact
    Set fNewContact = Nothing
    Me.SetFocus
    Screen.MousePointer = vbDefault
'
'********* Code Ends Here **************
'
   Exit Sub
'
msfContacts_DblClick_Error:
    DisplayError , "frmCustEdit.msfContacts_DblClick", vbExclamation
End Sub

Private Sub msfContacts_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'***************************************
' Module/Form Name   : frmCustEdit
'
' Procedure Name     : msfContacts_MouseDown
'
' Purpose            :
'
' Date Created       : 31/12/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : Button - Integer
'                    : Shift - Integer
'                    : X - Single
'                    : Y - Single
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo msfContacts_MouseDown_Error
'
'******** Code Starts Here *************
'
    Dim objContact As Contact
    Dim ShiftPressed As Integer
    ShiftPressed = (Shift And vbShiftMask) > 0

    '
    '   Only allow a left button click on the mail column.
    '
    Screen.MousePointer = vbHourglass
    If Button <> vbLeftButton Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    '
    With msfContacts
        '
        '   Inhibit multiple selection using the shift key.
        '
        If ShiftPressed Then
            HighLightRow msfContacts, .RowSel
            mlngCurrentContact = CLng(.TextMatrix(.RowSel, 8))
        End If
        '
        '   Set the current row.
        '
        mlngCurrentContact = CLng(.TextMatrix(.row, 8))
        If (X < .ColPos(0) + .ColWidth(0) + .ColWidth(1)) And _
           (X > .ColPos(0) + .ColWidth(0)) Then
            '
            '   The user has clicked on the mail column. Work out which row.
            '
            Set objContact = New Contact
            With mobjCustomer.Contacts.Item(msfContacts.TextMatrix(.row, 8))
                .Mail = Not .Mail
                On Error Resume Next
                .update
                If (Err.Number - vbObjectError) = 5 Then
                    DisplayError
                Else
                    If Err.Number <> 0 Then
                        ErrorSave
                        On Error GoTo msfContacts_MouseDown_Error
                        ErrorRestore
                    Else
                        On Error GoTo msfContacts_MouseDown_Error
                    End If
                End If
            End With
            Set objContact = Nothing
            DisplayContacts
        End If
    End With
    Screen.MousePointer = vbDefault
'
'********* Code Ends Here **************
'
   Exit Sub
'
msfContacts_MouseDown_Error:
    DisplayError , "frmCustEdit.msfContacts_MouseDown", vbExclamation
End Sub

Private Sub msfContacts_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'***************************************
' Module/Form Name   : frmCustEdit
'
' Procedure Name     : msfContacts_MouseMove
'
' Purpose            :
'
' Date Created       : 31/12/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : Button - Integer
'                    : Shift - Integer
'                    : X - Single
'                    : Y - Single
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo msfContacts_MouseMove_Error
'
'******** Code Starts Here *************
'
    With msfContacts
        If Button = vbLeftButton Then
            If .row <> .RowSel Then
                HighLightRow msfContacts, .RowSel
                mlngCurrentContact = CLng(.TextArray(FlexGridIndex(msfContacts, .RowSel, 8)))
            End If
        End If
    End With
'
'********* Code Ends Here **************
'
   Exit Sub
'
msfContacts_MouseMove_Error:
    DisplayError , "frmCustEdit.msfContacts_MouseMove", vbExclamation
End Sub

Private Sub msfContacts_Scroll()
'***************************************
' Module/Form Name   : frmCustEdit
'
' Procedure Name     : msfContacts_Scroll
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
On Error GoTo msfContacts_Scroll_Error
'
'******** Code Starts Here *************
'
    With msfContacts
        If .row <> .RowSel Then
            HighLightRow msfContacts, .RowSel
            mlngCurrentContact = CLng(.TextArray(FlexGridIndex(msfContacts, .RowSel, 8)))
        End If
    End With
'
'********* Code Ends Here **************
'
   Exit Sub
'
msfContacts_Scroll_Error:
    DisplayError , "frmCustEdit.msfContacts_Scroll", vbExclamation
End Sub

Private Sub msfContacts_SelChange()
'***************************************
' Module/Form Name   : frmCustEdit
'
' Procedure Name     : msfContacts_SelChange
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
On Error GoTo msfContacts_SelChange_Error
'
'******** Code Starts Here *************
'
    '
    '   Set tooltip text to be the Comments!
    '
    With msfContacts
        .ToolTipText = .TextArray(FlexGridIndex(msfContacts, .row, 7))
    End With
'
'********* Code Ends Here **************
'
   Exit Sub
'
msfContacts_SelChange_Error:
    DisplayError , "frmCustEdit.msfContacts_SelChange", vbExclamation
End Sub

Private Sub RefreshComments()
'***************************************
' Module/Form Name   : frmCustEdit
'
' Procedure Name     : RefreshComments
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
On Error GoTo RefreshComments_Error
'
'******** Code Starts Here *************
'
    Dim objComment As Comment

    With smgComments
        .Redraw = False
        .Clear
        For Each objComment In mobjCustomer.Comments
            .AddRow False, _
                    objComment.Author, _
                    objComment.Comment, _
                    objComment.DateWritten, _
                    objComment.Key, _
                    DateDiff("s", CDate("01/01/1900"), objComment.DateWritten)
        Next objComment
        .ResizeRows
        .Column(5).Sorted = smgDescending
        .GetKeyRow (mstrCurrentComment)
        .Redraw = True
    End With

'
'********* Code Ends Here **************
'
   Exit Sub
'
RefreshComments_Error:
    ErrorRaise "frmCustEdit.RefreshComments"
End Sub

Private Sub RefreshSageFields()
'***************************************
' Module/Form Name   : frmCustEdit
'
' Procedure Name     : RefreshSageFields
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
On Error GoTo RefreshSageFields_Error
'
'******** Code Starts Here *************
'
    With mobjCustomer
        lblSageTermsAgreed = IIf(.TermsAndConditionsApproved, "Yes", "No")
        lblSageOnHold = IIf(.OnHold, "Yes", "No")
    End With
'
'********* Code Ends Here **************
'
   Exit Sub
'
RefreshSageFields_Error:
    ErrorRaise "frmCustEdit.RefreshSageFields"
End Sub

Private Sub ResizeContactRows()
'***************************************
' Module/Form Name   : frmCustEdit
'
' Procedure Name     : ResizeContactRows
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
On Error GoTo ResizeContactRows_Error
'
'******** Code Starts Here *************
'
    Dim intCol As Integer
    Dim intRow As Integer
    Dim intSaveRow As Integer

    With msfContacts
        intSaveRow = .row
        For intRow = 0 To .Rows - 1
            .row = intRow
            .RowHeight(intRow) = 255
            For intCol = 0 To .Cols - 1
                If intCol <> 7 Then
                    .col = intCol
                    .CellAlignment = flexAlignLeftTop
                    '
                    '   For each column decide whether to increase size of row.
                    '
                    If .ColWidth(intCol) >= 100 Then
                        labTest.Width = .ColWidth(intCol) - 100
                        labTest.Caption = .TextArray(FlexGridIndex(msfContacts, intRow, intCol))
                        If .RowHeight(intRow) < labTest.Height Then
                            .RowHeight(intRow) = labTest.Height
                        End If
                    End If
                End If
            Next intCol
        Next intRow
        FlexGridHighlightRow msfContacts, intSaveRow
    End With
'
'********* Code Ends Here **************
'
   Exit Sub
'
ResizeContactRows_Error:
    ErrorRaise "frmCustEdit.ResizeContactRows"
End Sub

Private Sub ResizeContacts()
'***************************************
' Module/Form Name   : frmCustEdit
'
' Procedure Name     : ResizeContacts
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
On Error GoTo ResizeContacts_Error
'
'******** Code Starts Here *************
'
    Dim lngMaxHeight As Long

    With fraContacts
        .Width = sstCustomer.Width - fraContacts.Left * 2
        .Height = sstCustomer.Height - 1200 - .Top
    End With
    '
    '   Resize Contacts Buttons.
    '
    comDelete.Move fraContacts.Width - 200 - comDelete.Width, _
                   fraContacts.Height - 100 - comDelete.Height
    comNew.Move comDelete.Left - 200 - comNew.Width, _
                comDelete.Top
    comEdit.Move comNew.Left - 200 - comEdit.Width, _
                 comDelete.Top
    '
    '   Resize Contacts Grid.
    '
    On Error Resume Next
    With msfContacts
        lngMaxHeight = (comEdit.Top - 100 - .Top)
        .Width = fraContacts.Width - 200
        .Left = 50
        .Height = lngMaxHeight - ((lngMaxHeight - 75) Mod .RowHeight(0))
        .ColWidth(0) = 450
        .ColWidth(1) = 400
        .ColWidth(4) = 1200
        .ColWidth(5) = 1200
        .ColWidth(6) = 2400
        .ColWidth(3) = 1200
        .ColWidth(2) = Int(0.5 * (.Width - .ColWidth(0) - _
                                           .ColWidth(1) - _
                                           .ColWidth(4) - _
                                           .ColWidth(5) - _
                                           .ColWidth(6) - _
                                           .ColWidth(3) - scroll_bar_width - 2))
        .ColWidth(7) = .ColWidth(2)
        .ColWidth(8) = 0
        ResizeContactRows
        .Redraw = True
    End With
'
'********* Code Ends Here **************
'
   Exit Sub
'
ResizeContacts_Error:
    ErrorRaise "frmCustEdit.ResizeContacts"
End Sub

Private Sub smgComments_RowChanged(CurrentRow As String)
    If CurrentRow = "" Then
        mstrCurrentComment = "0"
    Else
        mstrCurrentComment = CurrentRow
    End If
End Sub

Private Sub ssc_cancel_Click()
    Unload Me
    Set frmCustEdit = Nothing
End Sub

Private Sub ssc_done_Click()
'***************************************
' Module/Form Name   : frmCustEdit
'
' Procedure Name     : ssc_done_Click
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
On Error GoTo ssc_done_Click_Error
'
'******** Code Starts Here *************
'
    Dim objCustomer As Customer2

    If Not valid_input Then
        Exit Sub
    End If
    If mMode = Add Then
        Set objCustomer = New Customer2
        With objCustomer
            .CustomerName = txt_customer_name.Text
            .Address1 = txt_address_line_1
            .BusinessType = cboBusinessType.Text
            On Error Resume Next
            .Create
            Select Case Err.Number - vbObjectError
                Case Is = 1
                    Beep
                    MsgBox "There is already a customer with this name and" + vbCr + _
                           "first line of address.", vbExclamation
                    Exit Sub
                Case Is = 2
                    msg = Err.Description & vbCr _
                          & "Do you wish to continue?"
                    style = vbYesNo + vbInformation + vbDefaultButton2

                    response = MsgBox(msg, style)

                    If response = vbNo Then
                        Exit Sub
                    End If
                    On Error GoTo ssc_done_Click_Error
                    .Create Force:=True
                Case Else
                    If Err.Number <> 0 Then
                        ErrorSave
                        On Error GoTo ssc_done_Click_Error
                        ErrorRestore
                    Else
                        On Error GoTo ssc_done_Click_Error
                    End If
            End Select
        End With
        Set objCustomer = Nothing
        Unload Me
        Set frmCustEdit = Nothing
    Else
        update_customer
        If Not retry_update Then
            Unload Me
            Set frmCustEdit = Nothing
        End If
    End If
'
'********* Code Ends Here **************
'
   Exit Sub
'
ssc_done_Click_Error:
    DisplayError , "frmCustEdit.ssc_done_Click", vbExclamation
End Sub

Private Sub sstCustomer_Click(PreviousTab As Integer)
'***************************************
' Module/Form Name   : frmCustEdit
'
' Procedure Name     : sstCustomer_Click
'
' Purpose            :
'
' Date Created       : 31/12/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : PreviousTab - Integer
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo sstCustomer_Click_Error
'
'******** Code Starts Here *************
'
    Dim Ctl         As Control
    Dim ctlTop      As Control
    
    If sstCustomer.Tab = 1 And mMode = Edit Then
        If Not ContactsDisplayed Then
            DisplayContacts
            msfContacts.Refresh
            msfContacts.Redraw = True
        End If
    End If
    '
    '   Disable controls not on the current tab to stop them getting focus.
    '
    On Error Resume Next
    For Each Ctl In Controls
        Set ctlTop = TopLevelControl(Me, Ctl)
        If TypeOf ctlTop.Container Is SSTab Then
            'Not all controls have the TabStop property
            Ctl.TabStop = ctlTop.Left > 0
        End If
    Next Ctl
    On Error GoTo sstCustomer_Click_Error
'
'********* Code Ends Here **************
'
   Exit Sub
'
sstCustomer_Click_Error:
    DisplayError , "frmCustEdit.sstCustomer_Click", vbExclamation
End Sub

Private Sub txtServiceFee_GotFocus()
    HighLightText txtServiceFee
End Sub

Private Sub txtServiceFee_KeyPress(KeyAscii As Integer)
    KeyAscii = AllowCurrencyOnly(KeyAscii, txtServiceFee)
End Sub

Private Sub update_customer()
'***************************************
' Module/Form Name   : frmCustEdit
'
' Procedure Name     : update_customer
'
' Purpose            :
'
' Date Created       : 03/04/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 03/04/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo update_customer_Error
'
'******** Code Starts Here *************
'
    retry_update = False
    '
    '   Update the customer record.
    '
    With mobjCustomer
      .CustomerName = txt_customer_name.Text
      .Address1 = txt_address_line_1.Text
      .Address2 = txt_address_line_2.Text
      .Address3 = txt_address_line_3.Text
      .CountyOrState = txt_county_or_state
      .Country = txt_country.Text
      .PostCode = txt_post_code.Text
      .Telephone = txt_telephone_no.Text
      .Fax = txt_fax_no.Text
      .Information = txt_information.Text
      .CountryCode = txt_country_code.Text
      .VATNo = txt_vat_no.Text
      .PreferredDelivery = txt_preferred_delivery_method.Text
      .BusinessType = cboBusinessType.Text
      If goSystemConfig.SageLink Then
          If cboSageURN.ListIndex = 0 Then
              .SageURN = ""
          Else
              .SageURN = cboSageURN.Text
          End If
      End If
      .ServiceFeeRequired = chkServiceFeeRequired.Value
      .ServiceFee = CCur(txtServiceFee.Text)
      .update
    End With

    Set dyna_customer = Nothing
    Unload Me
    Set frmCustEdit = Nothing

    Exit Sub
    '
    '********* Code Ends Here **************
    '
    Exit Sub
    '
update_customer_Error:
    Select Case Err.Number
        Case Is = 3022
            MsgBox "Customer number " + Str(dyna_customer!customer_no) + " already exists", vbCritical
        Case Is = 3262
            MsgBox "'Next Customer Number' table is currently in use." + vbCr + "Please try again", vbInformation
        Case Is = 16389
            MsgBox "'Next Customer Number' table is currently in use." + vbCr + "Please try again", vbInformation
        Case Else
            DisplayError , "frmCustEdit.update_customer", vbExclamation
    End Select
    Set dyna_customer = Nothing
    retry_update = True
End Sub

Private Function valid_input() As Boolean
'***************************************
' Module/Form Name   : frmCustEdit
'
' Procedure Name     : valid_input
'
' Purpose            :
'
' Date Created       : 03/04/2002
'
' Author             : GARETH SAUNDERS
'
' Returns            : Boolean
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 03/04/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo valid_input_Error
'
'******** Code Starts Here *************
'
    valid_input = False

    If Trim(txt_customer_name.Text) = "" Then
        MsgBox "Please enter a Customer Name", vbExclamation
        txt_customer_name.SetFocus
        sstCustomer.Tab = 0
        Exit Function
    End If
    
    If cboBusinessType.Text = "" Then
        MsgBox "Please select a Business Type", vbExclamation
        cboBusinessType.SetFocus
        sstCustomer.Tab = 1
        Exit Function
    End If
    '
    '   Validate the Service Fee - Don't allow more than two decimal places.
    '
    If InStr(txtServiceFee.Text, ".") <> 0 Then
        If Len(txtServiceFee.Text) - InStr(txtServiceFee.Text, ".") > 2 Then
            sstCustomer.Tab = 3
            MsgBox "Enter a valid amount with no more than two decimal places.", vbExclamation
            txtServiceFee.SetFocus
            Exit Function
        End If
    End If
    '
    valid_input = True
'
'********* Code Ends Here **************
'
    Exit Function
        '
valid_input_Error:
    ErrorRaise "frmCustEdit.valid_input"
End Function

Public Sub ForceResize()
    gblnResizeMaxedWindows = False
    Form_Resize
    gblnResizeMaxedWindows = True
End Sub


