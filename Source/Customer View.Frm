VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm_cust_view 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Customer"
   ClientHeight    =   6600
   ClientLeft      =   1935
   ClientTop       =   1845
   ClientWidth     =   10125
   Icon            =   "Customer View.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6600
   ScaleWidth      =   10125
   Begin ImageWhere.SimpleGrid smgComments 
      Height          =   2000
      Left            =   5000
      TabIndex        =   27
      Top             =   1200
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   3519
      Columns         =   1
      KeyCol          =   0
   End
   Begin ImageWhere.SimpleGrid smgContacts 
      Height          =   2000
      Left            =   2880
      TabIndex        =   38
      Top             =   2000
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   3519
      Columns         =   1
      KeyCol          =   0
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   675
      Left            =   4680
      Picture         =   "Customer View.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   5880
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin TabDlg.SSTab sstCustomer 
      Height          =   5760
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9900
      _ExtentX        =   17463
      _ExtentY        =   10160
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      Tab             =   3
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "&General"
      TabPicture(0)   =   "Customer View.frx":0B04
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lab_county_or_state"
      Tab(0).Control(1)=   "lab_county_or_state_lab(1)"
      Tab(0).Control(2)=   "lab_telephoneno_lab(0)"
      Tab(0).Control(3)=   "lab_faxno_lab(0)"
      Tab(0).Control(4)=   "lab_country_lab(0)"
      Tab(0).Control(5)=   "lab_country"
      Tab(0).Control(6)=   "lab_faxno"
      Tab(0).Control(7)=   "lab_telephoneno"
      Tab(0).Control(8)=   "lab_post_code"
      Tab(0).Control(9)=   "lab_post_code_lab(9)"
      Tab(0).Control(10)=   "lab_customername_lab"
      Tab(0).Control(11)=   "lab_address_lab"
      Tab(0).Control(12)=   "lab_customername"
      Tab(0).Control(13)=   "lab_addressline1"
      Tab(0).Control(14)=   "lab_addressline2"
      Tab(0).Control(15)=   "lab_addressline3"
      Tab(0).Control(16)=   "txtDateAmended"
      Tab(0).Control(17)=   "lblDateAmended"
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "&Delivery Information"
      TabPicture(1)   =   "Customer View.frx":0B20
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lab_pref_delivery_method(0)"
      Tab(1).Control(1)=   "lab_slash"
      Tab(1).Control(2)=   "lab_vat_no"
      Tab(1).Control(3)=   "lab_preferred_delivery_method_lab"
      Tab(1).Control(4)=   "lab_business_type(1)"
      Tab(1).Control(5)=   "lab_business_type_lab"
      Tab(1).Control(6)=   "lab_information_lab"
      Tab(1).Control(7)=   "lab_vat_no_country_code_lab"
      Tab(1).Control(8)=   "lab_country_code"
      Tab(1).Control(9)=   "fraContacts"
      Tab(1).Control(10)=   "txtInformation"
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "&Statistics"
      TabPicture(2)   =   "Customer View.frx":0B3C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraTransparencies"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "&Private"
      TabPicture(3)   =   "Customer View.frx":0B58
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "&Billing"
      TabPicture(4)   =   "Customer View.frx":0B74
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lblSageOnHold"
      Tab(4).Control(1)=   "lblSageTermsAgreed"
      Tab(4).Control(2)=   "lblOnHold"
      Tab(4).Control(3)=   "lblTermsAndConditionsApproved"
      Tab(4).Control(4)=   "lblSageURNLabel"
      Tab(4).Control(5)=   "lblSageURN"
      Tab(4).ControlCount=   6
      Begin VB.Frame fraTransparencies 
         Caption         =   "Transparencies"
         Height          =   1455
         Left            =   -74760
         TabIndex        =   39
         Top             =   720
         Width           =   9375
         Begin VB.TextBox txt_pending_photographs 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   3060
            Locked          =   -1  'True
            TabIndex        =   43
            Top             =   900
            Width           =   735
         End
         Begin VB.TextBox txt_outstanding_photographs 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   3060
            Locked          =   -1  'True
            TabIndex        =   41
            Top             =   480
            Width           =   735
         End
         Begin VB.Label lab_no_pending_photos_lab 
            Caption         =   "Number Pending:"
            Height          =   255
            Left            =   1080
            TabIndex        =   42
            Top             =   900
            Width           =   1635
         End
         Begin VB.Label lab_no_outstanding_photos_lab 
            Caption         =   "Number Outstanding:"
            Height          =   255
            Left            =   1080
            TabIndex        =   40
            Top             =   480
            Width           =   1635
         End
      End
      Begin VB.TextBox txtInformation 
         BackColor       =   &H8000000F&
         Height          =   495
         Left            =   -72840
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   540
         Width           =   7395
      End
      Begin VB.Frame fraContacts 
         Caption         =   "Contacts"
         Height          =   3015
         Left            =   -74880
         TabIndex        =   37
         Top             =   1680
         Width           =   9615
      End
      Begin VB.Label lblDateAmended 
         Caption         =   "Date Amended:"
         DataField       =   "faxno"
         DataSource      =   "dat_customer"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -74700
         TabIndex        =   35
         Top             =   3780
         Width           =   1215
      End
      Begin VB.Label txtDateAmended 
         BorderStyle     =   1  'Fixed Single
         DataField       =   "faxno"
         DataSource      =   "dat_customer"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   -72840
         TabIndex        =   34
         Top             =   3780
         Width           =   1815
      End
      Begin VB.Label lblSageURN 
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -72600
         TabIndex        =   33
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label lblSageURNLabel 
         Caption         =   "Sage URN:"
         Height          =   255
         Left            =   -74400
         TabIndex        =   32
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblTermsAndConditionsApproved 
         Caption         =   "Terms Agreed:"
         Height          =   255
         Left            =   -74400
         TabIndex        =   31
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblOnHold 
         Caption         =   "On Hold:"
         Height          =   255
         Left            =   -74400
         TabIndex        =   30
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label lblSageTermsAgreed 
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -72600
         TabIndex        =   29
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label lblSageOnHold 
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -72600
         TabIndex        =   28
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label lab_addressline3 
         BorderStyle     =   1  'Fixed Single
         DataField       =   "addressline3"
         DataSource      =   "dat_customer"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   -72840
         TabIndex        =   26
         Top             =   1620
         Width           =   4005
      End
      Begin VB.Label lab_addressline2 
         BorderStyle     =   1  'Fixed Single
         DataField       =   "addressline2"
         DataSource      =   "dat_customer"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   -72840
         TabIndex        =   25
         Top             =   1260
         Width           =   4005
      End
      Begin VB.Label lab_addressline1 
         BorderStyle     =   1  'Fixed Single
         DataField       =   "addressline1"
         DataSource      =   "dat_customer"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   -72840
         TabIndex        =   24
         Top             =   900
         Width           =   4005
      End
      Begin VB.Label lab_customername 
         BorderStyle     =   1  'Fixed Single
         DataField       =   "customername"
         DataSource      =   "dat_customer"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   -72840
         TabIndex        =   23
         Top             =   540
         Width           =   4935
      End
      Begin VB.Label lab_address_lab 
         Caption         =   "Address:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74760
         TabIndex        =   22
         Top             =   900
         Width           =   1395
      End
      Begin VB.Label lab_customername_lab 
         Caption         =   "Customer Name:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74760
         TabIndex        =   21
         Top             =   540
         Width           =   1395
      End
      Begin VB.Label lab_post_code_lab 
         Caption         =   "Post Code:"
         Height          =   255
         Index           =   9
         Left            =   -74700
         TabIndex        =   20
         Top             =   2700
         Width           =   1815
      End
      Begin VB.Label lab_post_code 
         BorderStyle     =   1  'Fixed Single
         DataField       =   "post_code"
         DataSource      =   "dat_customer"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   -72840
         TabIndex        =   19
         Top             =   2700
         Width           =   1095
      End
      Begin VB.Label lab_telephoneno 
         BorderStyle     =   1  'Fixed Single
         DataField       =   "telephoneno"
         DataSource      =   "dat_customer"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   -72840
         TabIndex        =   18
         Top             =   3060
         Width           =   1815
      End
      Begin VB.Label lab_faxno 
         BorderStyle     =   1  'Fixed Single
         DataField       =   "faxno"
         DataSource      =   "dat_customer"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   -72840
         TabIndex        =   17
         Top             =   3420
         Width           =   1815
      End
      Begin VB.Label lab_country 
         BorderStyle     =   1  'Fixed Single
         DataField       =   "post_code"
         DataSource      =   "dat_customer"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   -72840
         TabIndex        =   16
         Top             =   2340
         Width           =   2295
      End
      Begin VB.Label lab_country_lab 
         Caption         =   "Country:"
         Height          =   255
         Index           =   0
         Left            =   -74700
         TabIndex        =   15
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label lab_faxno_lab 
         Caption         =   "Fax No:"
         Height          =   255
         Index           =   0
         Left            =   -74700
         TabIndex        =   14
         Top             =   3420
         Width           =   675
      End
      Begin VB.Label lab_telephoneno_lab 
         Caption         =   "Telephone No:"
         Height          =   255
         Index           =   0
         Left            =   -74700
         TabIndex        =   13
         Top             =   3060
         Width           =   1815
      End
      Begin VB.Label lab_county_or_state_lab 
         Caption         =   "County/State:"
         Height          =   255
         Index           =   1
         Left            =   -74700
         TabIndex        =   12
         Top             =   1980
         Width           =   1815
      End
      Begin VB.Label lab_county_or_state 
         BorderStyle     =   1  'Fixed Single
         DataField       =   "addressline3"
         DataSource      =   "dat_customer"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   -72840
         TabIndex        =   11
         Top             =   1980
         Width           =   2295
      End
      Begin VB.Label lab_country_code 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   -72840
         TabIndex        =   10
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label lab_vat_no_country_code_lab 
         Caption         =   "Country Code/VAT No.:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74760
         TabIndex        =   9
         Top             =   1200
         Width           =   1755
      End
      Begin VB.Label lab_information_lab 
         Caption         =   "Information:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74760
         TabIndex        =   8
         Top             =   540
         Width           =   1635
      End
      Begin VB.Label lab_business_type_lab 
         Caption         =   "Business Type:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74760
         TabIndex        =   6
         Top             =   5280
         Width           =   1815
      End
      Begin VB.Label lab_business_type 
         BorderStyle     =   1  'Fixed Single
         DataField       =   "no_outstanding_photos"
         DataSource      =   "dat_customer"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   -72660
         TabIndex        =   5
         Top             =   5280
         Width           =   3435
      End
      Begin VB.Label lab_preferred_delivery_method_lab 
         Caption         =   "Preferred Delivery Method:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74760
         TabIndex        =   4
         Top             =   4860
         Width           =   2115
      End
      Begin VB.Label lab_vat_no 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   -71940
         TabIndex        =   3
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label lab_slash 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -71940
         TabIndex        =   2
         Top             =   1200
         Width           =   135
      End
      Begin VB.Label lab_pref_delivery_method 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   -72660
         TabIndex        =   1
         Top             =   4860
         Width           =   2595
      End
   End
End
Attribute VB_Name = "frm_cust_view"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public snap_customer As DAO.Recordset
Private SQL As String
Public cancel_list As Boolean
Dim mod_pic_moved(3) As String
Dim mod_statistics_loaded As Boolean
Private mobjCustomer As Customer2

Private Sub display_statistics()
'***************************************
' Module/Form Name   : frm_cust_view
'
' Procedure Name     : display_statistics
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
On Error GoTo display_statistics_Error
'
'******** Code Starts Here *************
'
    
    txt_outstanding_photographs = search_result.customers_outstanding_photographs(mobjCustomer.CustomerNo)
    txt_pending_photographs = search_result.customers_pending_photographs(mobjCustomer.CustomerNo)

'
'********* Code Ends Here **************
'
   Exit Sub
'
display_statistics_Error:
    ErrorRaise "frm_cust_view.display_statistics"
End Sub

Public Sub load_customer(view_customer_name As String)
'***************************************
' Module/Form Name   : frm_cust_view
'
' Procedure Name     : load_customer
'
' Purpose            :
'
' Date Created       : 06/12/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : view_customer_name - String
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo load_customer_Error
'
'******** Code Starts Here *************
'
    Set mobjCustomer = New Customer2
    On Error Resume Next
    mobjCustomer.Locate view_customer_name
    If Err.Number = vbObjectError + 7 Then  'Cancelled
        Exit Sub
    ElseIf Err.Number <> 0 Then
        ErrorSave
        On Error GoTo load_customer_Error
        ErrorRestore
    Else
        On Error GoTo load_customer_Error
    End If
    With mobjCustomer
        .Contacts.Refresh
        .Comments.Refresh
    End With
    display_customer
'
'********* Code Ends Here **************
'
   Exit Sub
'
load_customer_Error:
    ErrorRaise "frm_cust_view.load_customer"
End Sub

Private Sub display_customer()
'***************************************
' Module/Form Name   : frm_cust_view
'
' Procedure Name     : display_customer
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
On Error GoTo display_customer_Error
'
'******** Code Starts Here *************
'
    Dim oCustomer As Customer2
    
    With mobjCustomer
        lab_customername.Caption = Replace(.CustomerName, "&", "&&")
        lab_addressline1 = Replace(.Address1, "&", "&&")
        lab_addressline2 = Replace(.Address2, "&", "&&")
        lab_addressline3 = Replace(.Address3, "&", "&&")
        lab_post_code = Replace(.PostCode, "&", "&&")
        lab_telephoneno = Replace(.Telephone, "&", "&&")
        lab_county_or_state = Replace(.CountyOrState, "&", "&&")
        lab_country = Replace(.Country, "&", "&&")
        lab_faxno = Replace(.Fax, "&", "&&")
        lab_country_code = Replace(.CountryCode, "&", "&&")
        lab_vat_no = Replace(.VATNo, "&", "&&")
        If RTrim(lab_country_code) = "" Or RTrim(lab_vat_no) = "" Then
            lab_slash = ""
        Else
            lab_slash = "/"
        End If
        txtInformation = Replace(.Information, "&", "&&")
        lab_pref_delivery_method(0) = Replace(.PreferredDelivery, "&", "&&")
        lab_business_type(1) = Replace(.BusinessType, "&", "&&")
        If goSystemConfig.SageLink Then
            If .SageURN = "" Then
                lblSageURN.Caption = "<Not Set>"
            Else
                lblSageURN.Caption = .SageURN
            End If
            lblSageTermsAgreed = IIf(.TermsAndConditionsApproved, "Yes", "No")
            lblSageOnHold = IIf(.OnHold, "Yes", "No")
        End If
        txtDateAmended.Caption = Format(.DateAmended, "dd/mm/yyyy")
    End With
    Me.Show
    Exit Sub
'
'********* Code Ends Here **************
'
display_customer_Error:
    ErrorRaise "frm_cust_view.display_customer"
End Sub

Private Sub Form_Load()
'***************************************
' Module/Form Name   : frm_cust_view
'
' Procedure Name     : Form_Load
'
' Purpose            :
'
' Date Created       : 19/05/2002
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
    mod_statistics_loaded = False
    com_position_form Me
    '
    '   If there is no Sage Link don't show the Billing Tab.
    '
    If goSystemConfig.SageLink Then
        sstCustomer.TabVisible(4) = True
    Else
        sstCustomer.TabVisible(4) = False
    End If
    '
    '   Initialize the comments grid.
    '
    Set smgComments.Container = sstCustomer
    With smgComments
        .Height = 4935
        .Left = 240
        .Top = 600
        .Width = 9375
    End With
    
    With smgComments
        .Columns = 5
        .KeyCol = 4
        .Column(1).Header = "Who"
        .Column(2).Header = "Comment"
        .Column(3).Header = "Date"
        .Column(3).Width = 1800
        .Column(4).Width = 0
        .Column(5).Width = 0
        .Column(2).Width = .Width - .Column(1).Width - .Column(3).Width - 354
    End With
    '
    '   Initialize the contacts grid.
    '
    Set smgContacts.Container = fraContacts
    With smgContacts
        .Height = 2655
        .Left = 120
        .Top = 600
        .Width = 9375
    End With
    
    With smgContacts
        .Redraw = False
        .Columns = 8
        .Column(1).Header = "Web User"
        .ColumnProperty(1).PropertyType = smgPropTypeBooleanGraphic
        .Column(2).Header = "Mail"
        .ColumnProperty(2).PropertyType = smgPropTypeBooleanGraphic
        .Column(3).Header = "Name"
        .Column(4).Header = "Position"
        .Column(5).Header = "Phone"
        .Column(6).Header = "Fax"
        .Column(7).Header = "Email"
        .Column(8).Header = "Comments"
        .Column(1).Width = 500
        .Column(2).Width = 400
        .Column(3).Width = 1800
        .Column(5).Width = 1200
        .Column(6).Width = 1200
        .Column(7).Width = 1200
        .Column(4).Width = 1200
        .Column(8).Width = .Width - .Column(1).Width - _
                                    .Column(2).Width - _
                                    .Column(3).Width - _
                                    .Column(5).Width - _
                                    .Column(6).Width - _
                                    .Column(7).Width - _
                                    .Column(4).Width - 364
  
        .Redraw = True
    End With

    '
    '   Refresh Grids.
    '
    RefreshContacts
    RefreshComments
    
    sstCustomer.Tab = 0
    
'
'********* Code Ends Here **************
'
   Exit Sub
'
Form_Load_Error:
    DisplayError , "frm_cust_view.Form_Load", vbExclamation
End Sub

Private Sub cmdExit_Click()
    Unload Me
    Set frm_cust_view = Nothing

End Sub

Private Sub RefreshContacts()
'***************************************
' Module/Form Name   : frm_cust_view
'
' Procedure Name     : RefreshContacts
'
' Purpose            :
'
' Date Created       : 19/05/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo RefreshContacts_Error
'
'******** Code Starts Here *************
'
    Dim objContact As Contact
    
    With smgContacts
        .Redraw = False
        .Clear
        For Each objContact In mobjCustomer.Contacts
            .AddRow False, _
                    objContact.WebUser, _
                    objContact.Mail, _
                    objContact.Name, _
                    objContact.Position, _
                    objContact.Phone, _
                    objContact.Fax, _
                    objContact.Email, _
                    objContact.Comments
        Next objContact
        .ResizeRows
        .Redraw = True
    End With

'
'********* Code Ends Here **************
'
   Exit Sub
'
RefreshContacts_Error:
    ErrorRaise "frm_cust_view.RefreshContacts"
End Sub

Private Sub RefreshComments()
    '***************************************
    ' Module/Form Name   : frm_cust_view
    '
    ' Procedure Name     : RefreshComments
    '
    ' Purpose            :
    '
    ' Date Created       : 21/04/2002
    '
    ' Author             : GARETH SAUNDERS
    '
    ' Amendment History  : Date       Author    Description
    '                    : --------------------------------
    '                    : 21/04/2002 GARETH SAUNDERS
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
              Format(objComment.DateWritten, "yyyymmddhhmm")
  Next objComment
  .ResizeRows
  .Column(5).Sorted = smgDescending
  .Redraw = True
End With
    '
    '********* Code Ends Here **************
    '
Exit Sub
    '
RefreshComments_Error:
    ErrorRaise "frm_cust_view.RefreshComments"
End Sub

Private Sub sstCustomer_Click(PreviousTab As Integer)
'***************************************
' Module/Form Name   : frm_cust_view
'
' Procedure Name     : sstCustomer_Click
'
' Purpose            :
'
' Date Created       : 06/12/2002
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
    If sstCustomer.Tab = 2 Then
        If mod_statistics_loaded = False Then
            Screen.MousePointer = vbHourglass
            display_statistics
            Screen.MousePointer = vbDefault
            mod_statistics_loaded = True
        End If
    End If

'
'********* Code Ends Here **************
'
   Exit Sub
'
sstCustomer_Click_Error:
    DisplayError , "frm_cust_view.sstCustomer_Click", vbExclamation
End Sub

