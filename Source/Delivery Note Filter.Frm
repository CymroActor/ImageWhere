VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm_delivery_note 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delivery Note"
   ClientHeight    =   4965
   ClientLeft      =   6045
   ClientTop       =   1710
   ClientWidth     =   4695
   ControlBox      =   0   'False
   Icon            =   "Delivery Note Filter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4965
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkOpenDeliveryNotesOnly 
      Caption         =   "Open Trans'pncy D/Ns Only"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   4440
      Width           =   2415
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   555
      Left            =   3840
      Picture         =   "Delivery Note Filter.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4320
      UseMaskColor    =   -1  'True
      Width           =   795
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&OK"
      Height          =   555
      Left            =   2880
      Picture         =   "Delivery Note Filter.frx":089C
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4320
      UseMaskColor    =   -1  'True
      Width           =   795
   End
   Begin TabDlg.SSTab sst_delivery_note 
      Height          =   4215
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   7435
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "S&elect"
      TabPicture(0)   =   "Delivery Note Filter.frx":0E16
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fra_delivery_note_no"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Filter"
      TabPicture(1)   =   "Delivery Note Filter.frx":0E32
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fra_delivery_notes_for"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fra_delivery_note_contains"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame fra_delivery_note_no 
         Caption         =   "Delivery Note:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1815
         Left            =   -74940
         TabIndex        =   16
         Top             =   360
         Width           =   4395
         Begin VB.TextBox txt_delivery_note_no 
            Height          =   315
            Left            =   1560
            MaxLength       =   8
            TabIndex        =   1
            Top             =   720
            Width           =   915
         End
         Begin VB.Label lab_delivery_note_no 
            Caption         =   "&Number:"
            Height          =   255
            Left            =   780
            TabIndex        =   2
            Top             =   720
            Width           =   615
         End
         Begin VB.Image Img_delivery_note_no 
            Height          =   480
            Left            =   3060
            Picture         =   "Delivery Note Filter.frx":0E4E
            Top             =   660
            Width           =   480
         End
      End
      Begin VB.Frame fra_delivery_note_contains 
         Caption         =   "Delivery Note Contains:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   1575
         Left            =   120
         TabIndex        =   11
         Top             =   2580
         Width           =   4335
         Begin VB.TextBox txt_photograph_no 
            Height          =   315
            Left            =   1560
            MaxLength       =   8
            TabIndex        =   14
            Top             =   960
            Width           =   915
         End
         Begin VB.OptionButton opt_any_photograph 
            Caption         =   "A&ny photograph:"
            Height          =   315
            Left            =   120
            TabIndex        =   12
            Top             =   540
            Value           =   -1  'True
            Width           =   1515
         End
         Begin VB.OptionButton opt_photograph_no 
            Caption         =   "&Photograph no:"
            Height          =   315
            Left            =   120
            TabIndex        =   13
            Top             =   900
            Width           =   1395
         End
         Begin VB.Image img_photograph 
            Height          =   480
            Left            =   2580
            Picture         =   "Delivery Note Filter.frx":1158
            Top             =   300
            Width           =   480
         End
      End
      Begin VB.Frame fra_delivery_notes_for 
         Caption         =   "Delivery Notes For Customer:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   2175
         Left            =   60
         TabIndex        =   3
         Top             =   360
         Width           =   4395
         Begin VB.TextBox txt_customer_name 
            BackColor       =   &H8000000F&
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   1560
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   7
            Top             =   1200
            Width           =   2715
         End
         Begin VB.TextBox txt_first_line_of_address 
            BackColor       =   &H8000000F&
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   1560
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   8
            Top             =   1560
            Width           =   2715
         End
         Begin VB.CommandButton com_select_customer 
            Caption         =   "&SELECT..."
            Height          =   435
            Left            =   1560
            TabIndex        =   6
            Top             =   660
            Width           =   1395
         End
         Begin VB.OptionButton opt_all_customers 
            Caption         =   "&All Customers:"
            Height          =   315
            Left            =   120
            TabIndex        =   4
            Top             =   300
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton opt_customer 
            Caption         =   "C&ustomer:"
            Height          =   315
            Left            =   120
            TabIndex        =   5
            Top             =   660
            Width           =   1095
         End
         Begin VB.Label lab_customer_name 
            Caption         =   "Name"
            Height          =   255
            Left            =   420
            TabIndex        =   10
            Top             =   1200
            Width           =   555
         End
         Begin VB.Label lab_first_line_of_address 
            Caption         =   "First line of address"
            Height          =   375
            Left            =   420
            TabIndex        =   9
            Top             =   1560
            Width           =   855
         End
         Begin VB.Image img_customer 
            Height          =   480
            Left            =   3600
            Picture         =   "Delivery Note Filter.frx":1462
            Top             =   240
            Width           =   480
         End
      End
   End
End
Attribute VB_Name = "frm_delivery_note"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mod_customer_filter As Integer
Private mod_photograph_filter As String
Private mod_delivery_note_no As Long
'
'   Used for selecting the customer.
'
Public snap_customer As DAO.Recordset
Public cancel_list As Boolean
Public mstrCustomerSelect As String
Private moCustomer As Customer2
Public WithEvents fInputCustomer As frm_input_customer
Attribute fInputCustomer.VB_VarHelpID = -1

Private Function process_filter_tab()
'***************************************
' Module/Form Name   : frm_delivery_note
'
' Procedure Name     : process_filter_tab
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
On Error GoTo process_filter_tab_Error
'
'******** Code Starts Here *************
'
    process_filter_tab = False
    
    If opt_customer.Value = True And Len(RTrim(txt_customer_name)) = 0 Then
        MsgBox "You've chosen to filter on customer. Please select a customer.", vbExclamation
        com_select_customer.SetFocus
        Exit Function
    End If

    If opt_photograph_no.Value = True And Val(txt_photograph_no) = 0 Then
        MsgBox "You've chosen to filter on photograph. Please enter a non-zero photograph number.", vbExclamation
        txt_photograph_no.SetFocus
        Exit Function
    End If
    
    If opt_photograph_no.Value = True And Len(RTrim(txt_photograph_no)) = 0 Then
        MsgBox "You've chosen to filter on photograph. Please enter a photograph number.", vbExclamation
        txt_photograph_no.SetFocus
        Exit Function
    End If
'
'   if 'All Customers' and 'Any Photographs' are chosen, ask whether the user wishes to continue.
'
    If opt_all_customers.Value = True And opt_any_photograph.Value = True Then
        msg = "This will potentially involve a large retrieval of records." + vbCr _
            & "Do you wish to continue?"
        style = vbYesNo + vbInformation + vbDefaultButton2
        response = MsgBox(msg, style)
        If response = vbNo Then
            Exit Function
        End If
    End If
'
'   Remove a trailing full-stop from the photograph filter if necessary.
'
    mod_photograph_filter = LTrim(RTrim(mod_photograph_filter))
    If Len(mod_photograph_filter) > 0 Then
        If Mid(mod_photograph_filter, Len(mod_photograph_filter), 1) = "." Then
            mod_photograph_filter = Left(mod_photograph_filter, Len(mod_photograph_filter) - 1)
        End If
    End If
'
'   Initialise the delivery note filter.
'
    mod_delivery_note_no = 0
    process_filter_tab = True
'
'********* Code Ends Here **************
'
   Exit Function
'
process_filter_tab_Error:
    ErrorRaise "frm_delivery_note.process_filter_tab"
End Function

Private Sub display_customer()
'***************************************
' Module/Form Name   : frm_delivery_note
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
    With moCustomer
        txt_customer_name = .CustomerName
        txt_first_line_of_address = .Address1
        mod_customer_filter = .CustomerNo
    End With
'
'********* Code Ends Here **************
'
   Exit Sub
'
display_customer_Error:
    ErrorRaise "frm_delivery_note.display_customer"
End Sub

Public Sub load_customer(customer_name As String, customer As Customer2, ByRef retry As Boolean)
'***************************************
' Module/Form Name   : frm_delivery_note
'
' Procedure Name     : load_customer
'
' Purpose            :
'
' Date Created       : 06/12/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : customer_name - String
'                    : customer - Customer2
'                    : retry - Boolean
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
'
'********* Code Ends Here **************
'
   Exit Sub
'
load_customer_Error:
    ErrorRaise "frm_delivery_note.load_customer"
End Sub

Public Sub load_delivery_notes(customer_filter As Integer, _
                               photograph_filter As String, _
                               delivery_note_filter As Long)
'***************************************
' Module/Form Name   : frm_delivery_note
'
' Procedure Name     : load_delivery_notes
'
' Purpose            :
'
' Date Created       : 19/05/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : customer_filter - Integer
'                    : photograph_filter - String
'                    : delivery_note_filter - Integer
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 19/05/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo load_delivery_notes_Error
'
'******** Code Starts Here *************
'
    '
    '    Set up Module Filters
    '
    mod_delivery_note_no = delivery_note_filter
    '
    '   Get the last used Delivery Note No.
    '
    If delivery_note_filter = 0 Then txt_delivery_note_no.Text = GetSetting(App.Title, "Delivery Note", "DeliveryNoteNo", "")
    
    If customer_filter = 0 Then
        mod_customer_filter = GetSetting(App.Title, "Delivery Note", "Customer", 0)
    Else
        mod_customer_filter = customer_filter
    End If
        
    If photograph_filter = 0 Then
        mod_photograph_filter = Val(GetSetting(App.Title, "Delivery Note", "PhotographNo", 0))
    Else
        mod_photograph_filter = photograph_filter
    End If
    
    chkOpenDeliveryNotesOnly.Value = GetSetting(App.Title, "Delivery Note", "OpenDeliveryNotesOnly", 1)
    
    If Not mod_customer_filter = 0 Then
        Set moCustomer = Nothing
        Set moCustomer = New Customer2
        On Error Resume Next
        moCustomer.Read mod_customer_filter
        If Err.Number <> vbObjectError + 8 And _
           Err.Number <> 0 Then
            ErrorSave
            On Error GoTo load_delivery_notes_Error
            ErrorRestore
        Else
            On Error GoTo load_delivery_notes_Error
        End If
        display_customer
    End If
    
    If delivery_note_filter = 0 Then
        GoTo CustomerPhotoFiltering
    End If
    
DeliveryNoteNoFiltering:
    '
    '    sst_delivery_note.Tab = 1
    '
    sst_delivery_note.Tab = 0
    txt_delivery_note_no = Str(LTrim(delivery_note_filter))
    GoTo LoadEnd
    
CustomerPhotoFiltering:
    
    sst_delivery_note.Tab = 1
    
    If Not mod_photograph_filter = "0" Then
        txt_photograph_no.Text = mod_photograph_filter
    End If
    
LoadEnd:
    
    Me.Show 1
'
'********* Code Ends Here **************
'
    Exit Sub
    '
load_delivery_notes_Error:
    ErrorRaise "frm_delivery_note.load_delivery_notes"
End Sub


Private Sub process_select_tab()
'***************************************
' Module/Form Name   : frm_delivery_note
'
' Procedure Name     : process_select_tab
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
On Error GoTo process_select_tab_Error
'
'******** Code Starts Here *************
'
'
'   Initialise the customer and photograph filters.
'
''    mod_customer_filter = 0
''    mod_photograph_filter = 0
'
'   Set the delivery note no.
'
    mod_delivery_note_no = CLng(txt_delivery_note_no)
    
'
'********* Code Ends Here **************
'
   Exit Sub
'
process_select_tab_Error:
    ErrorRaise "frm_delivery_note.process_select_tab"
End Sub

Private Sub com_select_customer_Click()
'***************************************
' Module/Form Name   : frm_delivery_note
'
' Procedure Name     : com_select_customer_Click
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
On Error GoTo com_select_customer_Click_Error
'
'******** Code Starts Here *************
'
    Set fInputCustomer = New frm_input_customer
    With fInputCustomer
        .Caption = "Select Customer"
        .Show 1
    End With
    Set fInputCustomer = Nothing
'
'********* Code Ends Here **************
'
   Exit Sub
'
com_select_customer_Click_Error:
    DisplayError , "frm_delivery_note.com_select_customer_Click", vbExclamation
End Sub

Private Sub fInputCustomer_CustomerFound(oCustomer As Customer2, blnAbort As Boolean)
'***************************************
' Module/Form Name   : frm_delivery_note
'
' Procedure Name     : fInputCustomer_CustomerFound
'
' Purpose            :
'
' Date Created       : 06/12/2002
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
    Set moCustomer = oCustomer
    display_customer
'
'********* Code Ends Here **************
'
   Exit Sub
'
fInputCustomer_CustomerFound_Error:
    ErrorRaise "frm_delivery_note.fInputCustomer_CustomerFound"
End Sub

Private Sub Form_Activate()
'***************************************
' Module/Form Name   : frm_delivery_note
'
' Procedure Name     : Form_Activate
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
On Error GoTo Form_Activate_Error
'
'******** Code Starts Here *************
'
    If sst_delivery_note.Tab = 0 Then
        txt_delivery_note_no.SetFocus
    End If
  
'
'********* Code Ends Here **************
'
   Exit Sub
'
Form_Activate_Error:
    DisplayError , "frm_delivery_note.Form_Activate", vbExclamation
End Sub

Private Sub opt_all_customers_Click()
'***************************************
' Module/Form Name   : frm_delivery_note
'
' Procedure Name     : opt_all_customers_Click
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
On Error GoTo opt_all_customers_Click_Error
'
'******** Code Starts Here *************
'
    txt_customer_name.Text = ""
    txt_first_line_of_address = ""
    mod_customer_filter = 0
'
'********* Code Ends Here **************
'
   Exit Sub
'
opt_all_customers_Click_Error:
    DisplayError , "frm_delivery_note.opt_all_customers_Click", vbExclamation
End Sub

Private Sub opt_any_photograph_Click()
    txt_photograph_no.Text = ""
    mod_photograph_filter = 0
End Sub

Private Sub opt_customer_Click()
    '***************************************
    ' Module/Form Name   : frm_delivery_note
    '
    ' Procedure Name     : opt_customer_Click
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
    On Error GoTo opt_customer_Click_Error
    '
    '******** Code Starts Here *************
    '
    On Error Resume Next
    If Screen.ActiveForm.Name = "frm_delivery_note_filter" Then
''        On Error GoTo opt_customer_Click_Error
        com_select_customer.SetFocus
    End If
    '
    '********* Code Ends Here **************
    '
    Exit Sub
    '
opt_customer_Click_Error:
    DisplayError , "frm_delivery_note.opt_customer_Click", vbExclamation
End Sub

Private Sub opt_photograph_no_Click()
'***************************************
' Module/Form Name   : frm_delivery_note
'
' Procedure Name     : opt_photograph_no_Click
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
On Error GoTo opt_photograph_no_Click_Error
'
'******** Code Starts Here *************
'
    If Screen.ActiveForm.Name = "frm_delivery_note_filter" Then
        txt_photograph_no.SetFocus
    End If
'
'********* Code Ends Here **************
'
   Exit Sub
'
opt_photograph_no_Click_Error:
    DisplayError , "frm_delivery_note.opt_photograph_no_Click", vbExclamation
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    Set frm_delivery_note = Nothing
End Sub

Private Sub cmdSelect_Click()
'***************************************
' Module/Form Name   : frm_delivery_note
'
' Procedure Name     : cmdSelect_Click
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
On Error GoTo cmdSelect_Click_Error
'
'******** Code Starts Here *************
'
    Dim response As Integer
    Dim msg As String
    Dim style As String
    Dim SQL As String
        
    If sst_delivery_note.Tab = 0 Then
        process_select_tab
    Else
        If Not process_filter_tab Then
            Exit Sub
        End If
    End If
    
        
    frm_delivery_note_maint.CurrentDNoteNo = ""
    LockWindow frm_delivery_note_maint.hWnd
    frm_delivery_note_maint.list_delivery_notes mod_delivery_note_no, _
                                                mod_customer_filter, _
                                                mod_photograph_filter, _
                                                abort, _
                                                False, _
                                                chkOpenDeliveryNotesOnly.Value, _
                                                pblnFreezeWindow:=False
    If abort Then
        Exit Sub
    End If
  
    SaveSetting App.Title, "Delivery Note", "Customer", mod_customer_filter
    SaveSetting App.Title, "Delivery Note", "OpenDeliveryNotesOnly", CStr(chkOpenDeliveryNotesOnly.Value)
    SaveSetting App.Title, "Delivery Note", "DeliveryNoteNo", txt_delivery_note_no.Text
    SaveSetting App.Title, "Delivery Note", "PhotographNo", txt_photograph_no.Text
    
    Unload Me
    Set frm_delivery_note = Nothing
    
    frm_delivery_note_maint.Display
    UnlockWindow
'
'********* Code Ends Here **************
'
   Exit Sub
'
cmdSelect_Click_Error:
    DisplayError , "frm_delivery_note.cmdSelect_Click", vbExclamation
End Sub

Private Sub sst_delivery_note_Click(PreviousTab As Integer)
'***************************************
' Module/Form Name   : frm_delivery_note
'
' Procedure Name     : sst_delivery_note_Click
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
On Error GoTo sst_delivery_note_Click_Error
'
'******** Code Starts Here *************
'
    If sst_delivery_note.Tab = 0 Then
        fra_delivery_notes_for.Enabled = False
        fra_delivery_note_contains.Enabled = False
        com_select_customer.TabStop = False
        txt_delivery_note_no.TabStop = True
        If Screen.ActiveForm.Name = Me.Name Then
            txt_delivery_note_no.SetFocus
        End If
    Else
        fra_delivery_notes_for.Enabled = True
        fra_delivery_note_contains.Enabled = True
        com_select_customer.TabStop = True
        txt_delivery_note_no.TabStop = False
        'SendKeys "{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}"
    End If
    
'
'********* Code Ends Here **************
'
   Exit Sub
'
sst_delivery_note_Click_Error:
    DisplayError , "frm_delivery_note.sst_delivery_note_Click", vbExclamation
End Sub

Private Sub txt_customer_name_Change()
'***************************************
' Module/Form Name   : frm_delivery_note
'
' Procedure Name     : txt_customer_name_Change
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
On Error GoTo txt_customer_name_Change_Error
'
'******** Code Starts Here *************
'
    If Not Len(RTrim(txt_customer_name)) = 0 Then
        On Error Resume Next
        opt_customer.Value = True
    End If
'
'********* Code Ends Here **************
'
   Exit Sub
'
txt_customer_name_Change_Error:
    DisplayError , "frm_delivery_note.txt_customer_name_Change", vbExclamation
End Sub

Private Sub txt_delivery_note_no_GotFocus()
    txt_delivery_note_no.SelStart = 0
    txt_delivery_note_no.SelLength = Len(txt_delivery_note_no.Text)
End Sub

Private Sub txt_delivery_note_no_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And _
        Not KeyAscii = 8 Then
        KeyAscii = 0
        Beep
    End If
End Sub

Private Sub txt_photograph_no_Change()
'***************************************
' Module/Form Name   : frm_delivery_note
'
' Procedure Name     : txt_photograph_no_Change
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
On Error GoTo txt_photograph_no_Change_Error
'
'******** Code Starts Here *************
'
    If Len(RTrim(txt_photograph_no.Text)) = 0 Then
        opt_any_photograph.Value = True
    Else
        opt_photograph_no.Value = True
        mod_photograph_filter = txt_photograph_no.Text
    End If
  
'
'********* Code Ends Here **************
'
   Exit Sub
'
txt_photograph_no_Change_Error:
    DisplayError , "frm_delivery_note.txt_photograph_no_Change", vbExclamation
End Sub


Private Sub txt_photograph_no_GotFocus()
    txt_photograph_no.SelStart = 0
    txt_photograph_no.SelLength = Len(txt_photograph_no.Text)
End Sub


Private Sub txt_photograph_no_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And _
        Not KeyAscii = 8 And _
        Not KeyAscii = 46 Then
        KeyAscii = 0
        Beep
    End If
End Sub


Private Sub txt_photograph_no_KeyUp(keycode As Integer, Shift As Integer)
'***************************************
' Module/Form Name   : frm_delivery_note
'
' Procedure Name     : txt_photograph_no_KeyUp
'
' Purpose            :
'
' Date Created       : 06/12/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : keycode - Integer
'                    : Shift - Integer
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo txt_photograph_no_KeyUp_Error
'
'******** Code Starts Here *************
'
   
    Static photograph_no_save As String
    
    If Len(LTrim(txt_photograph_no)) = 0 Then
         photograph_no_save = ""
    Else
        If Not IsNumeric(txt_photograph_no.Text) Then
            txt_photograph_no.Text = photograph_no_save
            txt_photograph_no.SelStart = Len(txt_photograph_no.Text)
            Beep
        Else
            photograph_no_save = txt_photograph_no.Text
        End If
    End If

'
'********* Code Ends Here **************
'
   Exit Sub
'
txt_photograph_no_KeyUp_Error:
    DisplayError , "frm_delivery_note.txt_photograph_no_KeyUp", vbExclamation
End Sub


