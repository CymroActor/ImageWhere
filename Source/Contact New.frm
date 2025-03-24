VERSION 5.00
Begin VB.Form frmNewContact 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Contact for Customer "
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4500
   Icon            =   "Contact New.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboPreferredPhotoType 
      Height          =   315
      Left            =   1620
      TabIndex        =   15
      Text            =   "Combo1"
      Top             =   4200
      Width           =   2295
   End
   Begin VB.CheckBox chkWebUser 
      Alignment       =   1  'Right Justify
      Caption         =   "&Web User"
      Height          =   315
      Left            =   2760
      TabIndex        =   13
      Top             =   3780
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   555
      Left            =   3600
      Picture         =   "Contact New.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5220
      UseMaskColor    =   -1  'True
      Width           =   795
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   555
      Left            =   2640
      Picture         =   "Contact New.frx":0C84
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5220
      UseMaskColor    =   -1  'True
      Width           =   795
   End
   Begin VB.CheckBox chkMailingList 
      Alignment       =   1  'Right Justify
      Caption         =   "&Mailing List?   "
      Height          =   315
      Left            =   200
      TabIndex        =   12
      Top             =   3780
      Width           =   1620
   End
   Begin VB.TextBox txtEmail 
      Height          =   315
      Left            =   1620
      MaxLength       =   50
      TabIndex        =   11
      Top             =   3360
      Width           =   2715
   End
   Begin VB.TextBox txtFax 
      Height          =   315
      Left            =   1620
      MaxLength       =   20
      TabIndex        =   9
      Top             =   2940
      Width           =   2715
   End
   Begin VB.TextBox txtPhone 
      Height          =   315
      Left            =   1620
      MaxLength       =   20
      TabIndex        =   7
      Top             =   2520
      Width           =   2715
   End
   Begin VB.TextBox txtComments 
      Height          =   1335
      Left            =   1620
      MaxLength       =   200
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   1080
      Width           =   2715
   End
   Begin VB.TextBox txtPosition 
      Height          =   315
      Left            =   1620
      MaxLength       =   20
      TabIndex        =   3
      Top             =   660
      Width           =   2715
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   1620
      MaxLength       =   20
      TabIndex        =   1
      Top             =   240
      Width           =   2715
   End
   Begin VB.Label lblPreferredPhotoType 
      Caption         =   "Preferred Photo Type:"
      Height          =   435
      Left            =   240
      TabIndex        =   14
      Top             =   4200
      Width           =   1275
   End
   Begin VB.Label lblDateAmended 
      Caption         =   "Date Amended:"
      DataField       =   "faxno"
      DataSource      =   "dat_customer"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label txtDateAmended 
      DataField       =   "faxno"
      DataSource      =   "dat_customer"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   1680
      TabIndex        =   17
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label labEmail 
      Caption         =   "Emai&l:"
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   3420
      Width           =   1155
   End
   Begin VB.Label labFax 
      Caption         =   "F&ax:"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   3000
      Width           =   1155
   End
   Begin VB.Label labPhone 
      Caption         =   "&Phone:"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   2580
      Width           =   1155
   End
   Begin VB.Label labComments 
      Caption         =   "C&omments:"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   1140
      Width           =   1155
   End
   Begin VB.Label labPosition 
      Caption         =   "&Position:"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1155
   End
   Begin VB.Label labName 
      Caption         =   "&Name:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   300
      Width           =   1155
   End
End
Attribute VB_Name = "frmNewContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbolEditContact As Boolean
Private mintCustomerNo As Integer
Private mstrName As String
Private mlngID          As Long
Private objContact As New Contact
Public Event ContactCreated(ContactID As Long)

Private Function ValidFormInput() As Boolean
'***************************************
' Module/Form Name   : frmNewContact
'
' Procedure Name     : ValidFormInput
'
' Purpose            :
'
' Date Created       : 31/12/2002
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
On Error GoTo ValidFormInput_Error
'
'******** Code Starts Here *************
'
    msg_title = "New Contact"
    ValidFormInput = False

    If Len(Trim(txtName.Text)) = 0 Then
        MsgBox "A Name must be Entered", vbExclamation, msg_title
        txtName.SetFocus
        Exit Function
    End If
    ValidFormInput = True
'
'********* Code Ends Here **************
'
   Exit Function
'
ValidFormInput_Error:
    ErrorRaise "frmNewContact.ValidFormInput"
End Function

Public Sub DisplayForm(intCustomerNo As Integer, Optional EditContact As Variant)
'***************************************
' Module/Form Name   : frmNewContact
'
' Procedure Name     : DisplayForm
'
' Purpose            :
'
' Date Created       : 31/12/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : intCustomerNo - Integer
'                    : EditContact - Variant
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo DisplayForm_Error
'
'******** Code Starts Here *************
'
    Screen.MousePointer = vbHourglass
    mintCustomerNo = intCustomerNo
    If IsMissing(EditContact) Then
        mbolEditContact = False
    Else
        mbolEditContact = True
        'Set objContact = Contact
        'mstrName = Contact.Name
'        mstrName = EditContact
        mlngID = EditContact
    End If

    If mbolEditContact Then
        Me.Caption = "Edit Contact for " & customer.Read(mintCustomerNo)!customer_name
    Else
        Me.Caption = "New Contact for " & customer.Read(mintCustomerNo)!customer_name
    End If
    '
    Me.Show 1
    Screen.MousePointer = vbDefault
'
'********* Code Ends Here **************
'
   Exit Sub
'
DisplayForm_Error:
    ErrorRaise "frmNewContact.DisplayForm"
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    Set frmNewContact = Nothing
    DoEvents
End Sub

Private Sub cmdDone_Click()
'***************************************
' Module/Form Name   : frmNewContact
'
' Procedure Name     : cmdDone_Click
'
' Purpose            :
'
' Date Created       : 04/04/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo cmdDone_Click_Error
'
'******** Code Starts Here *************
'
    Screen.MousePointer = vbHourglass
    '
    '   Validate Form Input.
    '
    If Not ValidFormInput Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    With objContact
        '.read mintCustomerNo, mstrName
        If mbolEditContact Then
            .CustomerNo = mintCustomerNo
            .Name = txtName.Text
        End If
        .Position = txtPosition.Text
        .Comments = txtComments.Text
        .Phone = txtPhone.Text
        .Fax = txtFax.Text
        .Email = txtEmail.Text
        .Mail = chkMailingList.Value
        .WebUser = chkWebUser.Value
        .PreferredPhotoType = IIf(cboPreferredPhotoType.ListIndex = 0, "T", "D")
        If mbolEditContact Then
            On Error Resume Next
            .update
            If Err.Number - vbObjectError = 1 Then
                MsgBox Err.Description, vbExclamation, App.Title
                DisplayFields
                RaiseEvent ContactCreated(objContact.Name)
                Screen.MousePointer = vbDefault
                Exit Sub
            ElseIf Err.Number - vbObjectError = 2 Then
                MsgBox Err.Description, vbExclamation, App.Title
                txtName.SetFocus
                Screen.MousePointer = vbDefault
                Exit Sub
            ElseIf Err.Number - vbObjectError = 3 Or _
                   Err.Number - vbObjectError = 4 Then
                MsgBox Err.Description, vbExclamation, App.Title
                txtEmail.SetFocus
                Screen.MousePointer = vbDefault
                Exit Sub
            ElseIf Err.Number <> 0 Then
                ErrorSave
                On Error GoTo cmdDone_Click_Error
                ErrorRestore
            Else
                On Error GoTo cmdDone_Click_Error
            End If
        Else
            On Error Resume Next
            .Add mintCustomerNo, txtName.Text
            If Err.Number - vbObjectError = 3 Or _
               Err.Number - vbObjectError = 4 Then
                MsgBox Err.Description, vbExclamation, App.Title
                txtEmail.SetFocus
                Screen.MousePointer = vbDefault
                Exit Sub
            ElseIf Err.Number <> 0 Then
                ErrorSave
                On Error GoTo cmdDone_Click_Error
                ErrorRestore
            Else
                On Error GoTo cmdDone_Click_Error
            End If
        End If
    End With
    '
    '   Refresh the Customer Edit form.
    '
    RaiseEvent ContactCreated(objContact.ID)
    '
    RefreshScreens
    '
    Unload Me
    Set frmNewContact = Nothing
    Screen.MousePointer = vbDefault
'
'********* Code Ends Here **************
'
   Exit Sub
'
cmdDone_Click_Error:
    DisplayError , "frmNewContact.cmdDone_Click", vbExclamation
End Sub

Private Sub Form_Initialize()
    mbolEditContact = False
End Sub

Private Sub Form_Load()
'***************************************
' Module/Form Name   : frmNewContact
'
' Procedure Name     : Form_Load
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
On Error GoTo Form_Load_Error
'
'******** Code Starts Here *************
'
    With cboPreferredPhotoType
        .AddItem goSystemConfig.PhotoTypeDescription("T")
        .AddItem goSystemConfig.PhotoTypeDescription("D")
        .ListIndex = 0
    End With
    '
    If mbolEditContact Then
        DisplayFields
'        SendKeys "{TAB}"
    Else
''        txtName.Locked = False
''        txtName.BackColor = vbWindowBackground
    End If

    Screen.MousePointer = vbDefault
'
'********* Code Ends Here **************
'
   Exit Sub
'
Form_Load_Error:
    DisplayError , "frmNewContact.Form_Load", vbExclamation
End Sub

Private Sub Form_Paint()
'***************************************
' Module/Form Name   : frmNewContact
'
' Procedure Name     : Form_Paint
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
On Error GoTo Form_Paint_Error
'
'******** Code Starts Here *************
'
    If mbolEditContact Then
        txtPosition.SetFocus
    End If
'
'********* Code Ends Here **************
'
   Exit Sub
'
Form_Paint_Error:
    DisplayError , "frmNewContact.Form_Paint", vbExclamation
End Sub

Private Sub txtComments_GotFocus()
    HighLightText txtComments
End Sub

Private Sub txtEmail_GotFocus()
    HighLightText txtEmail
End Sub

Private Sub txtEmail_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(LCase(Chr(KeyAscii)))
End Sub

Private Sub txtFax_GotFocus()
    HighLightText txtFax
End Sub

Private Sub txtName_GotFocus()
    HighLightText txtName
End Sub

Private Sub txtPhone_GotFocus()
    HighLightText txtPhone
End Sub

Private Sub txtPosition_GotFocus()
    HighLightText txtPosition
End Sub

Private Sub DisplayFields()
'***************************************
' Module/Form Name   : frmNewContact
'
' Procedure Name     : DisplayFields
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
On Error GoTo DisplayFields_Error
'
'******** Code Starts Here *************
'
''    txtName.Locked = True
''    txtName.BackColor = vbButtonFace
    Set objContact = New Contact
    With objContact
        .CheckPoint
''        .Read mintCustomerNo, mstrName
        .Read mlngID
        txtName = .Name
        txtPosition.Text = .Position
        txtComments.Text = .Comments
        txtPhone.Text = .Phone
        txtFax.Text = .Fax
        txtEmail.Text = .Email
        chkMailingList = IIf(.Mail = True, 1, 0)
        chkWebUser = IIf(.WebUser = True, 1, 0)
        txtDateAmended.Caption = IIf(.DateAmended = 0, "", Format(.DateAmended, "dd/mm/yyyy"))
        cboPreferredPhotoType.ListIndex = IIf(.PreferredPhotoType = "T", 0, 1)
    End With
'
'********* Code Ends Here **************
'
   Exit Sub
'
DisplayFields_Error:
    ErrorRaise "frmNewContact.DisplayFields"
End Sub

Private Sub RefreshScreens()
'***************************************
' Module/Form Name   : frmNewContact
'
' Procedure Name     : RefreshScreens
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
On Error GoTo RefreshScreens_Error
'
'******** Code Starts Here *************
'
    Dim f As Form

    For Each f In Forms
        '
        '   Refresh the Edit Job screen.
        '
        If f.Name = "frm_job_edit" Then
            f.RefreshJob
        End If
''        '
''        '   Refresh the Job Details screen.
''        '
''        If f.Name = "frm_job_details" Then
''            f.RefreshContactDetails
''        End If
    Next f

'
'********* Code Ends Here **************
'
   Exit Sub
'
RefreshScreens_Error:
    ErrorRaise "frmNewContact.RefreshScreens"
End Sub
