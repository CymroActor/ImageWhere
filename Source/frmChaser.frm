VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmChaser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chase"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5430
   Icon            =   "frmChaser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   5430
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtMultipleDNotes 
      BackColor       =   &H8000000F&
      Height          =   375
      Left            =   4440
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   23
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtOriginalReturnByDate 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1200
      Width           =   1275
   End
   Begin VB.TextBox txtDeliveryNote 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   795
   End
   Begin VB.TextBox txtJobDescription 
      BackColor       =   &H8000000F&
      Height          =   555
      Left            =   2040
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   540
      Width           =   3315
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   555
      Left            =   3300
      Picture         =   "frmChaser.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5520
      Width           =   795
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   555
      Left            =   4560
      Picture         =   "frmChaser.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5520
      Width           =   795
   End
   Begin VB.Frame fraAction 
      Caption         =   "Action"
      Height          =   3735
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   5235
      Begin VB.ComboBox cboBlankNextAction 
         Enabled         =   0   'False
         Height          =   315
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   3360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox cboBlankChaserDate 
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   3360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox cboNextAction 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   3120
         Width           =   3240
      End
      Begin VB.ComboBox cboAction 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2220
         Width           =   3240
      End
      Begin VB.ComboBox cboContact 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   660
         Width           =   3240
      End
      Begin VB.TextBox txtUser 
         Height          =   285
         Left            =   1920
         MaxLength       =   20
         TabIndex        =   8
         Top             =   240
         Width           =   3240
      End
      Begin VB.TextBox txtComment 
         Height          =   1065
         Left            =   1920
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   1080
         Width           =   3240
      End
      Begin MSComCtl2.DTPicker dtpReturnBy 
         Height          =   315
         Left            =   3825
         TabIndex        =   16
         Top             =   2640
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   69992451
         CurrentDate     =   36779
         MinDate         =   29221
      End
      Begin VB.Label lblContact 
         Caption         =   "C&ontact:"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblAction 
         Caption         =   "&Action Completed:"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label lblUser 
         Caption         =   "&User:"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblReturnByDate 
         Caption         =   "Next &Chaser Date:"
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label lblNextAction 
         Caption         =   "&Next Action:"
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label lblComment 
         Caption         =   "Co&mments:"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   1080
         Width           =   1095
      End
   End
   Begin VB.Label lblOriginalReturnByDate 
      Caption         =   "Original Return By Date:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1260
      Width           =   1755
   End
   Begin VB.Label lblDeliveryNote 
      Caption         =   "Delivery Note No:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   180
      Width           =   1455
   End
   Begin VB.Label lblJobDescription 
      Caption         =   "Job Description:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmChaser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event ChaserUpdated(ID As Integer)
Public Event ChaserAdded(ID As Long)

Private moJob                   As Job2
Private moDeliveryNote          As DeliveryNote
Private moCustomer              As Customer2
Private moChaser                As Chaser
Private moChasers               As Chasers
Private moBusinessType          As BusinessType
Private moDeliveryNotes         As DeliveryNotes
Private mblnMultipleSelection   As Boolean
Public mMode                    As UpdateMode
Private mintNoSelectedDNotes    As Integer

Public Sub Display(ByVal Mode As UpdateMode, _
                   ByRef oJob As Job2, _
                   ByRef oChasers As Chasers, _
                   Optional ByRef oChaser As Chaser, _
                   Optional ByVal pstrNextAction As String = "")
'***************************************
' Module/Form Name   : frmChaser
'
' Procedure Name     : Display
'
' Purpose            :
'
' Date Created       : 29/12/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : Mode - UpdateMode
'                    : oJob - clsJob
'                    : oChasers - Chasers
'                    : oChaser - Chaser
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo Display_Error
'
'******** Code Starts Here *************
'
    mblnMultipleSelection = False
    mMode = Mode
    Set moJob = oJob
    Set moChasers = oChasers
    '
    '   Get the customer.
    '
    Set moCustomer = New Customer2
    moCustomer.Read moJob.CustomerNo
    '
    '   Get the Business Type.
    '
    Set moBusinessType = New BusinessType
    moBusinessType.Read moCustomer.BusinessType
    '
    '   Get the Delivery Note
    '
    Set moDeliveryNote = New DeliveryNote
    moDeliveryNote.Read moJob.DeliveryNoteNo
    '
    '   Initialise the screen controls.
    '
    InitialiseControls IIf(pstrNextAction = "", moChasers.LatestNextAction, pstrNextAction)
    '
    '   Display the title fields.
    '
    txtDeliveryNote.Text = moJob.DeliveryNoteNo
    txtJobDescription.Text = moJob.reference
    '
    If mMode = Edit Then
        Set moChaser = oChaser
        Me.Caption = "Edit Chaser for Delivery Note " & oChaser.DeliveryNoteNo
        DisplayChaser
        SendKeys "{TAB}"
    Else
        Set moChaser = New Chaser
        Me.Caption = "Add Chaser for Delivery Note " & moJob.DeliveryNoteNo
        SendKeys "{TAB}{TAB}{TAB}"
    End If
    Screen.MousePointer = vbDefault
    Me.Show 1
'
'********* Code Ends Here **************
'
   Exit Sub
'
Display_Error:
    ErrorRaise "frmChaser.Display"
End Sub

Public Sub DisplayMultiple(ByRef poDeliveryNotes As DeliveryNotes)
'***************************************
' Module/Form Name   : frmChaser
'
' Procedure Name     : DisplayMultiple
'
' Purpose            :
'
' Date Created       : 22/05/2006 23:43
'
' Author             : Gareth
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
    On Error GoTo ErrorProc
    
    Dim oDeliveryNote       As DeliveryNote
    Dim oCustomer           As Customer2

    mblnMultipleSelection = True
    '
    '   Display the title fields.
    '
    Me.Caption = "Add Chaser for Multiple Delivery Notes"
    '
    txtDeliveryNote.Visible = False
    lblDeliveryNote.Caption = "Delivery Note Nos:"
    txtJobDescription.Visible = False
    txtOriginalReturnByDate.Visible = False
    lblJobDescription.Visible = False
    lblOriginalReturnByDate.Visible = False
    '
    With txtMultipleDNotes
        .Visible = True
        .Top = txtDeliveryNote.Top
        .Left = txtDeliveryNote.Left
        .Height = fraAction.Top - txtDeliveryNote.Top
        .Width = 1000
        mintNoSelectedDNotes = 0
        For Each oDeliveryNote In poDeliveryNotes
            If oDeliveryNote.Selected = True Then
                mintNoSelectedDNotes = mintNoSelectedDNotes + 1
                .Text = .Text & IIf(.Text <> "", vbCrLf, "") & oDeliveryNote.DNoteNo
            End If
        Next oDeliveryNote
    End With
    '
    '   Get the customer.
    '
    Set moCustomer = New Customer2
    For Each oDeliveryNote In poDeliveryNotes
        If oDeliveryNote.Selected Then
            moCustomer.Read oDeliveryNote.CustomerNo
            Exit For
        End If
    Next oDeliveryNote
    '
    '   Get the Business Type.
    '
    Set moBusinessType = New BusinessType
    moBusinessType.Read moCustomer.BusinessType
    '
    '   Set the module level delivery note.
    '
    Set moDeliveryNotes = poDeliveryNotes
    '
    '   Initialise the screen controls.
    '
    InitialiseControls moDeliveryNotes.LatestNextAction
    '
    Me.Show 1

    On Error GoTo 0
    Exit Sub
'
'********* Code Ends Here **************
'
ErrorProc:
    ErrorRaise "frmChaser.DisplayMultiple"
End Sub

Private Sub cboAction_Click()
'***************************************
' Module/Form Name   : frmChaser
'
' Procedure Name     : cboAction_Click
'
' Purpose            :
'
' Date Created       : 29/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo cboAction_Click_Error
'
'******** Code Starts Here *************
'
    DetermineReturnByDate
    '
    '   If Invoiced then disable Next Action etc.
    '
    If cboAction.ListIndex = 6 Then
        dtpReturnBy.Enabled = False
        With cboBlankChaserDate
            .Top = dtpReturnBy.Top
            .Left = dtpReturnBy.Left
            .Width = dtpReturnBy.Width
            .Visible = True
        End With
        cboNextAction.Enabled = False
        With cboBlankNextAction
            .Top = cboNextAction.Top
            .Left = cboNextAction.Left
            .Width = cboNextAction.Width
            .Visible = True
        End With
    Else
        dtpReturnBy.Enabled = True
        cboBlankChaserDate.Visible = False
        cboNextAction.Enabled = True
        cboBlankNextAction.Visible = False
        cboNextAction.Text = DetermineNextAction(cboAction.Text)
    End If
'
'********* Code Ends Here **************
'
   Exit Sub
'
cboAction_Click_Error:
    DisplayError , "frmChaser.cboAction_Click", vbExclamation
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    DoEvents
End Sub

Private Sub DisplayChaser()
'***************************************
' Module/Form Name   : frmChaser
'
' Procedure Name     : DisplayChaser
'
' Purpose            :
'
' Date Created       : 29/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo DisplayChaser_Error
'
'******** Code Starts Here *************
'
    If Not LocateComboItem(cboContact, moChaser.Contact) Then
        MsgBox "Contact '" & moChaser.Contact & "' is not currently recorded as working for this Customer." & vbCrLf & "Please create another Chaser record."
        cboContact.ListIndex = -1
    End If
'
'********* Code Ends Here **************
'
   Exit Sub
'
DisplayChaser_Error:
    ErrorRaise "frmChaser.DisplayChaser"
End Sub

Private Sub InitialiseControls(strAction As String)
'***************************************
' Module/Form Name   : frmChaser
'
' Procedure Name     : InitialiseControls
'
' Purpose            :
'
' Date Created       : 29/12/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : strAction - String
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo InitialiseControls_Error
'
'******** Code Starts Here *************
'
    Dim oContact As Contact

    moCustomer.Contacts.Refresh
    For Each oContact In moCustomer.Contacts
        cboContact.AddItem oContact.Name
    Next oContact
    cboContact.ListIndex = -1
    '
    With cboNextAction
        .AddItem "SL1"
        .AddItem "SL2"
        .AddItem "Phone 1"
        .AddItem "Phone 2"
        .AddItem "Loss/Fee"
        .AddItem "Miscellaneous"
        .ListIndex = -1
    End With
    '
    With cboAction
        .AddItem "SL1"
        .AddItem "SL2"
        .AddItem "Phone 1"
        .AddItem "Phone 2"
        .AddItem "Loss/Fee"
        .AddItem "Miscellaneous"
        .AddItem "Invoiced"
        If strAction = "" Then
            cboAction.Text = "Miscellaneous"
        ElseIf Not LocateComboItem(cboAction, strAction) Then
            Err.Raise vbObjectError + 1, , "Next Action from previous Chaser record - " & strAction & " - does not exist"
        End If
    End With
    '
    '   Default the Contact.
    '
    If Not mblnMultipleSelection Then
        If Not LocateComboItem(cboContact, moChasers.LatestContact) Then
            MsgBox moChasers.LatestContact & " is no longer a contact for " & moCustomer.CustomerName, vbInformation
        End If
    '
        txtOriginalReturnByDate.Text = Format(moChasers.OriginalReturnByDate, "dd/mm/yyyy")
    End If
    'cboContact.Text = moChasers.LatestContact
    txtUser.Text = Left(goCompanyInfo.Signatory, InStr(goCompanyInfo.Signatory, " "))
'
'********* Code Ends Here **************
'
   Exit Sub
'
InitialiseControls_Error:
    ErrorRaise "frmChaser.InitialiseControls"
End Sub

Private Sub cmdOK_Click()
'***************************************
' Module/Form Name   : frmChaser
'
' Procedure Name     : cmdOK_Click
'
' Purpose            :
'
' Date Created       : 29/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo cmdOK_Click_Error
'
'******** Code Starts Here *************
'
    If Not ValidInput Then
        Exit Sub
    End If

    If mblnMultipleSelection Then
        AddMultipleChasers
    ElseIf mMode = Edit Then
        With moChaser
            .Contact = cboContact.Text
            .Action = cboAction.Text
            .User = txtUser.Text
            .ChaserDate = dtpReturnBy.Value
            .NextAction = cboNextAction.Text
            .Comment = txtComment.Text
            On Error Resume Next
            .update
            If Err.Number = vbObjectError + 1 Then
                DisplayError
                DisplayChaser
                Exit Sub
            Else
                If Err.Number <> 0 Then
                    ErrorSave
                    On Error GoTo cmdOK_Click_Error
                    ErrorRestore
                Else
                    On Error GoTo cmdOK_Click_Error
                End If
            End If
            RaiseEvent ChaserUpdated(.ID)
        End With
    Else
        begin_trans
            '
            '   Set any digital statuses to Invoiced if necessary.
            '
            SetInvoicedDigitalStatus moDeliveryNote
            '
            moChaser.CreateDAO moJob.DeliveryNoteNo, _
                               cboContact.Text, _
                               cboAction.Text, _
                               txtUser.Text, _
                               IIf(cboAction.Text = "Invoiced", 0, dtpReturnBy.Value), _
                               IIf(cboAction.Text = "Invoiced", "", cboNextAction.Text), _
                               txtComment.Text
        commit_trans
        RaiseEvent ChaserAdded(moChaser.ID)
    End If

    Unload Me
    DoEvents
'
'********* Code Ends Here **************
'
   Exit Sub
'
cmdOK_Click_Error:
    DisplayError , "frmChaser.cmdOK_Click", vbExclamation
End Sub

Private Sub SetInvoicedDigitalStatus(ByRef poDeliveryNote As DeliveryNote)
    Dim oSearchResult   As SearchResult
    
'***************************************
' Module/Form Name   : frmChaser
'
' Procedure Name     : SetInvoicedDigitalStatus
'
' Purpose            :
'
' Date Created       : 10/10/2006 00:04
'
' Author             : Gareth
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
    On Error GoTo ErrorProc

    If cboAction.Text = "Invoiced" Then
        poDeliveryNote.SearchResults.Refresh plngDNoteNo:=poDeliveryNote.DNoteNo
        For Each oSearchResult In poDeliveryNote.SearchResults
            If oSearchResult.DigitalStatus = "" Then
                oSearchResult.DigitalStatus = "A"
                oSearchResult.UpdateDAO
            End If
        Next oSearchResult
    End If

    On Error GoTo 0
    Exit Sub
'
'********* Code Ends Here **************
'
ErrorProc:
    ErrorRaise "frmChaser.SetInvoicedDigitalStatus"
End Sub

Private Sub AddMultipleChasers()
    Dim oDeliveryNote       As DeliveryNote
    
'***************************************
' Module/Form Name   : frmChaser
'
' Procedure Name     : AddMultipleChasers
'
' Purpose            :
'
' Date Created       : 06/08/2006 12:16
'
' Author             : Gareth
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
    On Error GoTo ErrorProc

    For Each oDeliveryNote In moDeliveryNotes
        If oDeliveryNote.Selected = True Then
            begin_trans
                '
                '   Set any digital statuses to Invoiced if necessary.
                '
                SetInvoicedDigitalStatus oDeliveryNote
                '
                Set moChaser = Nothing
                Set moChaser = New Chaser
                moChaser.CreateDAO oDeliveryNote.DNoteNo, _
                                   cboContact.Text, _
                                   cboAction.Text, _
                                   txtUser.Text, _
                                   dtpReturnBy.Value, _
                                   cboNextAction.Text, _
                                   txtComment.Text
            commit_trans
        End If
    
    Next oDeliveryNote
    RaiseEvent ChaserAdded(moChaser.ID)
    Set moChaser = Nothing

    On Error GoTo 0
    Exit Sub
'
'********* Code Ends Here **************
'
ErrorProc:
    ErrorRaise "frmChaser.AddMultipleChasers"
End Sub

Private Function ValidInput() As Boolean
'***************************************
' Module/Form Name   : frmChaser
'
' Procedure Name     : ValidInput
'
' Purpose            :
'
' Date Created       : 29/12/2002
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
On Error GoTo ValidInput_Error
'
'******** Code Starts Here *************
'
    Dim oDeliveryNote               As DeliveryNote
    Dim sDNotesOutstandingMess      As String
    
    ValidInput = False
    If cboContact.ListIndex = -1 Then
        MsgBox "Select a Contact", vbExclamation
        cboContact.SetFocus
        Exit Function
    End If

    If Trim(txtUser.Text) = "" Then
        MsgBox "Enter your User ID", vbExclamation
        txtUser.SetFocus
        Exit Function
    End If

    If cboAction.ListIndex <> 6 And cboNextAction.ListIndex = -1 Then
        MsgBox "Select a 'Next Action'", vbExclamation
        cboNextAction.SetFocus
        Exit Function
    End If
    '
    '   Check O/S pictures if setting to Invoiced.
    '
    If cboAction.Text = "Invoiced" Then
        If mblnMultipleSelection Then
            sDNotesOutstandingMess = ""
            For Each oDeliveryNote In moDeliveryNotes
                If oDeliveryNote.Selected Then
                    If oDeliveryNote.TotalOutstandingTrans > 0 Then
                        sDNotesOutstandingMess = sDNotesOutstandingMess & vbCrLf & CStr(oDeliveryNote.DNoteNo)
                    End If
                End If
            Next oDeliveryNote
            If sDNotesOutstandingMess <> "" Then
                MsgBox "There are outstanding pictures on the following delivery notes:" & vbCrLf & _
                       sDNotesOutstandingMess & vbCrLf & vbCrLf & _
                       "These must be returned before invoicing can be completed.", vbExclamation
                On Error Resume Next
                cboAction.SetFocus
                Exit Function
            End If
        Else
                    
            Set oDeliveryNote = New DeliveryNote
            oDeliveryNote.Read moJob.DeliveryNoteNo
            If oDeliveryNote.TotalOutstandingTrans > 0 Then
                MsgBox "There are outstanding pictures on this delivery note." & vbCrLf & _
                       "These must be returned before invoicing can be completed.", vbExclamation
                Set oDeliveryNote = Nothing
                On Error Resume Next
                cboAction.SetFocus
                Exit Function
            End If
            Set oDeliveryNote = Nothing
        End If
    End If
    '
    ValidInput = True
'
'********* Code Ends Here **************
'
   Exit Function
'
ValidInput_Error:
    ErrorRaise "frmChaser.ValidInput"
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set moCustomer = Nothing
    Set moBusinessType = Nothing
    Set moDeliveryNotes = Nothing
    Set moDeliveryNote = Nothing
End Sub

Private Sub txtUser_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub DetermineReturnByDate()
'***************************************
' Module/Form Name   : frmChaser
'
' Procedure Name     : DetermineReturnByDate
'
' Purpose            :
'
' Date Created       : 29/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo DetermineReturnByDate_Error
'
'******** Code Starts Here *************
'
    Select Case cboAction.ListIndex
        Case Is = 0
            dtpReturnBy.Value = DateAdd("d", moBusinessType.SL1ReturnPeriod, Now)
        Case Is = 1
            dtpReturnBy.Value = DateAdd("d", moBusinessType.SL2ReturnPeriod, Now)
        Case Is = 2
            dtpReturnBy.Value = DateAdd("d", moBusinessType.Phone1ReturnPeriod, Now)
        Case Is = 3
            dtpReturnBy.Value = DateAdd("d", moBusinessType.Phone2ReturnPeriod, Now)
        Case Is = 4
            dtpReturnBy.Value = DateAdd("d", moBusinessType.LossFeeReturnPeriod, Now)
        Case Is = 5
            If mblnMultipleSelection Then
                dtpReturnBy.Value = Format(moDeliveryNotes.LatestChaserDate, "dd/mm/yyyy")
            Else
                dtpReturnBy.Value = Format(moChasers.LatestChaserDate, "dd/mm/yyyy")
            End If
        Case Is = 6
            dtpReturnBy.Value = dtpReturnBy.MinDate
    End Select
'
'********* Code Ends Here **************
'
   Exit Sub
'
DetermineReturnByDate_Error:
    ErrorRaise "frmChaser.DetermineReturnByDate"
End Sub

Private Function DetermineNextAction(Action As String) As String
'***************************************
' Module/Form Name   : frmChaser
'
' Procedure Name     : DetermineNextAction
'
' Purpose            :
'
' Date Created       : 29/12/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : Action - String
'
' Returns            : String
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo DetermineNextAction_Error
'
'******** Code Starts Here *************
'
    Select Case Action
        Case Is = "None"
            DetermineNextAction = "SL1"
        Case Is = "SL1"
            DetermineNextAction = "SL2"
        Case Is = "SL2"
            DetermineNextAction = "Phone 1"
        Case Is = "Phone 1"
            DetermineNextAction = "Phone 2"
        Case Is = "Phone 2"
            DetermineNextAction = "Loss/Fee"
        Case Is = "Loss/Fee"
            DetermineNextAction = "Loss/Fee"
        Case Is = "Miscellaneous"
            If mblnMultipleSelection Then
                DetermineNextAction = "Miscellaneous"
            ElseIf moChasers.LatestNextAction = "" Then
                DetermineNextAction = "Miscellaneous"
            Else
                DetermineNextAction = moChasers.LatestNextAction
            End If
        End Select
'
'********* Code Ends Here **************
'
   Exit Function
'
DetermineNextAction_Error:
    ErrorRaise "frmChaser.DetermineNextAction"
End Function

