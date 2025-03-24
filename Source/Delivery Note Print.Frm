VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_delivery_note_print 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Delivery Note"
   ClientHeight    =   4620
   ClientLeft      =   2760
   ClientTop       =   1815
   ClientWidth     =   6930
   Icon            =   "Delivery Note Print.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4620
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkBundleDelivery 
      Alignment       =   1  'Right Justify
      Caption         =   "Bundle Delivery"
      Height          =   255
      Left            =   5160
      TabIndex        =   16
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Frame fraDispatchMethod 
      Caption         =   "Dispatch Method"
      ForeColor       =   &H80000001&
      Height          =   1035
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   6675
      Begin VB.TextBox txtDispatchDigital 
         Height          =   285
         Left            =   1740
         MaxLength       =   30
         TabIndex        =   13
         Top             =   600
         Width           =   4815
      End
      Begin VB.TextBox txtDispatchTransparency 
         Height          =   285
         Left            =   1740
         MaxLength       =   30
         TabIndex        =   11
         Top             =   240
         Width           =   4815
      End
      Begin VB.Label lblDispatchDigital 
         Caption         =   "Digital:"
         Height          =   240
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   1395
      End
      Begin VB.Label lblDispatchTransparency 
         Caption         =   "Transparency:"
         Height          =   240
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1395
      End
   End
   Begin VB.TextBox txt_memo 
      Height          =   495
      Left            =   1860
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   1440
      Width           =   4815
   End
   Begin VB.Frame fraCustomerJob 
      Caption         =   "Customer Job"
      ForeColor       =   &H80000001&
      Height          =   1275
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6675
      Begin VB.Label lab_address_line_1 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1740
         TabIndex        =   4
         Top             =   540
         Width           =   4815
      End
      Begin VB.Label lab_address_line_1_lab 
         Caption         =   "First line of address:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   540
         Width           =   1455
      End
      Begin VB.Label lab_customer_name 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1740
         TabIndex        =   2
         Top             =   240
         Width           =   4815
      End
      Begin VB.Label lab_customer_name_lab 
         Caption         =   "Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lab_job_reference 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1740
         TabIndex        =   6
         Top             =   840
         Width           =   4815
      End
      Begin VB.Label lab_job_reference_lab 
         Caption         =   "Reference:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   855
      End
   End
   Begin VB.CommandButton ssc_cancel 
      Caption         =   "&Cancel"
      Height          =   675
      Left            =   6000
      Picture         =   "Delivery Note Print.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   3840
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   675
      Left            =   5040
      Picture         =   "Delivery Note Print.frx":089C
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3840
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.TextBox txt_loss_damage_fee 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1860
      MaxLength       =   8
      TabIndex        =   22
      Top             =   4200
      Width           =   855
   End
   Begin VB.TextBox txt_weekly_holding_fee 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1860
      MaxLength       =   8
      TabIndex        =   20
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox txt_service_fee 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1860
      MaxLength       =   8
      TabIndex        =   18
      Top             =   3480
      Width           =   855
   End
   Begin Crystal.CrystalReport crs_delivery_note 
      Left            =   3720
      Top             =   3840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowBorderStyle=   1
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin MSComCtl2.DTPicker txt_return_date 
      Height          =   285
      Left            =   1860
      TabIndex        =   15
      Top             =   3120
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   115736579
      CurrentDate     =   36844
   End
   Begin VB.Label lab_loss_damage_fee_lab 
      Caption         =   "Loss/Damage Fee:"
      Height          =   195
      Left            =   240
      TabIndex        =   21
      Top             =   4200
      Width           =   1515
   End
   Begin VB.Label lab_weekly_holding_fee_lab 
      Caption         =   "Weekly Holding Fee:"
      Height          =   195
      Left            =   240
      TabIndex        =   19
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label lab_service_fee_lab 
      Caption         =   "Service Fee:"
      Height          =   195
      Left            =   240
      TabIndex        =   17
      Top             =   3480
      Width           =   1035
   End
   Begin VB.Label lab_requested_by 
      Caption         =   "Return date:"
      Height          =   195
      Left            =   240
      TabIndex        =   14
      Top             =   3120
      Width           =   1035
   End
   Begin VB.Label lab_memo_lab 
      Caption         =   "Memo:"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   1440
      Width           =   1395
   End
End
Attribute VB_Name = "frm_delivery_note_print"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
Private moJob                       As Job2
Private moChasers                   As Chasers
Private SQL                         As String
Private mod_customer_no             As Integer
Private mod_job_no                  As Long
Private mblnDispatchTransparency    As Boolean
Private mblnDispatchDigital         As Boolean
Private mblnInvoiceCreated          As Boolean
Private WithEvents mfChaser         As frmChaser
Attribute mfChaser.VB_VarHelpID = -1
Public cancel_list As Boolean

Public Sub display_customer_and_job(job_no As Long, Optional pblnRedisplay As Boolean = False)
'***************************************
' Module/Form Name   : frm_delivery_note_print
'
' Procedure Name     : display_customer_and_job
'
' Purpose            :
'
' Date Created       : 29/07/2001
'
' Author             : GARETH
'
' Parameters         : job_no - Long
'                    : pblnRedisplay -
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 19/05/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo display_customer_and_job_Error
'
'******** Code Starts Here *************
'
    Dim oBusinessType               As BusinessType
    Dim oCustomer                   As Customer2
    Dim oSearch                     As Search2
    Dim oSearchResult               As SearchResult
    Dim oDeliveryNote               As DeliveryNote
    '
    '   Get job.
    '
    Set moJob = New Job2
    moJob.Read job_no
    '
    lab_job_reference = Duplicate_character(moJob.reference, "&")
    '
    '   Which dispatch methods are appropriate. Also check to see if bundling should be defaulted.
    '
    mblnDispatchDigital = False
    mblnDispatchTransparency = False
    For Each oSearch In moJob.Searches
        If oSearch.PhotoType = "D" Then
            For Each oSearchResult In oSearch.SearchResults
                If oSearchResult.Confirmed Then
                    mblnDispatchDigital = True
                    Exit For
                End If
            Next oSearchResult
        ElseIf oSearch.PhotoType = "T" Then
            For Each oSearchResult In oSearch.SearchResults
                If oSearchResult.Confirmed Then
                    mblnDispatchTransparency = True
                    Exit For
                End If
            Next oSearchResult
        End If
    Next oSearch
    '
    '   Get customer and display fields.
    '
    Set oCustomer = New Customer2
    On Error Resume Next
    oCustomer.Read moJob.CustomerNo
    If Err.Number = vbObjectError + 8 And _
     Err.Number <> 0 Then      '   Customer Not Found
        ErrorSave
        On Error GoTo display_customer_and_job_Error
        ErrorRestore
    Else
        On Error GoTo display_customer_and_job_Error
    End If
    '
    With oCustomer
        mod_customer_no = .CustomerNo
        lab_customer_name.Caption = Replace(.CustomerName, "&", "&&")
        lab_address_line_1.Caption = Replace(.Address1, "&", "&&")
    End With
''    '
''    '   Only allow bundling if a web job.
''    '
''    chkBundleDelivery.Visible = (moJob.WebStatus <> "")
    '
    '   If job has been delivered, then only allow a reprint.
    '
    If moJob.DeliveryNoteNo > 0 Then
    
        Me.Caption = "Reprint Delivery Note - " & CStr(moJob.DeliveryNoteNo)
        '
        '   Get delivery note.
        '
        Set oDeliveryNote = New DeliveryNote
        oDeliveryNote.Read moJob.DeliveryNoteNo
        
        txtDispatchDigital.Text = oDeliveryNote.DispatchDigital
        txtDispatchTransparency.Text = oDeliveryNote.DispatchTransparency
        txt_service_fee.Text = oDeliveryNote.ServiceFee
        txt_weekly_holding_fee.Text = oDeliveryNote.WeeklyHoldingFee
        txt_loss_damage_fee.Text = oDeliveryNote.LossDamageFee
        txt_return_date.Value = oDeliveryNote.ReturnByDate
        txt_memo.Text = oDeliveryNote.Memo
        chkBundleDelivery.Value = vbChecked
        chkBundleDelivery.Enabled = False
        
        txtDispatchDigital.Locked = True
        txtDispatchDigital.BackColor = vbButtonFace
        txtDispatchTransparency.Locked = True
        txtDispatchTransparency.BackColor = vbButtonFace
        txt_service_fee.Locked = True
        txt_service_fee.BackColor = vbButtonFace
        txt_weekly_holding_fee.Locked = True
        txt_weekly_holding_fee.BackColor = vbButtonFace
        txt_loss_damage_fee.Locked = True
        txt_loss_damage_fee.BackColor = vbButtonFace
        txt_return_date.Enabled = False
        txt_memo.Locked = True
        txt_memo.BackColor = vbButtonFace
    Else
                
        '
        txtDispatchDigital.Enabled = mblnDispatchDigital
        txtDispatchDigital.BackColor = IIf(mblnDispatchDigital, vbWindowBackground, vbButtonFace)
        txtDispatchTransparency.Enabled = mblnDispatchTransparency
        txtDispatchTransparency.BackColor = IIf(mblnDispatchTransparency, vbWindowBackground, vbButtonFace)
        '
        '   Display Customer fields.
        '
        With oCustomer
            If Not pblnRedisplay Then
                txtDispatchTransparency.Text = IIf(mblnDispatchTransparency, .PreferredDelivery, "")
                txtDispatchDigital.Text = ""
            End If
            If .ServiceFeeRequired Then
                txt_service_fee.Text = Format(IIf(mblnDispatchTransparency, .ServiceFee, 0), "##,##0.00")
            Else
                txt_service_fee = Format(0, "##,##0.00")
            End If
        End With
        
        Set oBusinessType = New BusinessType
        oBusinessType.Read oCustomer.BusinessType
        txt_return_date.Value = DateAdd("d", oBusinessType.InitialReturnPeriod, Date)     '   Default date for return
        
        Set oBusinessType = Nothing
        '
        '   Default Fees to be set up.
        '
        txt_weekly_holding_fee = Format(IIf(mblnDispatchTransparency, 4, 0), "##,##0.00")
        txt_loss_damage_fee = Format(IIf(mblnDispatchTransparency, 500, 0), "##,##0.00")
        '
        '   Retrieve and display the previous memo for this workstation.
        '
        txt_memo = GetSetting(App.Title, "DeliveryNote", "Memo", "")
        '
        chkBundleDelivery.Value = IIf(mblnDispatchDigital, vbChecked, vbUnchecked)
        chkBundleDelivery.Visible = IIf(mblnDispatchDigital, True, False)
    End If
    '
    If Not pblnRedisplay Then
        Me.Show 1
    End If
    Set oCustomer = Nothing
'
'********* Code Ends Here **************
'
    Exit Sub
    '
display_customer_and_job_Error:
    ErrorRaise "frm_delivery_note_print.display_customer_and_job"
End Sub

Private Function valid_input() As Boolean
'***************************************
' Module/Form Name   : frm_delivery_note_print
'
' Procedure Name     : valid_input
'
' Purpose            :
'
' Date Created       : 07/12/2002
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
    
    If Not IsDate(txt_return_date) Then
        MsgBox "Enter a valid date", vbExclamation
        txt_return_date.SetFocus
        valid_input = False
        Exit Function
    End If
    '
    '   Check to see if date is in past.
    '
    If DateDiff("d", Now, CDate(txt_return_date.Value)) < 2 Then
        If MsgBox("Return date is less than two days from now." & vbCrLf & _
                   "Are you sure you want to use this date", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            txt_return_date.SetFocus
            valid_input = False
            Exit Function
        End If
    End If
    '
    '   Validate the Service Fee - Don't allow more than two decimal places.
    '
    If Not IsNumeric(txt_service_fee) Then
        MsgBox "Enter a valid number", vbExclamation
        txt_service_fee.SetFocus
        valid_input = False
        Exit Function
    End If
    
    If InStr(txt_service_fee.Text, ".") <> 0 Then
        If Len(txt_service_fee.Text) - InStr(txt_service_fee.Text, ".") > 2 Then
            MsgBox "Enter a valid amount with no more than two decimal places.", vbExclamation
            txt_service_fee.SetFocus
            valid_input = False
            Exit Function
        End If
    End If

    If Not IsNumeric(txt_weekly_holding_fee) Then
        MsgBox "Enter a valid number", vbExclamation
        txt_weekly_holding_fee.SetFocus
        valid_input = False
        Exit Function
    End If

    If Not IsNumeric(txt_loss_damage_fee) Then
        MsgBox "Enter a valid number", vbExclamation
        txt_loss_damage_fee.SetFocus
        valid_input = False
        Exit Function
    End If
    '
    '   Dispatch Methods:
    '
    If mblnDispatchTransparency Then
        If Trim(txtDispatchTransparency.Text) = "" Then
            MsgBox "Enter a Transparency Dispatch Method for the Delivery", vbExclamation
            txtDispatchTransparency.SetFocus
            valid_input = False
            Exit Function
        End If
    End If
    '
    If mblnDispatchDigital Then
        If Trim(txtDispatchDigital.Text) = "" Then
            MsgBox "Enter a Digital Dispatch Method for the Delivery", vbExclamation
            txtDispatchDigital.SetFocus
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
    ErrorRaise "frm_delivery_note_print.valid_input"
End Function

Private Sub Form_Load()
    com_position_form Me
    mblnInvoiceCreated = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set moJob = Nothing
End Sub

Private Sub mfChaser_ChaserAdded(ID As Long)
    mblnInvoiceCreated = True
End Sub

Private Sub ssc_cancel_Click()
    Unload Me
    Set frm_delivery_note_print = Nothing
End Sub

Private Sub cmdPrint_Click()
'***************************************
' Module/Form Name   : frm_delivery_note_print
'
' Procedure Name     : cmdPrint_Click
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
On Error GoTo cmdPrint_Click_Error
'
'******** Code Starts Here *************
'
    Dim oOutlook                As OutlookClass
    Dim msg                     As String
    Dim style                   As String
    Dim response                As Integer
    Dim abort                   As Boolean
    Dim delivery_note_no        As Long
    Dim oContact                As Contact
    Dim strEmail                As String
    Dim oSearch                 As Search2
    Dim blnReprint              As Boolean
    Dim blnInvoiceDeliveryNote  As Boolean
    Dim oChaser                 As Chaser
    Dim fChaser                 As frmChaser
    Dim fPreparedFiles          As frmPreparedFiles
    Dim strPreparedFilesMessage As String
    
    Screen.MousePointer = vbHourglass
    
    blnReprint = (moJob.DeliveryNoteNo <> 0)
    
    If Not valid_input Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    '
    SaveSetting App.Title, "DeliveryNote", "Memo", txt_memo
    If Not blnReprint Then
        '
        '
        '   Issue a warning that the Job Sheet will become definitive etc.
        '
        msg = "PRINTING THE DELIVERY NOTE WILL MAKE THE JOB DEFINITIVE." + vbCr _
            + "YOU WILL NO LONGER BE ABLE TO EDIT THE JOB." + vbCr + "DO YOU WISH TO CONTINUE?"
    
        style = vbYesNo + vbInformation + vbDefaultButton2
        response = MsgBox(msg, style)
    
        If response = vbNo Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        '
        '   Remind the user to raise a manual invoice if necessary for any service fees.
        '
        If CCur(txt_service_fee.Text) <> 0 Then
            MsgBox "A Service Fee has been applied to this Delivery Note." & vbCrLf & _
                   "A manual invoice must therefore be raised for " & Format(CCur(txt_service_fee.Text), "£#,###,##0.00"), vbInformation, App.Title
        End If
        '
        '   If this is a Digital Job only and the pictures have been bundled, then determine whether the user wishes to raise an invoice for
        '   the whole job. If so then set Delivery Note to Invoiced.
        '
        blnInvoiceDeliveryNote = False
        If mblnDispatchDigital And Not mblnDispatchTransparency Then
            If chkBundleDelivery.Value = vbChecked Then
                If MsgBox("Do you wish an invoice for the digital images to be raised immediately for this job?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                    blnInvoiceDeliveryNote = True
                End If
            End If
        End If
        '
        goLog.WriteDetailedLog "Starting delivery of job - '" & CStr(moJob.JobNo) & "'", "frm_delivery_note_print.cmdPrint_Click"
        '
        begin_trans
            '
            '   Deliver the job sheet.
            '
            goLog.WriteDetailedLog "Starting delivery of job - '" & CStr(moJob.JobNo) & "'", "frm_delivery_note_print.cmdPrint_Click"
            On Error Resume Next
            Set oContact = moJob.Contact
            If Err.Number - vbObjectError = 9 Then
                MsgBox "Contact for the Job '" & moJob.RequestedBy & "' cannot be found." & vbCrLf & _
                       "Please update the Customer's contacts or change the contact for the job.", vbExclamation
                roll_back
                Screen.MousePointer = vbDefault
                Exit Sub
            ElseIf Err.Number <> 0 Then
                ErrorSave
                On Error GoTo cmdPrint_Click_Error
                ErrorRestore
            Else
                On Error GoTo cmdPrint_Click_Error
            End If
            '
            moJob.DispatchTransparency = txtDispatchTransparency.Text
            moJob.DispatchDigital = txtDispatchDigital.Text
            moJob.ServiceFee = txt_service_fee.Text
            job.deliver frm_delivery_note_print, moJob, delivery_note_no, abort
            If abort = True Then
                roll_back
                display_customer_and_job moJob.JobNo, True
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            '
            '   If invoice required then use chaser form to select a contact.
            '
            If blnInvoiceDeliveryNote Then
                mblnInvoiceCreated = False
                
                Set moChasers = New Chasers
                moChasers.DeliveryNoteNo = moJob.DeliveryNoteNo
                
                Set mfChaser = New frmChaser
                mfChaser.Display Add, _
                                 moJob, _
                                 moChasers, _
                                 pstrNextAction:="Invoiced"
                If Not mblnInvoiceCreated Then
                    Set mfChaser = Nothing
                    roll_back
                    display_customer_and_job moJob.JobNo, True
                    MsgBox "Delivery Cancelled"
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                Set mfChaser = Nothing
            End If
            '
        commit_trans
        '
        '   Prepare all Digital Images.
        '
        Screen.MousePointer = vbHourglass
        moJob.PrepareDigitalSearchesForDelivery
        '
        '   Now print the delivery Note.
        '   (Hide the form first so that IW doesn't take the focus away from Crystal.)
        '
        Me.Hide
        job.print_delivery_note crs_delivery_note, _
                                delivery_note_no, _
                                moJob.CustomerNo, _
                                "A", _
                                IIf(chkBundleDelivery.Value = vbChecked, jdnPDF, jdnOnLine), _
                                moJob.DigitalDeliveryNote, _
                                abort
        If abort = True Then
            roll_back
            display_customer_and_job moJob.JobNo, True
            Set moJob = Nothing
            Screen.MousePointer = vbDefault
            goLog.WriteDetailedLog "Delivery note NOT printed for delivery note number - '" & CStr(delivery_note_no) & "'", "frm_delivery_note_print.cmdPrint_Click"
            Exit Sub
        End If
    End If
    '
    '   If Required, email a Customer to let them know that the Delivery Note has been sent.
    '
    If mblnDispatchDigital Then
        If Not EmailDelivery(chkBundleDelivery.Value = vbChecked) Then
            If Not blnReprint Then
                roll_back
            End If
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If
    
    If Not blnReprint Then
        If moJob.DigitalSearchesFolder <> "" Then
            msg = ""
            For Each oSearch In moJob.Searches
                If oSearch.PhotoType = "D" Then
                    msg = msg & vbCrLf & "'" & oSearch.DisplayDescription & "', "
                End If
            Next oSearch
            If msg <> "" Then
                msg = Mid(Left(msg, Len(msg) - 2), 3)
                strPreparedFilesMessage = "The following searches have Digital Images that have been copied to folder '" & moJob.DigitalDeliveryNoteFolder & "' in preparation for Delivery:"
                Set fPreparedFiles = New frmPreparedFiles
                fPreparedFiles.lblMessage = strPreparedFilesMessage
                fPreparedFiles.DigitalSearchesFolder = moJob.DigitalDeliveryNoteFolder
                fPreparedFiles.txtSearches = msg
                Screen.MousePointer = vbDefault
                fPreparedFiles.Show vbModal
                Screen.MousePointer = vbHourglass
                Set fPreparedFiles = Nothing
            End If
        End If
        goLog.WriteDetailedLog "Digital Pictures prepared for job - '" & CStr(moJob.JobNo) & "'", "frm_delivery_note_print.cmdPrint_Click"
        '
        goLog.WriteDetailedLog "Delivery note printed for delivery note number - '" & CStr(delivery_note_no) & "'", "frm_delivery_note_print.cmdPrint_Click"
        
        goLog.WriteDetailedLog "Completed delivery of job - '" & CStr(moJob.JobNo) & "'", "frm_delivery_note_print.cmdPrint_Click"
    End If
    
    Set oOutlook = Nothing
refresh_screens:
    '
    '   Refresh all required screens.
    '
    If is_form_loaded("frm_delivery_note_maint") Then
        frm_delivery_note_maint.RefreshDeliveryNotes
    End If

    If is_form_loaded("frm_job_maint") Then
        frm_job_maint.redisplay
    End If

    If is_form_loaded("frm_photograph_view") Then
        frm_photograph_view.RedisplayTransparencies
    End If

    If is_form_loaded("frm_photograph_edit") Then
        frm_photograph_edit.redisplay_images
    End If

print_delivery_note:
    Hide
    Screen.MousePointer = vbDefault
    '
    '********* Code Ends Here **************
    '
    Exit Sub
    '
cmdPrint_Click_Error:
    DisplayError , "frm_delivery_note_print.cmdPrint_Click", vbExclamation
End Sub

Private Function EmailDelivery(Optional ByVal pblnBundle As Boolean = False) As Boolean
    Dim oOutlook            As OutlookClass
    Dim colAttachments      As Collection
    Dim varAttach           As Variant
    Dim strEmail            As String
    Dim oCustomer           As Customer2
    Dim oFSO                As Scripting.FileSystemObject
    Dim oFolder             As Scripting.Folder
    Dim oSearchFolder       As Scripting.Folder
    Dim oFile               As Scripting.File

'***************************************
' Module/Form Name   : frm_delivery_note_print
'
' Procedure Name     : EmailDelivery
'
' Purpose            :
'
' Date Created       : 10/06/2006 17:27
'
' Author             : Gareth
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
    On Error GoTo ErrorProc

    EmailDelivery = False
    
    If goSystemConfig.WebSearchTestEmail = "" Then
        strEmail = moJob.Contact.Email
    Else
        strEmail = goSystemConfig.WebSearchTestEmail
    End If
    
    Set oCustomer = New Customer2
    oCustomer.Read moJob.CustomerNo
    
    Set oOutlook = New OutlookClass
    '
    '   If requested, add bundled attachments.
    '
    If pblnBundle Then
        Set colAttachments = New Collection
        '
        varAttach = moJob.DigitalDeliveryNote
        colAttachments.Add varAttach                 '   Delivery Note
        '
        'varAttach = App.Path & "\Reports\TermsAndConditions.doc"
        varAttach = goSystemConfig.ServerLocation & "\Reports\TermsAndConditions.doc"
        colAttachments.Add varAttach                 '   Terms and Conditions
        '
        Set oFSO = New Scripting.FileSystemObject
        Set oFolder = oFSO.GetFolder(moJob.DigitalDeliveryNoteFolder)
        For Each oSearchFolder In oFolder.SubFolders
            For Each oFile In oSearchFolder.Files
                varAttach = oFile.Path
                colAttachments.Add varAttach         '   Pictures
            Next oFile
        Next oSearchFolder
        '
        varAttach = moJob.DigitalSearchesFolder & "\Printing Clive Nichols.doc"
        colAttachments.Add varAttach                 '   Terms and Conditions
    End If
    
    oOutlook.HTMLEmail strEmail, _
                       oCustomer.CustomerName & "  Job: " & moJob.reference & " (" & moJob.JobNo & ")", _
                       IIf(moJob.WebStatus <> "", moJob.ConfirmationHTMLEmail, moJob.BundledHTMLEmail), _
                       colAttachments

    If Not oOutlook.MailSent Then
        On Error Resume Next
        Me.SetFocus
        MsgBox "Email not sent. Job has not been delivered.", vbExclamation, App.Title
        roll_back
        Set oOutlook = Nothing
        Set colAttachments = Nothing
        Set oCustomer = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    
    goLog.WriteDetailedLog "Email sent for job - '" & CStr(moJob.JobNo) & "'", "frm_delivery_note_print.EmailDelivery"
    
    Set colAttachments = Nothing
    Set oOutlook = Nothing
    Set oCustomer = Nothing
    
    EmailDelivery = True

    On Error GoTo 0
    Exit Function
'
'********* Code Ends Here **************
'
ErrorProc:
    ErrorRaise "frm_delivery_note_print.EmailDelivery"
End Function

Private Sub txt_loss_damage_fee_Change()
   
   Static loss_damage_fee_save As String
    
    If Len(LTrim(txt_loss_damage_fee)) = 0 Then
         loss_damage_fee_save = ""
    Else
        If Not IsNumeric(txt_loss_damage_fee.Text) Then
            txt_loss_damage_fee.Text = loss_damage_fee_save
            txt_loss_damage_fee.SelStart = Len(txt_loss_damage_fee.Text)
            Beep
        Else
            loss_damage_fee_save = txt_loss_damage_fee.Text
        End If
    End If
End Sub

Private Sub txt_service_fee_Change()
   
   Static service_fee_save As String
    
    If Len(LTrim(txt_service_fee)) = 0 Then
         service_fee_save = ""
    Else
        If Not IsNumeric(txt_service_fee.Text) Then
            txt_service_fee.Text = service_fee_save
            txt_service_fee.SelStart = Len(txt_service_fee.Text)
            Beep
        Else
            service_fee_save = txt_service_fee.Text
        End If
    End If
End Sub

Private Sub txt_weekly_holding_fee_Change()
   
   Static weekly_holding_fee_save As String
    
    If Len(LTrim(txt_weekly_holding_fee)) = 0 Then
         weekly_holding_fee_save = ""
    Else
        If Not IsNumeric(txt_weekly_holding_fee.Text) Then
            txt_weekly_holding_fee.Text = weekly_holding_fee_save
            txt_weekly_holding_fee.SelStart = Len(txt_weekly_holding_fee.Text)
            Beep
        Else
            weekly_holding_fee_save = txt_weekly_holding_fee.Text
        End If
    End If
End Sub

