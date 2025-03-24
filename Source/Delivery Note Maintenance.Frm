VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_delivery_note_maint 
   AutoRedraw      =   -1  'True
   Caption         =   "Delivery Note Maintenance"
   ClientHeight    =   5190
   ClientLeft      =   2130
   ClientTop       =   1755
   ClientWidth     =   10965
   Icon            =   "Delivery Note Maintenance.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5190
   ScaleWidth      =   10965
   Begin ImageWhere.SimpleGrid smgDeliveryNote 
      Height          =   975
      Left            =   1000
      TabIndex        =   8
      Top             =   2000
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   1720
      Columns         =   1
      KeyCol          =   0
   End
   Begin VB.CommandButton cmdReprintAll 
      Caption         =   "&Reprint All"
      Height          =   675
      Left            =   8280
      Picture         =   "Delivery Note Maintenance.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   60
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CheckBox chkMultipleSelection 
      Caption         =   "Multiple Selection"
      Height          =   195
      Left            =   7680
      TabIndex        =   10
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdSummary 
      Caption         =   "&Summary"
      Height          =   675
      Left            =   3600
      Picture         =   "Delivery Note Maintenance.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   60
      UseMaskColor    =   -1  'True
      Width           =   1035
   End
   Begin VB.Frame fraDeliveryNote 
      Height          =   3375
      Left            =   60
      TabIndex        =   7
      Top             =   780
      Width           =   7935
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   675
      Left            =   9840
      Picture         =   "Delivery Note Maintenance.frx":0E0E
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   60
      UseMaskColor    =   -1  'True
      Width           =   1035
   End
   Begin VB.CommandButton cmdChase 
      Caption         =   "&Chase"
      Height          =   675
      Left            =   6000
      Picture         =   "Delivery Note Maintenance.frx":14D0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   60
      UseMaskColor    =   -1  'True
      Width           =   1035
   End
   Begin VB.CommandButton cmdOutstanding 
      Caption         =   "&Outstanding"
      Height          =   675
      Left            =   2400
      Picture         =   "Delivery Note Maintenance.frx":1C12
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   60
      UseMaskColor    =   -1  'True
      Width           =   1035
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "&Return"
      Height          =   675
      Left            =   4800
      Picture         =   "Delivery Note Maintenance.frx":22D4
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   60
      UseMaskColor    =   -1  'True
      Width           =   1035
   End
   Begin VB.CommandButton cmdFilter 
      Caption         =   "&Filter"
      Height          =   675
      Left            =   1200
      Picture         =   "Delivery Note Maintenance.frx":2996
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   60
      UseMaskColor    =   -1  'True
      Width           =   1035
   End
   Begin VB.CommandButton cmdReprint 
      Caption         =   "&Reprint"
      Height          =   675
      Left            =   60
      Picture         =   "Delivery Note Maintenance.frx":3058
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   60
      UseMaskColor    =   -1  'True
      Width           =   1035
   End
   Begin MSComctlLib.StatusBar stb_delivery_note 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   4935
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8096
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8096
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
   End
   Begin Crystal.CrystalReport crs_delivery_note 
      Left            =   4080
      Top             =   4320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowBorderStyle=   1
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
End
Attribute VB_Name = "frm_delivery_note_maint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private SQL                             As String
Private abort                           As Boolean
Private mod_delivery_note_no            As Long
Private mod_customer_filter             As Integer
Private mod_photograph_filter           As String
Private mblnOpenDeliveryNotesOnly       As Boolean
Private dyna_delivery_note              As DAO.Recordset
Private mrstDeliveryNote                As DAO.Recordset
Private dyna_search_result              As DAO.Recordset
''Private mblnDontResize                  As Boolean
Private mblnMultipleSelection           As Boolean
Private mblnDeliveryNoteDataClicked     As Boolean
Private moDeliveryNotes                 As DeliveryNotes
Private mblnForceResize                 As Boolean
Private mstrCurrentDNoteNo              As String
Private WithEvents fChaser              As frmChaser
Attribute fChaser.VB_VarHelpID = -1

Public Property Let CurrentDNoteNo(vData As String)
    mstrCurrentDNoteNo = vData
End Property

Public Property Get CurrentDNoteNo() As String
    CurrentDNoteNo = mstrCurrentDNoteNo
End Property

Private Sub chkMultipleSelection_Click()
    mblnMultipleSelection = (chkMultipleSelection = vbChecked)
    EnableButtons
    mblnForceResize = True
    Call Form_Resize
    mblnForceResize = False
End Sub

Private Sub cmdExit_Click()
    Unload Me
    Set frm_delivery_note_maint = Nothing
End Sub

Private Sub close_delivery_recordsets()
    On Error GoTo error_proc
    
    dyna_delivery_note.Close
    dyna_search_result.Close
    
    Exit Sub
    
error_proc:
    Select Case Err.Number
        Case Is = 91
            Resume Next
        Case Is = 3420
            Resume Next
        Case Else
            MsgBox Str(Err.Number) + " - " + Err.Description, vbCritical
        End Select
End Sub

Private Sub get_job_dyna_search_result()
'***************************************
' Module/Form Name   : frm_delivery_note_maint
'
' Procedure Name     : get_job_dyna_search_result
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
On Error GoTo get_job_dyna_search_result_Error
'
'******** Code Starts Here *************
'
    
    SQL = "SELECT  search_result.*,photograph.status " _
        & "FROM delivery_note, job, search, search_result, photograph, " _
        & "delivery_note INNER JOIN job ON delivery_note.delivery_note_no= job.delivery_note_no, " _
        & "job INNER JOIN search ON job.job_no = search.job_no, " _
        & "search INNER JOIN search_result ON search.search_no = search_result.search_no, " _
        & "search_result INNER JOIN photograph ON search_result.photograph_key = photograph.photograph_key " _
        & "WHERE delivery_note.delivery_note_no = " + Str(dyna_delivery_note!delivery_note_no) + ";"
    
    Set dyna_search_result = db.OpenRecordset(SQL, dbOpenDynaset)

'
'********* Code Ends Here **************
'
   Exit Sub
'
get_job_dyna_search_result_Error:
    ErrorRaise "frm_delivery_note_maint.get_job_dyna_search_result"
End Sub
''
''Private Sub deliver_job_search_results()
''      '***************************************
''      ' Module/Form Name   : frm_delivery_note_maint
''      '
''      ' Procedure Name     : deliver_job_search_results
''      '
''      ' Purpose            :
''      '
''      ' Date Created       : 06/12/2002
''      '
''      ' Author             : GARETH SAUNDERS
''      '
''      ' Amendment History  : Date       Author    Description
''      '                    : --------------------------------
''      '
''      '***************************************
''      '
''10    On Error GoTo deliver_job_search_results_Error
''      '
''      '******** Code Starts Here *************
''      '
''
''20        get_job_dyna_search_result
''
''30        If dyna_search_result.RecordCount = 0 Then
''40            Exit Sub
''50        End If
''
''60        dyna_search_result.MoveFirst
''
''70        With dyna_search_result
''80            While Not .EOF
''90                .Edit
''100               If Not !Status = "P" Then
''110                   MsgBox "Photograph " + Str(!batch_no) + "/" + Str(!photograph_no) + _
''                             "is not at Status 'Pending'." + Chr(10) + "Aborting transaction", vbCritical
''120                   close_delivery_recordsets
''130                   ws.Rollback
''140               End If
''150               !Status = "O"
''160               !delivery_note_no = dyna_delivery_note!delivery_note_no
''170               .Update
''180               dyna_search_result.MoveNext
''190           Wend
''200       End With
''
''210       dyna_search_result.Close
''
''      '
''      '********* Code Ends Here **************
''      '
''220      Exit Sub
''      '
''deliver_job_search_results_Error:
''230       ErrorRaise "frm_delivery_note_maint.deliver_job_search_results"
''End Sub

Public Sub list_delivery_notes(delivery_note_no, _
                               customer_filter, _
                               photograph_filter, _
                               abort, _
                               bolPicturesReturned As Boolean, _
                               Optional ByVal pblnOutstandingOnly As Boolean = True, _
                               Optional ByVal pblnFreezeWindow As Boolean = True)
'***************************************
' Module/Form Name   : frm_delivery_note_maint
'
' Procedure Name     : list_delivery_notes
'
' Purpose            :
'
' Date Created       : 09/06/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : delivery_note_no - photograph_filter
'                    : abort - As
'                    : Boolean -
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 09/06/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo list_delivery_notes_Error
'
'******** Code Starts Here *************
'

    Dim pnlX                        As Panel
    Dim lngTotalPicturesDelivered   As Long
    Dim lngTotalOutstandingTrans    As Long
    Dim lngTotalOutstandingDigital  As Long
    Dim oDeliveryNote               As DeliveryNote
    Dim colSelected                 As Collection
    Dim lngDNoteNo                  As Long
    Dim strFirstDeliveryNoteNo      As String
    
    abort = True
    Screen.MousePointer = ccHourglass
    If pblnFreezeWindow Then
        LockWindow Me.hWnd
    End If
    '
    '   Establish the access.
    '
    If delivery_note_no = 0 Then
        GoTo filter_dnote
    End If
    '
select_dnote:

    SQL = "SELECT dnote.delivery_note_no, date_printed, returnbydate, "
    SQL = SQL & "customer.customer_name, customer.customer_no, job.job_reference, "
    SQL = SQL & "(SELECT count(*) FROM search_result, photograph "
    SQL = SQL & "WHERE dnote.delivery_note_no = search_result.delivery_note_no and "
    SQL = SQL & "photograph.photograph_key = search_result.photograph_key and photograph.deleted = false) as total_pictures, "
    SQL = SQL & "Job.Job_No, "
    SQL = SQL & "sum(iif(search_result.photograph_no <> 0 and (search_result.date_returned is null or search_result.date_returned = 0) , 1,0)) as total_outstanding_trans, "
    SQL = SQL & "Sum(iif(search_result.photograph_no = 0 and (search_result.date_returned Is Null or search_result.date_returned =0) and (search_result.digitalstatus is null or search_result.digitalstatus = ''),1,0)) as  total_outstanding_digital, "
    SQL = SQL & "(select top 1 action from chaser where chaser.deliverynoteno = dnote.delivery_note_no and chaser.dateamended = (select max(dateamended) from chaser where chaser.deliverynoteno = dnote.delivery_note_no)) as DNStatus "
    SQL = SQL & "FROM delivery_note as DNote, customer, search_result, job, photograph, Chaser "
    SQL = SQL & "WHERE dnote.delivery_note_no = " & CStr(Trim(delivery_note_no)) & " and customer.customer_no = dnote.customer_no and "
    SQL = SQL & "dnote.delivery_note_no = search_result.delivery_note_no and "
    If pblnOutstandingOnly Then
        SQL = SQL & "(search_result.photograph_no <> 0 and (search_result.date_returned is null or search_result.date_returned = 0)) and "
    End If
    SQL = SQL & "job.delivery_note_no = dnote.delivery_note_no and photograph.photograph_key = search_result.photograph_key and photograph.deleted = false AND "
    SQL = SQL & "Chaser.DeliveryNoteNo = DNote.Delivery_Note_No and Chaser.Action = 'None' and Chaser.NextAction = 'None' "
    SQL = SQL & "GROUP BY dnote.delivery_note_no, customer_name, customer.customer_no, job.job_reference, job_no, date_printed, returnbydate"
    
    Set mrstDeliveryNote = db.OpenRecordset(SQL, dbOpenSnapshot)
    If mrstDeliveryNote.EOF Then
        If Not bolPicturesReturned Then
            MsgBox "Delivery note '" + CStr(delivery_note_no) + "' not found." & IIf(pblnOutstandingOnly, vbCrLf & "Try unchecking '" & frm_delivery_note.chkOpenDeliveryNotesOnly.Caption & "'.", ""), vbExclamation, "Select Delivery Note"
            Screen.MousePointer = ccDefault
            UnlockWindow
            Exit Sub
        End If
    End If
    
    GoTo end_dnote
    '
filter_dnote:
    '
    '   Build a recordset with delivery notes and establish the SQL also.
    '
    delivery_note.read_customers_outstanding_delivery_notes customer_filter, _
                                                            photograph_filter, _
                                                            SQL, _
                                                            pblnOutstandingOnly
    
    Set mrstDeliveryNote = db.OpenRecordset(SQL, dbOpenSnapshot)
    If mrstDeliveryNote.EOF Then
        If Not bolPicturesReturned Then
            MsgBox "No delivery note found with requested filters", vbExclamation, "Filter Delivery Note"
            Screen.MousePointer = ccDefault
            UnlockWindow
            Exit Sub
        End If
    End If
    '
end_dnote:
    '
    '   Set the module level variables to contain the filtering information.
    '
    mod_delivery_note_no = delivery_note_no
    mod_customer_filter = customer_filter
    mod_photograph_filter = photograph_filter
    mblnOpenDeliveryNotesOnly = pblnOutstandingOnly
    '
    '   Update the DbGrid.
    '
    With mrstDeliveryNote
        If Not .EOF Then .MoveFirst
        lngTotalOutstandingTrans = 0
        '
        '   Remember any selected Delivery Notes
        '
        If chkMultipleSelection.Value = vbChecked Then
            If Not moDeliveryNotes Is Nothing Then
                Set colSelected = Nothing
                Set colSelected = New Collection
                For Each oDeliveryNote In moDeliveryNotes
                    If oDeliveryNote.Selected = True Then
                        lngDNoteNo = oDeliveryNote.DNoteNo
                        colSelected.Add lngDNoteNo, CStr(lngDNoteNo)
                    End If
                Next oDeliveryNote
            End If
        End If
        '
        Set moDeliveryNotes = Nothing
        Set moDeliveryNotes = New DeliveryNotes
        While Not .EOF
            lngTotalPicturesDelivered = lngTotalPicturesDelivered + !total_pictures
            lngTotalOutstandingTrans = lngTotalOutstandingTrans + !total_outstanding_trans
            lngTotalOutstandingDigital = lngTotalOutstandingDigital + !total_outstanding_digital
            '
            '   Add to the Delivery Notes collection.
            '
            Set oDeliveryNote = moDeliveryNotes.Add(!delivery_note_no, False)
            oDeliveryNote.CustomerNo = !customer_no
            oDeliveryNote.CustomerName = !customer_name
            oDeliveryNote.JobReference = !job_reference
            oDeliveryNote.JobNo = !job_no
            oDeliveryNote.DatePrinted = !date_printed
            oDeliveryNote.ReturnByDate = !ReturnByDate
            oDeliveryNote.TotalOutstandingTrans = !total_outstanding_trans
            oDeliveryNote.TotalOutstandingDigital = !total_outstanding_digital
            oDeliveryNote.TotalPictures = !total_pictures
            oDeliveryNote.Status = !DNStatus
            '
            '   Is it selected?
            '
            On Error Resume Next
            lngDNoteNo = colSelected.Item(CStr(oDeliveryNote.DNoteNo))
            If Err.Number = 0 Then
                oDeliveryNote.Selected = True
            End If
            On Error GoTo list_delivery_notes_Error
            '
            .MoveNext
        Wend
        '
        '   Display.
        '
        If moDeliveryNotes.Count = 0 Then
            strFirstDeliveryNoteNo = ""
        Else
            strFirstDeliveryNoteNo = moDeliveryNotes.Item(1).DNoteNo
        End If
        
        DisplayDeliveryNotes IIf(moDeliveryNotes.Count = 0, _
                             "", _
                             IIf(mstrCurrentDNoteNo = "", strFirstDeliveryNoteNo, mstrCurrentDNoteNo)), _
                             pblnFreezeWindow:=pblnFreezeWindow
        '
        If Not pblnFreezeWindow Then
            smgDeliveryNote.GetKeyRow IIf(mstrCurrentDNoteNo = "", moDeliveryNotes.Item(1).DNoteNo, mstrCurrentDNoteNo)
        End If
        EnableButtons
    End With
    '
    '   Output the filters to the screen.
    '   Build the Panels up from scratch each time.
    '
    While stb_delivery_note.Panels.Count <> 0
        stb_delivery_note.Panels.Remove 1
    Wend
    '
    If mod_delivery_note_no <> 0 Then
        '
        '   Five Panels.
        '
        Set pnlX = stb_delivery_note.Panels.Add(, , "Delivery Note: " & Str(LTrim(mod_delivery_note_no)))
        pnlX.AutoSize = sbrNoAutoSize
    Else
        If mod_customer_filter <> 0 Then
            Set pnlX = stb_delivery_note.Panels.Add(, , "Customer: " & customer.Read(mod_customer_filter)!customer_name)
            pnlX.AutoSize = sbrNoAutoSize
        End If
        
        If mod_photograph_filter <> "0" Then
            Set pnlX = stb_delivery_note.Panels.Add(, , "Photograph: " & Str(mod_photograph_filter))
            pnlX.AutoSize = sbrNoAutoSize
        End If
    End If
    '
    If stb_delivery_note.Panels.Count = 0 Then
        Set pnlX = stb_delivery_note.Panels.Add(, , "")
    End If
    '
    Set pnlX = stb_delivery_note.Panels.Add(, , "Totals:")
    pnlX.AutoSize = sbrNoAutoSize
    '
    Set pnlX = stb_delivery_note.Panels.Add(, , CStr(lngTotalPicturesDelivered))
    pnlX.AutoSize = sbrNoAutoSize
    pnlX.Alignment = sbrRight
    '
    Set pnlX = stb_delivery_note.Panels.Add(, , CStr(lngTotalOutstandingTrans))
    pnlX.AutoSize = sbrNoAutoSize
    pnlX.Alignment = sbrRight
    '
    Set pnlX = stb_delivery_note.Panels.Add(, , CStr(lngTotalOutstandingDigital))
    pnlX.AutoSize = sbrNoAutoSize
    pnlX.Alignment = sbrRight
    '
    ResizePanels
    '
    DoEvents
    
    abort = False
    
    Screen.MousePointer = ccDefault
    '
    '********* Code Ends Here **************
    '
    Exit Sub
    '
list_delivery_notes_Error:
    ErrorRaise "frm_delivery_note_maint.list_delivery_notes"
End Sub

Private Sub ResizePanels()
'***************************************
' Module/Form Name   : frm_delivery_note_maint
'
' Procedure Name     : ResizePanels
'
' Purpose            :
'
' Date Created       : 01/10/2006 16:54
'
' Author             : Gareth
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
    Dim lngPanelsTotalsPosition As Long
    Dim lngNoTextPanels         As Long
    Const CONST_TOTALS_PANEL_WIDTH = 1000
    
    On Error GoTo ErrorProc
    '
    '   Size the panels.
    '
    lngPanelsTotalsPosition = CLng(smgDeliveryNote.Column(1).Width) + _
                              CLng(smgDeliveryNote.Column(2).Width) + _
                              CLng(smgDeliveryNote.Column(3).Width) + _
                              CLng(smgDeliveryNote.Column(4).Width) + _
                              CLng(smgDeliveryNote.Column(5).Width) + _
                              CLng(smgDeliveryNote.Column(6).Width) + _
                              CLng(smgDeliveryNote.Column(7).Width) + _
                              CLng(smgDeliveryNote.Left) + _
                              CLng(fraDeliveryNote.Left)
    
    lngNoTextPanels = 0
    
    If mod_delivery_note_no <> 0 Then
        lngNoTextPanels = lngNoTextPanels + 1
    Else
        If mod_customer_filter <> 0 Then
            lngNoTextPanels = lngNoTextPanels + 1
        End If
        If mod_photograph_filter <> "0" Then
            lngNoTextPanels = lngNoTextPanels + 1
        End If
    End If
    
    If lngNoTextPanels = 0 Then lngNoTextPanels = lngNoTextPanels + 1
   
    If lngNoTextPanels = 1 Then
        stb_delivery_note.Panels(2).Width = CONST_TOTALS_PANEL_WIDTH
        stb_delivery_note.Panels(1).Width = lngPanelsTotalsPosition - stb_delivery_note.Panels(2).Width
    ElseIf lngNoTextPanels = 2 Then
        stb_delivery_note.Panels(3).Width = CONST_TOTALS_PANEL_WIDTH
        stb_delivery_note.Panels(2).Width = (lngPanelsTotalsPosition - stb_delivery_note.Panels(3).Width) / 2
        stb_delivery_note.Panels(1).Width = stb_delivery_note.Panels(1).Width
    End If
    '
    stb_delivery_note.Panels(lngNoTextPanels + 2).Width = smgDeliveryNote.Column(8).Width
    stb_delivery_note.Panels(lngNoTextPanels + 3).Width = smgDeliveryNote.Column(9).Width
    stb_delivery_note.Panels(lngNoTextPanels + 4).Width = smgDeliveryNote.Column(10).Width

    On Error GoTo 0
    Exit Sub
'
'********* Code Ends Here **************
'
ErrorProc:
    ErrorRaise "frm_delivery_note_maint.ResizePanels"
End Sub

Private Sub DisplayDeliveryNotes(Optional ByVal pstrDNoteNo As String = "", _
                                 Optional ByVal pblnFreezeWindow As Boolean = True)
    
    Dim oDeliveryNote       As DeliveryNote
    Dim intTopRow           As Integer
    
'***************************************
' Module/Form Name   : frm_delivery_note_maint
'
' Procedure Name     : DisplayDeliveryNotes
'
' Purpose            :
'
' Date Created       : 22/05/2006 22:49
'
' Author             : Gareth
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
    On Error GoTo ErrorProc

    If pblnFreezeWindow Then
        LockWindow Me.hWnd
        smgDeliveryNote.Redraw = False
    End If
    
    intTopRow = smgDeliveryNote.TopRow
    smgDeliveryNote.Clear
    For Each oDeliveryNote In moDeliveryNotes
        With oDeliveryNote
            smgDeliveryNote.AddRow False, _
                                   .Selected, _
                                   .DNoteNo, _
                                   .CustomerNo, _
                                   .CustomerName, _
                                   .JobReference, _
                                   .DatePrinted, _
                                   .ReturnByDate, _
                                   .TotalPictures, _
                                   .TotalOutstandingTrans, _
                                   .TotalOutstandingDigital, _
                                   IIf(LCase(.Status) = "none", "", IIf(Len(.Status) > 12, Left(.Status, 9) & "...", .Status))
                                   
        End With
    Next oDeliveryNote

    With smgDeliveryNote
        If pblnFreezeWindow Then
            .ResizeRows
            If pstrDNoteNo <> "" Then
                .KeyCol = 2
                .GetKeyRow pstrDNoteNo
            End If
            On Error Resume Next
            smgDeliveryNote.TopRow = intTopRow
            On Error GoTo ErrorProc
            .Redraw = True
            UnlockWindow
        End If
    End With
    
    On Error GoTo 0
    Exit Sub
'
'********* Code Ends Here **************
'
ErrorProc:
    ErrorRaise "frm_delivery_note_maint.DisplayDeliveryNotes"
End Sub

Private Sub cmdReprintAll_Click()
'***************************************
' Module/Form Name   : frm_delivery_note_maint
'
' Procedure Name     : cmdReprint_Click
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
On Error GoTo cmdReprint_Click_Error
'
'******** Code Starts Here *************
'
    frm_delivery_note_maint.Enabled = False
    Screen.MousePointer = vbHourglass
    
    '
    '   Refresh all the Court Case Jobs.
    '
    Dim oCourtCaseJobs As Jobs
    Dim oJob As Job2
    
    Set oCourtCaseJobs = New Jobs
    oCourtCaseJobs.RefreshCourtCaseJobs (False)
        
    For Each oJob In oCourtCaseJobs
        job.print_delivery_note crs_delivery_note, _
                                oJob.DeliveryNoteNo, _
                                oJob.CustomerNo, _
                                "A", _
                                jdnPDF, _
                                "C:\Data\Image Where\Development\DigitalSearches\All", _
                                abort
    
    Next oJob
    Screen.MousePointer = vbDefault
    
    MsgBox "Delivery Notes Exported", vbInformation
    
    
    If abort Then
        frm_delivery_note_maint.Enabled = True
        Exit Sub
    End If
    
    frm_delivery_note_maint.Enabled = True
'
'********* Code Ends Here **************
'
   Exit Sub
'
cmdReprint_Click_Error:
    DisplayError , "frm_delivery_note_maint.cmdReprint_Click", vbExclamation
    frm_delivery_note_maint.Enabled = True
End Sub

Private Sub cmdSummary_Click()
'***************************************
' Module/Form Name   : frm_delivery_note_maint
'
' Procedure Name     : cmdSummary_Click
'
' Purpose            :
'
' Date Created       : 06/03/2005
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo cmdSummary_Click_Error
'
'******** Code Starts Here *************
'
    Dim strSQL                  As String
    Dim qdfSummary              As QueryDef
    Dim strSelectionFormula     As String
    Dim strTitle                As String
    Dim oCustomer               As New Customer2
    Dim strQueryName            As String
    Dim oReportRequest          As ReportRequest
    Const CONST_SUMMARY_OUTSTANDING_PICTURES_REPORT_REQUEST = "SUMOUTSTANDINGPICS"
    
    Screen.MousePointer = vbHourglass
    crs_delivery_note.Reset
    '
    '   Get the SQL.
    '
    delivery_note.read_customers_outstanding_delivery_notes mod_customer_filter, _
                                                            mod_photograph_filter, _
                                                            strSQL, _
                                                            mblnOpenDeliveryNotesOnly
    '
    '   Create a query and store on the database.
    '
    strQueryName = "SummaryOfOutstandingPictures"
    On Error Resume Next
    Set qdfSummary = db.QueryDefs(strQueryName)
    If Err.Number = 0 Then
        db.QueryDefs.Delete strQueryName
    End If
    '
    On Error GoTo cmdSummary_Click_Error
    Set qdfSummary = db.CreateQueryDef(strQueryName, strSQL)
    '
    '   Set up the Report Request Information.
    '
    strTitle = "Summary of Outstanding Pictures"
    If mod_customer_filter <> 0 Then
        On Error Resume Next
        oCustomer.Read mod_customer_filter
        If Err.Number = 0 Then
            strTitle = strTitle & " - Customer: " & oCustomer.CustomerName
        End If
        On Error GoTo cmdSummary_Click_Error
    End If
    '
    If mod_photograph_filter <> 0 Then
        strTitle = strTitle & ", Photograph: " & CStr(mod_photograph_filter)
    End If
    Set oCustomer = Nothing
    '
    Set oReportRequest = New ReportRequest
    With oReportRequest
        If Not .Read(CONST_SUMMARY_OUTSTANDING_PICTURES_REPORT_REQUEST) Then
            '
            '   Create a request.
            '
            .Name = CONST_SUMMARY_OUTSTANDING_PICTURES_REPORT_REQUEST
            .Parameter1 = Left(strTitle, 100)
            .Create
        End If
        '
        '   Update the request with the details of the report request.
        '
        .Parameter1 = strTitle
        .update
    End With
    Set oReportRequest = Nothing
    '
    '   Run the report against the query.
    '
    With crs_delivery_note
        .Destination = 0
        .WindowState = 2
        .WindowControls = True
        .WindowControlBox = True
        .WindowShowCloseBtn = True
    End With
    '
    '   Set the Record Selection formula.
    '
    With crs_delivery_note
        '
        .WindowTitle = strTitle
        '.ReportFileName = App.Path & "\Reports\Outstanding Summary.rpt"
        .ReportFileName = goSystemConfig.ServerLocation & "\Reports\Outstanding Summary.rpt"
        '
        '   Set location of the database file to the location from
        '   which the user ran the app
        '
        .SelectionFormula = ""
        .SubreportToChange = ""
        .DataFiles(0) = glo_dbname
        .DataFiles(1) = glo_dbname
        .SubreportToChange = "TITLES"
        .DataFiles(0) = glo_dbname
        '   EXECUTE PRINT CALL
        .Action = 1
        'Print/Export a report to a dummy file in the backgroud, the purpose of this section of code
        'is to just release first report from memory:
        '.ReportFileName = App.Path & "\Reports\Dummy2.rpt"
        .ReportFileName = goSystemConfig.ServerLocation & "\Reports\Dummy2.rpt"
        .Destination = crptToFile
        .PrintFileType = crptCrystal
        .PrintFileName = "c:\temp\temp.rpt"
        'To turn off ProgressDialog so the user won't notice a report is printing on the background:
        .ProgressDialog = False
        On Error Resume Next
        .Action = 1
    End With
    
    abort = False
    Screen.MousePointer = vbDefault
'
'********* Code Ends Here **************
'
   Exit Sub
'
cmdSummary_Click_Error:
    DisplayError , "frm_delivery_note_maint.cmdSummary_Click", vbExclamation
End Sub

Private Sub update_loaded_screens()
    If is_form_loaded("frm_delivery_note_maint") Then
        frm_delivery_note_maint.RefreshDeliveryNotes pblnPicturesReturned:=True
    End If
End Sub

Private Sub fChaser_ChaserAdded(ID As Long)
    update_loaded_screens
End Sub

Private Sub Form_Initialize()
    mstrCurrentDNoteNo = ""
End Sub

Private Sub Form_Load()

    Dim intWindowState  As Integer
    
    ''mblnDontResize = False
    mblnMultipleSelection = False
    mblnForceResize = False
    '
    '   Basic Image Where?
    '
    If goSystemConfig.BasicImageWhere Then
        cmdReturn.Visible = False
        cmdChase.Visible = False
    End If
    
    Set smgDeliveryNote.Container = fraDeliveryNote
    With smgDeliveryNote
        .Height = 975
        .Left = 120
        .Top = 1680
        .Width = 7695
    End With
    
    With smgDeliveryNote
        .Columns = 11
        .Column(1).Header = "Selected"
        .ColumnProperty(1).PropertyType = smgPropTypeBooleanGraphic
        .Column(2).Header = "Number"
        .Column(3).Header = "Customer No"
        .Column(4).Header = "Customer Name"
        .Column(5).Header = "Job"
        .Column(6).Header = "Date Printed"
        .Column(7).Header = "Due By"
        .Column(8).Header = "Total pictures delivered"
        .Column(8).Align = flexAlignRightTop
        .Column(9).Header = "Transpar-" & vbCrLf & "encies" & vbCrLf & "o/s"
        .Column(9).Align = flexAlignRightTop
        .Column(10).Header = "Digitial o/s"
        .Column(10).Align = flexAlignRightTop
        .Column(11).Header = "D/N Status"
    End With
    
    intWindowState = GetSetting(App.Title, Me.Name, "Load", vbNormal)
    If intWindowState = vbMinimized Then
        intWindowState = vbNormal
    End If
    '
    Me.WindowState = intWindowState
    
    com_position_form Me

End Sub

Private Sub Form_Resize()
'***************************************
' Module/Form Name   : frm_delivery_note_maint
'
' Procedure Name     : Form_Resize
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
On Error GoTo Form_Resize_Error
'
'******** Code Starts Here *************
'
    Dim MIN_WIDTH               As Integer
    Dim MIN_HEIGHT              As Integer
    Dim header_height           As Integer
    
    MIN_WIDTH = 7500
    MIN_HEIGHT = 4650
    header_height = 440
    
    ''If mblnDontResize Then Exit Sub
    
    If Not mblnForceResize Then
        If Not ResizeForm(Me) Then
            Exit Sub
        End If
    End If
    
    If WindowState = 1 Then
        Exit Sub
    End If
  
    If WindowState = 0 Then
        If Width < MIN_WIDTH Then Width = MIN_WIDTH
        If Height < MIN_HEIGHT Then Height = MIN_HEIGHT
    End If
    
    With fraDeliveryNote
        .Width = frm_delivery_note_maint.Width - _
                 (.Left * 2) - 100
        .Height = frm_delivery_note_maint.Height - _
                  .Top - stb_delivery_note.Height - 500
    End With

    cmdExit.Left = frm_delivery_note_maint.Width - cmdExit.Width - 160
    chkMultipleSelection.Move cmdSummary.Left + cmdSummary.Width + 200, cmdSummary.Top + 200, 1695
'
'   Size the DbGrid.
'
    With smgDeliveryNote
        LockWindow Me.hWnd
        .Redraw = False
        .Move 100, _
              200, _
              fraDeliveryNote.Width - (smgDeliveryNote.Left * 2), _
              fraDeliveryNote.Height - 300
        '
        '   Show an exact number of rows on the screen.
        '
        '.Height = Int((fraDeliveryNote.Height - (.Top * 2) - header_height) / .RowHeight) * .RowHeight + header_height
        If mblnMultipleSelection Then
            .Column(1).Width = 720
        Else
            .Column(1).Width = 0
        End If
        .Column(2).Width = 720
        .Column(3).Width = 0
        .Column(6).Width = 1000
        .Column(7).Width = 1000
        .Column(8).Width = 1000
        .Column(9).Width = 1000
        .Column(10).Width = 1000
        .Column(11).Width = 1000
        .Column(4).Width = (.Width - scroll_bar_width - 100 - _
                                     .Column(1).Width - _
                                     .Column(2).Width - _
                                     .Column(6).Width - _
                                     .Column(7).Width - _
                                     .Column(8).Width - _
                                     .Column(9).Width - _
                                     .Column(10).Width - _
                                     .Column(11).Width) / 2
        .Column(5).Width = .Column(4).Width
        DoEvents
        '
        ResizePanels
        '
        '   Sometime produces a 3009 - invalid row error!
        '
        On Error Resume Next
        .ResizeRows
        On Error GoTo Form_Resize_Error
        .Redraw = True
        UnlockWindow
    End With
'
'********* Code Ends Here **************
'
   Exit Sub
'
Form_Resize_Error:
    DisplayError , "frm_delivery_note_maint.Form_Resize", vbExclamation
End Sub

Private Sub cmdFilter_Click()
'***************************************
' Module/Form Name   : frm_delivery_note_maint
'
' Procedure Name     : cmdFilter_Click
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
On Error GoTo cmdFilter_Click_Error
'
'******** Code Starts Here *************
'
    frm_delivery_note.load_delivery_notes mod_customer_filter, _
                                          mod_photograph_filter, _
                                          mod_delivery_note_no
'
'********* Code Ends Here **************
'
   Exit Sub
'
cmdFilter_Click_Error:
    DisplayError , "frm_delivery_note_maint.cmdFilter_Click", vbExclamation
End Sub

Private Sub cmdOutstanding_Click()
'***************************************
' Module/Form Name   : frm_delivery_note_maint
'
' Procedure Name     : cmdOutstanding_Click
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
On Error GoTo cmdOutstanding_Click_Error
'
'******** Code Starts Here *************
'
    If smgDeliveryNote.CurrentRow = 0 Then
        MsgBox "Select a Delivery Note"
        Exit Sub
    End If
'
'   Set up Crystal Reports parameters.
'
print_delivery_note:
    frm_delivery_note_maint.Enabled = False
    Screen.MousePointer = vbHourglass
'
'   Request an 'outstanding' delivery note.
'
    job.print_delivery_note crs_delivery_note, _
                            smgDeliveryNote.Column(2).Value, _
                            smgDeliveryNote.Column(3).Value, _
                            "O", _
                            jdnOnLine, _
                            "", _
                            abort
                
    Screen.MousePointer = ccDefault
    
    If abort Then
        frm_delivery_note_maint.Enabled = True
        Exit Sub
    End If
    
    frm_delivery_note_maint.Enabled = True
'
'********* Code Ends Here **************
'
   Exit Sub
'
cmdOutstanding_Click_Error:
    DisplayError , "frm_delivery_note_maint.cmdOutstanding_Click", vbExclamation
    frm_delivery_note_maint.Enabled = True
End Sub

Private Sub cmdReprint_Click()
'***************************************
' Module/Form Name   : frm_delivery_note_maint
'
' Procedure Name     : cmdReprint_Click
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
On Error GoTo cmdReprint_Click_Error
'
'******** Code Starts Here *************
'
    Dim oJob        As Job2
    
    If smgDeliveryNote.CurrentRow = 0 Then
        MsgBox "Select a Delivery Note"
        Exit Sub
    End If
    
'
'   Set up Crystal Reports parameters.
'
print_delivery_note:
    '
    frm_delivery_note_maint.Enabled = False
    Screen.MousePointer = vbHourglass
    Set oJob = New Job2
    oJob.Read moDeliveryNotes.Item(CStr(smgDeliveryNote.Column(2).Value)).JobNo
    job.print_delivery_note crs_delivery_note, _
                            smgDeliveryNote.Column(2).Value, _
                            smgDeliveryNote.Column(3).Value, _
                            "A", _
                            jdnOnLine, _
                            oJob.DigitalSearchesFolder, _
                            abort
    Set oJob = Nothing
    Screen.MousePointer = vbDefault
    
    If abort Then
        frm_delivery_note_maint.Enabled = True
        Exit Sub
    End If
    
    frm_delivery_note_maint.Enabled = True
'
'********* Code Ends Here **************
'
   Exit Sub
'
cmdReprint_Click_Error:
    DisplayError , "frm_delivery_note_maint.cmdReprint_Click", vbExclamation
    frm_delivery_note_maint.Enabled = True
End Sub

Private Sub cmdReturn_Click()
'***************************************
' Module/Form Name   : frm_delivery_note_maint
'
' Procedure Name     : cmdReturn_Click
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
On Error GoTo cmdReturn_Click_Error
'
'******** Code Starts Here *************
'
    If smgDeliveryNote.CurrentRow = 0 Then
        MsgBox "Select a Delivery Note"
        Exit Sub
    End If
    
    mstrCurrentDNoteNo = CStr(smgDeliveryNote.Column(2).Value)
    frm_delivery_note_return_photos.load_delivery_note smgDeliveryNote.Column(2).Value, _
                                                       mod_delivery_note_no, _
                                                       mod_customer_filter, _
                                                       mod_photograph_filter, _
                                                       mblnOpenDeliveryNotesOnly
'
'********* Code Ends Here **************
'
   Exit Sub
'
cmdReturn_Click_Error:
    DisplayError , "frm_delivery_note_maint.cmdReturn_Click", vbExclamation
End Sub

Private Sub cmdChase_Click()
'***************************************
' Module/Form Name   : frm_delivery_note_maint
'
' Procedure Name     : cmdChase_Click
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
On Error GoTo cmdChase_Click_Error
'
'******** Code Starts Here *************
'
    Dim fChasers            As frmChasers
    Dim lngDNoteNo          As Long
    Dim oDeliveryNote       As DeliveryNote
    
    lngDNoteNo = smgDeliveryNote.Column(2).Value
    mstrCurrentDNoteNo = CStr(lngDNoteNo)
    If chkMultipleSelection.Value = vbUnchecked Then
        Set fChasers = New frmChasers
        fChasers.Display lngDNoteNo, IIf(smgDeliveryNote.Column(7).Value = 0, False, True)
        Set fChasers = Nothing
    Else
        If Not ValidMultipleChaserSelection Then
            Exit Sub
        End If
        '
        Set fChaser = New frmChaser
        fChaser.DisplayMultiple moDeliveryNotes
        Set fChaser = Nothing
        '
        '   Reset the chasers as requested by Jane in email of 31/10/2006
        '
        For Each oDeliveryNote In moDeliveryNotes
            oDeliveryNote.Selected = False
        Next oDeliveryNote
        
        RefreshDeliveryNotes
    End If
'
'********* Code Ends Here **************
'
    Exit Sub
'
cmdChase_Click_Error:
    DisplayError , "frm_delivery_note_maint.cmdChase_Click", vbExclamation
End Sub

Private Function ValidMultipleChaserSelection()
    Dim blnSelected                 As Boolean
    Dim blnAllSameCustomer          As Boolean
    Dim intLastCustomerSelected     As Integer
    
    Dim oDeliveryNote       As DeliveryNote

    ValidMultipleChaserSelection = False
    '
    '   Check at least one delivery note has been selected and same customer is selected each time.
    '
    blnSelected = False
    blnAllSameCustomer = True
    intLastCustomerSelected = 0
    For Each oDeliveryNote In moDeliveryNotes
        If oDeliveryNote.Selected Then
            blnSelected = True
            '
            If intLastCustomerSelected <> 0 And _
               intLastCustomerSelected <> oDeliveryNote.CustomerNo Then
                blnAllSameCustomer = False
                Exit For
            End If
            intLastCustomerSelected = oDeliveryNote.CustomerNo
        End If
    Next oDeliveryNote
    '
    If Not blnSelected Then
        MsgBox "Please select at least one Delivery Note.", vbExclamation
        Exit Function
    End If
    '
    '   Are all delivery notes for the same customer?
    '
    If Not blnAllSameCustomer Then
        MsgBox "Please select the same customer for each Delivery Note.", vbExclamation
        Exit Function
    End If
    '
    ValidMultipleChaserSelection = True

End Function

Private Sub ResetMultipleChaserSelection()
    Dim blnSelected                 As Boolean
    Dim oDeliveryNote       As DeliveryNote
    
    For Each oDeliveryNote In moDeliveryNotes
        oDeliveryNote.Selected = False
    Next oDeliveryNote
    '
End Sub

Public Sub RefreshDeliveryNotes(Optional pblnPicturesReturned As Boolean = False)
'***************************************
' Module/Form Name   : frm_delivery_note_maint
'
' Procedure Name     : RefreshDeliveryNotes
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
On Error GoTo RefreshDeliveryNotes_Error
'
'******** Code Starts Here *************
'
    Dim blnAbort As Boolean
    
    list_delivery_notes mod_delivery_note_no, _
                        mod_customer_filter, _
                        mod_photograph_filter, _
                        blnAbort, _
                        pblnPicturesReturned, _
                        mblnOpenDeliveryNotesOnly
'
'********* Code Ends Here **************
'
   Exit Sub
'
RefreshDeliveryNotes_Error:
    ErrorRaise "frm_delivery_note_maint.RefreshDeliveryNotes"
End Sub

Public Sub Display()
'***************************************
' Module/Form Name   : frm_delivery_note_maint
'
' Procedure Name     : Display
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
On Error GoTo Display_Error
'
'******** Code Starts Here *************
'
    If Not Me.Visible Then
        ''mblnDontResize = True
        Me.Show
        ''mblnDontResize = False
    End If
    Me.SetFocus
'
'********* Code Ends Here **************
'
   Exit Sub
'
Display_Error:
    ErrorRaise "frm_delivery_note_maint.Display"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    gcolMaxedWindows.Remove CStr(Me.hWnd)
    Set moDeliveryNotes = Nothing
    SaveSetting App.Title, Me.Name, "Load", Me.WindowState
End Sub

Private Sub smgDeliveryNote_Click()
    
'***************************************
' Module/Form Name   : frm_delivery_note_maint
'
' Procedure Name     : smgDeliveryNote_Click
'
' Purpose            :
'
' Date Created       : 30/07/2006 17:51
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
    
    If mblnDeliveryNoteDataClicked Then
        Set oDeliveryNote = moDeliveryNotes.Item(CStr(smgDeliveryNote.Column(2).Value))
        oDeliveryNote.Selected = Not oDeliveryNote.Selected
        With oDeliveryNote
            smgDeliveryNote.Redraw = False
            smgDeliveryNote.UpdateRow smgDeliveryNote.CurrentRow, _
                                      .Selected, _
                                      .DNoteNo, _
                                      .CustomerNo, _
                                      .CustomerName, _
                                      .JobReference, _
                                      .DatePrinted, _
                                      .ReturnByDate, _
                                      .TotalPictures, _
                                      .TotalOutstandingTrans, _
                                      .TotalOutstandingDigital, _
                                      IIf(LCase(.Status) = "none", "", IIf(Len(.Status) > 12, Left(.Status, 9) & "...", .Status))
            smgDeliveryNote.Redraw = True
        End With
    End If

    On Error GoTo 0
    Exit Sub
'
'********* Code Ends Here **************
'
ErrorProc:
    DisplayError , "frm_delivery_note_maint.smgDeliveryNote_Click", vbExclamation
End Sub

Private Sub smgDeliveryNote_DblClick()
'***************************************
' Module/Form Name   : frm_delivery_note_maint
'
' Procedure Name     : smgDeliveryNote_DblClick
'
' Purpose            :
'
' Date Created       : 07/04/2003
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo smgDeliveryNote_DblClick_Error
'
'******** Code Starts Here *************
'
    '
    '   Basic Image Where?
    '
    If goSystemConfig.BasicImageWhere Then
        Exit Sub
    End If

    If mblnMultipleSelection Then
        Exit Sub
    End If
    
    mstrCurrentDNoteNo = CStr(smgDeliveryNote.Column(2).Value)
    frm_delivery_note_return_photos.load_delivery_note smgDeliveryNote.Column(2).Value, _
                                                       mod_delivery_note_no, _
                                                       mod_customer_filter, _
                                                       mod_photograph_filter, _
                                                       mblnOpenDeliveryNotesOnly
'
'********* Code Ends Here **************
'
   Exit Sub
'
smgDeliveryNote_DblClick_Error:
    DisplayError , "frm_delivery_note_maint.smgDeliveryNote_DblClick", vbExclamation
End Sub

Public Sub ForceResize()
    gblnResizeMaxedWindows = False
    Form_Resize
    gblnResizeMaxedWindows = True
End Sub

Private Sub smgDeliveryNote_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim dblDataHeight       As Double
    Dim intRow              As Integer
    Dim intShiftTest        As Integer
    
'***************************************
' Module/Form Name   : frm_delivery_note_maint
'
' Procedure Name     : smgDeliveryNote_MouseUp
'
' Purpose            :
'
' Date Created       : 21/05/2006 18:28
'
' Author             : Gareth
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
    On Error GoTo ErrorProc
    '
    '   Only do this for multiple selection.
    '
    If chkMultipleSelection.Value <> vbChecked Then
        Exit Sub
    End If
    
    For intRow = 1 To smgDeliveryNote.Rows
        dblDataHeight = dblDataHeight + smgDeliveryNote.RowHeight(intRow)
    Next intRow
    
    intShiftTest = Shift And 7
    
    If Y > smgDeliveryNote.HeaderHeight And _
       Y < dblDataHeight + smgDeliveryNote.HeaderHeight And _
       X > 0 And X < smgDeliveryNote.Column(2).Width Then
        mblnDeliveryNoteDataClicked = True
    Else
        mblnDeliveryNoteDataClicked = False
    End If

    On Error GoTo 0
    Exit Sub
'
'********* Code Ends Here **************
'
ErrorProc:
    DisplayError , "frm_delivery_note_maint.smgDeliveryNote_MouseUp", vbExclamation
End Sub

Private Sub smgDeliveryNote_RowChanged(CurrentRow As String)
    EnableButtons
End Sub

Private Sub EnableButtons()
'***************************************
' Module/Form Name   : frm_delivery_note_maint
'
' Procedure Name     : EnableButtons
'
' Purpose            :
'
' Date Created       : 30/07/2006 17:57
'
' Author             : Gareth
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
    On Error GoTo ErrorProc

    If chkMultipleSelection.Value = vbChecked Then
        cmdOutstanding.Enabled = False
        cmdReturn.Enabled = False
        cmdReprint.Enabled = False
    Else
        If smgDeliveryNote.Column(9).Value = 0 Then
            cmdOutstanding.Enabled = False
        Else
            cmdOutstanding.Enabled = True
        End If
        cmdReturn.Enabled = True
        cmdReprint.Enabled = True
    End If

    On Error GoTo 0
    Exit Sub
'
'********* Code Ends Here **************
'
ErrorProc:
    ErrorRaise "frm_delivery_note_maint.EnableButtons"
End Sub

