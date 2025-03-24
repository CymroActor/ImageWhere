VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_job_maint 
   Caption         =   "Job Maintenance - Open Jobs"
   ClientHeight    =   4335
   ClientLeft      =   2280
   ClientTop       =   1770
   ClientWidth     =   8730
   Icon            =   "Job Maintenance.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4335
   ScaleWidth      =   8730
   Begin ImageWhere.SimpleGrid smgJobs 
      Height          =   3135
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   5530
      Columns         =   1
      KeyCol          =   0
   End
   Begin VB.Timer timRefresh 
      Left            =   5760
      Top             =   240
   End
   Begin VB.CommandButton cmdFilter 
      Caption         =   "&Filter"
      Height          =   675
      Left            =   4800
      Picture         =   "Job Maintenance.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton cmdDeliveryNote 
      Caption         =   "&Delivery"
      Height          =   675
      Left            =   3000
      Picture         =   "Job Maintenance.frx":0B04
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   675
      Left            =   300
      Picture         =   "Job Maintenance.frx":11C6
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   675
      Left            =   1200
      Picture         =   "Job Maintenance.frx":1888
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   675
      Left            =   2100
      Picture         =   "Job Maintenance.frx":1F4A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   675
      Left            =   3900
      Picture         =   "Job Maintenance.frx":260C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   675
      Left            =   7440
      Picture         =   "Job Maintenance.frx":2CCE
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin Crystal.CrystalReport crs_job 
      Left            =   6720
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "JOB SHEET"
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
Attribute VB_Name = "frm_job_maint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private abort               As Boolean
Private job_selected        As Boolean
Private moJobs              As Jobs
Private mstrCurrentKeyCol   As String
Private mblnDontResize      As Boolean
Private mvarCustomer        As New Customer2
Private mvarShowAll         As Boolean
Private mvarShowOpenOnly    As Boolean
Private mblnRefreshPrimed   As Boolean

Public Property Get CurrentJobNo() As Long
    CurrentJobNo = moJobs.CurrentJobNo
End Property

Public Property Get DontResize() As Boolean
    DontResize = mblnDontResize
End Property

Public Property Let DontResize(vData As Boolean)
    mblnDontResize = vData
End Property

Public Sub redisplay(Optional pblnRefresh As Boolean = True, _
                     Optional plngJobNo As Long = 0)
    '***************************************
    ' Module/Form Name   : frm_job_maint
    '
    ' Procedure Name     : Redisplay
    '
    ' Purpose            :
    '
    ' Date Created       : 08/04/2002
    '
    ' Author             : GARETH SAUNDERS
    '
    ' Amendment History  : Date       Author    Description
    '                    : --------------------------------
    '                    : 08/04/2002 GARETH SAUNDERS
    '
    '***************************************
    '
On Error GoTo Redisplay_Error
    '
    '******** Code Starts Here *************
    '
    Dim lngJobNo    As Long
    Dim oJob        As Job2
    '
    '   First refresh the data.
    '
    If pblnRefresh Then
        If plngJobNo <> 0 Then
            moJobs.CurrentJobNo = plngJobNo
        End If
        '
        '   If the job won't be shown then change the filter.
        '
        If plngJobNo <> 0 Then
            Set oJob = New Job2
            oJob.Read plngJobNo
            If mvarCustomer.CustomerNo <> 0 Then
                If mvarCustomer.CustomerNo <> oJob.CustomerNo Then
                    mvarCustomer.CustomerNo = 0
                End If
            End If
        End If
        '
        moJobs.Refresh mvarCustomer.CustomerNo, mvarShowOpenOnly
    Else
        moJobs.UpdateCountsWarm
    End If
''    On Error Resume Next
    DisplayJobs pblnResizeRows:=True
    '
    '********* Code Ends Here **************
    '
Exit Sub
    '
Redisplay_Error:
    ErrorRaise "frm_job_maint.Redisplay"
End Sub

Public Sub PendingRedisplay()
    moJobs.PendingRefresh = True
    mblnRefreshPrimed = False
End Sub

Private Sub cmdDelete_Click()
    '***************************************
    ' Module/Form Name   : frm_job_maint
    '
    ' Procedure Name     : cmdDelete_Click
    '
    ' Purpose            :
    '
    ' Date Created       : 08/04/2002
    '
    ' Author             : GARETH SAUNDERS
    '
    ' Amendment History  : Date       Author    Description
    '                    : --------------------------------
    '                    : 09/04/2002 GARETH SAUNDERS
    '
    '***************************************
    '
On Error GoTo cmdDelete_Click_Error
    '
    '******** Code Starts Here *************
    '
DeleteJob
    '
    '********* Code Ends Here **************
    '
Exit Sub
    '
cmdDelete_Click_Error:
DisplayError , "frm_job_maint.cmdDelete_Click", vbExclamation
End Sub

Public Sub DeleteJob()
'***************************************
' Module/Form Name   : frm_job_maint
'
' Procedure Name     : DeleteJob
'
' Purpose            :
'
' Date Created       : 14/04/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 14/04/2002 GARETH SAUNDERS
'
'***************************************
'
    On Error GoTo DeleteJob_Error
    '
    '******** Code Starts Here *************
    '
    Dim oJob As Job2
    If smgJobs.CurrentRow = 0 Then
        MsgBox "Select a Job", vbExclamation, App.Title
        Exit Sub
    End If

    Set oJob = moJobs.Item(smgJobs.Column(smgJobs.KeyCol).Value)
    If MsgBox("Are you certain you wish to Delete Job '" & oJob.reference & "'", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
        Exit Sub
    End If

    On Error Resume Next
    oJob.Delete

    If Err.Number = vbObjectError + 4 Then
        MsgBox Err.Description, vbExclamation, App.Title
        Exit Sub
    ElseIf Err.Number = vbObjectError + 3 Then
        MsgBox Err.Description, vbExclamation, App.Title
        redisplay
        Exit Sub
    ElseIf Err.Number <> 0 Then
        ErrorSave
        On Error GoTo DeleteJob_Error
        ErrorRestore
    Else
        On Error GoTo DeleteJob_Error
    End If
    '
    '   Job has been deleted.
    '
    MsgBox "Job No: " & oJob.JobNo & " (" & oJob.reference & ")" & " has been deleted.", vbInformation, App.Title
    redisplay
    '
    '********* Code Ends Here **************
    '
    Exit Sub
    '
DeleteJob_Error:
    ErrorRaise "frm_job_maint.DeleteJob"
End Sub

Private Sub cmdEdit_Click()
'***************************************
' Module/Form Name   : frm_job_maint
'
' Procedure Name     : cmdEdit_Click
'
' Purpose            :
'
' Date Created       : 14/04/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 14/04/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo cmdEdit_Click_Error
'
'******** Code Starts Here *************
'
    EditJob
'
'********* Code Ends Here **************
'
    Exit Sub
    '
cmdEdit_Click_Error:
    DisplayError , "frm_job_maint.cmdEdit_Click", vbExclamation
    redisplay
End Sub

Public Sub EditJob()
'***************************************
' Module/Form Name   : frm_job_maint
'
' Procedure Name     : EditJob
'
' Purpose            :
'
' Date Created       : 14/04/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 14/04/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo EditJob_Error
'
'******** Code Starts Here *************
'
    If smgJobs.CurrentRow = 0 Then
        MsgBox "Select a Job Sheet", vbExclamation, App.Title
        Exit Sub
    End If

    Screen.MousePointer = vbHourglass
    If Not frm_job_edit.display_job(moJobs.Item(smgJobs.Column(smgJobs.KeyCol).Value).CustomerNo, _
                                smgJobs.Column(smgJobs.KeyCol).Value, _
                                moJobs.Item(smgJobs.Column(smgJobs.KeyCol).Value), _
                                Initialize:=True) Then
        redisplay
    End If
    Screen.MousePointer = vbDefault
    '
    '********* Code Ends Here **************
    '
    Exit Sub
    '
EditJob_Error:
    ErrorRaise "frm_job_maint.EditJob"
End Sub

Private Sub cmdExit_Click()
    Unload Me
    Set frm_job_maint = Nothing
End Sub

Private Sub cmdFilter_Click()
'***************************************
' Module/Form Name   : frm_job_maint
'
' Procedure Name     : cmdFilter_Click
'
' Purpose            :
'
' Date Created       : 19/04/2006 10:54
'
' Author             : Gareth
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
    Dim strCustomerName     As String
    
    On Error GoTo ErrorProc
    
    With frmJobFilter
        Set .customer = mvarCustomer
        .ShowAll = mvarShowAll
        .ShowOpenOnly = mvarShowOpenOnly
        .Display
        If .Cancel Then Exit Sub
        Set mvarCustomer = .customer
        mvarShowAll = .ShowAll
        mvarShowOpenOnly = .ShowOpenOnly
        moJobs.Refresh mvarCustomer.CustomerNo, mvarShowOpenOnly
        '
        Set mvarCustomer = .customer
        DisplayJobs
        '
        SaveSetting App.Title, "JobMaintenance", "FilterShowAll", .ShowAll
        SaveSetting App.Title, "JobMaintenance", "FilterShowOpenOnly", .ShowOpenOnly
        SaveSetting App.Title, "JobMaintenance", "FilterCustomerNo", .customer.CustomerNo
    End With

    
    On Error GoTo 0
    Exit Sub
'
'********* Code Ends Here **************
'
ErrorProc:
    DisplayError , "frm_job_maint.cmdFilter_Click", vbExclamation
End Sub

Private Sub cmdNew_Click()
'***************************************
' Module/Form Name   : frm_job_maint
'
' Procedure Name     : cmdNew_Click
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
On Error GoTo cmdNew_Click_Error
'
'******** Code Starts Here *************
'
    NewJob
'
'********* Code Ends Here **************
'
   Exit Sub
'
cmdNew_Click_Error:
    DisplayError , "frm_job_maint.cmdNew_Click", vbExclamation
End Sub

Public Sub NewJob()
'***************************************
' Module/Form Name   : frm_job_maint
'
' Procedure Name     : NewJob
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
On Error GoTo NewJob_Error
'
'******** Code Starts Here *************
'
    frm_job_details.input_details "C"
'
'********* Code Ends Here **************
'
   Exit Sub
'
NewJob_Error:
    ErrorRaise "frm_job_maint.NewJob"
End Sub

Private Sub Form_Activate()
    mblnDontResize = False
    mblnRefreshPrimed = False
End Sub

Private Sub Form_Initialize()
    mvarShowAll = False
    mvarShowOpenOnly = True
    timRefresh.Interval = 500
End Sub

Private Sub Form_Resize()
'***************************************
' Module/Form Name   : frm_job_maint
'
' Procedure Name     : Form_Resize
'
' Purpose            :
'
' Date Created       : 08/04/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 08/04/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo Form_Resize_Error
'
'******** Code Starts Here *************
'
    Dim MIN_WIDTH As Integer
    Dim MIN_HEIGHT As Integer
    Dim header_height As Integer

    MIN_WIDTH = 9000
    MIN_HEIGHT = 4650
    header_height = 440
''    '
''    '   Resizing Rules for MID Children.
''    '
''    If Me.WindowState <> vbMaximized And _
''       ChildWindowsMaxed Then
''        Exit Sub
''    End If
''    '
''    If Me.WindowState <> vbMaximized And _
''       Not ChildWindowsMaxed And _
''       Not mblnResizedToMax Then
''        Exit Sub
''    End If
''    '
''    If Me.WindowState = vbMaximized Then
''        If mblnResizedToMax Then
''            Exit Sub
''        Else
''            mblnResizedToMax = True
''            gcolMaxedWindows.Add Me, CStr(Me.hWnd)
''        End If
''    Else
''        mblnResizedToMax = False
''        gcolMaxedWindows.Remove CStr(Me.hWnd)
''    End If
''    ResizeMaxedWindows Me
''
    If Not ResizeForm(Me) Then
        Exit Sub
    End If
    '
    On Error Resume Next
    If WindowState = 1 Then
        Exit Sub
    End If
    
    If WindowState = 0 Then
        If Width < MIN_WIDTH Then
            Width = MIN_WIDTH
            Exit Sub
        End If
        If Height < MIN_HEIGHT Then
            Height = MIN_HEIGHT
            Exit Sub
        End If
    End If
    '
''    MakeAllControlsInvisible Me
    '
    cmdNew.Move 100, 100
    cmdEdit.Move cmdNew.Left + cmdNew.Width + 100, cmdNew.Top
    cmdDelete.Move cmdEdit.Left + cmdEdit.Width + 100, cmdNew.Top
    cmdDeliveryNote.Move cmdDelete.Left + cmdDelete.Width + 100, cmdNew.Top
    cmdPrint.Move cmdDeliveryNote.Left + cmdDeliveryNote.Width + 100, cmdNew.Top
    cmdFilter.Move cmdPrint.Left + cmdPrint.Width + 100, cmdNew.Top
    cmdExit.Left = frm_job_maint.Width - cmdExit.Width - 200
    '
    '   Size the DbGrid.
    '
    With smgJobs
        .Redraw = False
        .Move 100, _
              cmdExit.Top + cmdExit.Height + 100, _
              frm_job_maint.Width - 300, _
              frm_job_maint.Height - (cmdExit.Top + cmdExit.Height + 100) - 600
          '
          '           Show an exact number of rows on the screen.
          '
          '        .Height = Int((fraJob.Height - (.Top * 2) - header_height) / .RowHeight) * _
          '                    .RowHeight + header_height
        .Column(1).Width = 0
        .Column(2).Width = 700
        .Column(5).Width = 1200
        .Column(6).Width = 800
        .Column(7).Width = 800
        .Column(8).Width = 800
        .Column(9).Width = 600
        .Column(10).Width = 1200
        .Column(11).Width = 0
        .Column(3).Width = (.Width - scroll_bar_width - _
                                      .Column(2).Width - _
                                      .Column(5).Width - _
                                      .Column(6).Width - _
                                      .Column(7).Width - _
                                      .Column(8).Width - _
                                      .Column(9).Width - _
                                      .Column(10).Width - _
                                      .Column(11).Width) / 2
        .Column(4).Width = .Column(3).Width
        .Redraw = True
    End With
    smgJobs.ResizeRows
''    MakeAllControlsVisible Me
    '
    '********* Code Ends Here **************
    '
    Exit Sub
    '
Form_Resize_Error:
    DisplayError , "frm_job_maint.Form_Resize", vbExclamation
End Sub

Private Sub cmdDeliveryNote_Click()
'***************************************
' Module/Form Name   : frm_job_maint
'
' Procedure Name     : cmdDeliveryNote_Click
'
' Purpose            :
'
' Date Created       : 08/04/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 08/04/2002 GARETH SAUNDERS
'                    : 05/07/2005 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo cmdDeliveryNote_Click_Error
'
'******** Code Starts Here *************
'
    Dim oJob                As Job2
    Dim oActivity           As Activity
    Dim strRefineMessage    As String
    Dim strBurnMessage      As String
    Dim oSearch             As Search2
    
    If smgJobs.CurrentRow = 0 Then
        MsgBox "Select a Job Sheet", vbExclamation, App.Title
        Exit Sub
    End If
    '
    Set oJob = moJobs.Item(smgJobs.Column(smgJobs.KeyCol).Value)
    '
    '   Are there any outstanding Activities?
    '
    With oJob.Activities
        .ActivityType = ""
        .JobNo = oJob.JobNo
        .Refresh
        If .Count <> 0 Then
            strRefineMessage = ""
            strBurnMessage = ""
            For Each oActivity In oJob.Activities
                Select Case oActivity.ActivityType
                Case Is = "REFS"
                    On Error Resume Next
                    Set oSearch = oJob.Searches.Item(CStr(oActivity.SearchNo))
                    If Err.Number = 0 Then
                        On Error GoTo cmdDeliveryNote_Click_Error
                        strRefineMessage = strRefineMessage & vbCrLf & "'" & oSearch.DisplayDescription & "', "
                    End If
                    On Error GoTo cmdDeliveryNote_Click_Error
                End Select
            Next oActivity
            If strRefineMessage <> "" Then
                strRefineMessage = Left(strRefineMessage, Len(strRefineMessage) - 2)
                MsgBox "You have not yet confirmed your digital images for the following Searches: " & vbCrLf & _
                       strRefineMessage, vbExclamation
                Exit Sub
            End If
        End If
    End With
    '
    '   Is this a reprint?
    '
    If moJobs.Item(CStr(moJobs.CurrentJobNo)).DeliveryNoteNo > 0 Then
        If MsgBox("This job has already been delivered. Do you wish to reprint the paperwork?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            Exit Sub
        End If
    End If
    '
    If moJobs.Item(smgJobs.Column(smgJobs.KeyCol).Value).NoConfirmed = 0 Then
        MsgBox "There are no confirmed photographs to deliver yet!", vbInformation
    Else
        frm_delivery_note_print.display_customer_and_job smgJobs.Column(smgJobs.KeyCol).Value
    End If
'
'********* Code Ends Here **************
'
    Exit Sub
    '
cmdDeliveryNote_Click_Error:
    DisplayError , "frm_job_maint.cmdDeliveryNote_Click", vbExclamation
End Sub

Private Sub cmdPrint_Click()
    '***************************************
    ' Module/Form Name   : frm_job_maint
    '
    ' Procedure Name     : cmdPrint_Click
    '
    ' Purpose            :
    '
    ' Date Created       : 08/04/2002
    '
    ' Author             : GARETH SAUNDERS
    '
    ' Amendment History  : Date       Author    Description
    '                    : --------------------------------
    '                    : 08/04/2002 GARETH SAUNDERS
    '
    '***************************************
    '
On Error GoTo cmdPrint_Click_Error
    '
    '******** Code Starts Here *************
    '

If smgJobs.CurrentRow = 0 Then
  MsgBox "Select a Job Sheet", vbExclamation, App.Title
  Exit Sub
End If

job.print_job_sheet crs_job, _
                  smgJobs.Column(smgJobs.KeyCol).Value, _
                  abort

    '
    '********* Code Ends Here **************
    '
Exit Sub
    '
cmdPrint_Click_Error:
    DisplayError , "frm_job_maint.cmdPrint_Click", vbExclamation
End Sub

Private Sub DisplayJobs(Optional pblnResizeRows As Boolean = True)
'***************************************
' Module/Form Name   : frm_job_maint
'
' Procedure Name     : DisplayJobs
'
' Purpose            :
'
' Date Created       : 08/04/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 08/04/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo DisplayJobs_Error
'
'******** Code Starts Here *************
'
    Dim oJob            As Job2
    Dim strCaption      As String
    '
    '   Set up form caption.
    '
    strCaption = "Job Maintenance"
    If mvarShowAll = True Then
        strCaption = strCaption & " - All Jobs"
    ElseIf mvarShowOpenOnly Then
        strCaption = strCaption & " - Open Jobs"
    End If
    '
    If mvarCustomer.CustomerNo <> 0 Then
        strCaption = strCaption & " for Customer '" & mvarCustomer.CustomerName & "'"
    End If
    '
    Me.Caption = strCaption
    
    With smgJobs
        LockWindow Me.hWnd
        .Redraw = False
        .Clear
        For Each oJob In moJobs
            .AddRow False, _
                    "", _
                    oJob.JobNo, _
                    oJob.DeliveryCustomerName, _
                    oJob.reference, _
                    oJob.DateRequired, _
                    oJob.Searches.Count, _
                    oJob.NoPhotos, _
                    oJob.NoConfirmed, _
                    oJob.WebStatus, _
                    oJob.WebDatePostedFormatted, _
                    IIf(oJob.DeliveryNoteNo = 0, "", oJob.DeliveryNoteNo)
            '
            '   Jobs that have been posted to the Web and the Customer has not supplied
            '   any feedback within 5 days will be highlighted in bold.
            '
            If oJob.WebDatePosted <> 0 Then
                If DateDiff("d", oJob.WebDatePosted, Now) > goSystemConfig.WebNonFeedbackAlertPeriod Then
                    .RowBold(.CurrentRow) = True
                End If
            End If
        Next oJob
        If pblnResizeRows Then
            On Error Resume Next
            .ResizeRows
            On Error GoTo DisplayJobs_Error
        End If
        .Column(smgJobs.KeyCol).Sorted = smgAscending
        .GetKeyRow (moJobs.CurrentJobNo)
        If smgJobs.Rows > 0 Then
            EnableButtons
        End If
        .Redraw = True
        UnlockWindow
    End With
'
'********* Code Ends Here **************
'
    Exit Sub
'
DisplayJobs_Error:
    ErrorRaise "frm_job_maint.DisplayJobs"
End Sub

Private Sub EnableButtons()
    Dim oJob        As Job2
    
'***************************************
' Module/Form Name   : frm_job_maint
'
' Procedure Name     : EnableButtons
'
' Purpose            :
'
' Date Created       : 27/10/2006 00:06
'
' Author             : Gareth
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
    On Error GoTo ErrorProc

    Set oJob = moJobs.Item(CStr(moJobs.CurrentJobNo))
    If oJob.DeliveryNoteNo > 0 Then
        If oJob.WebStatus = "" Then
            cmdDeliveryNote.Enabled = False
        Else
            cmdDeliveryNote.Enabled = True
        End If
        cmdDelete.Enabled = False
    Else
        cmdDelete.Enabled = True
        cmdDeliveryNote.Enabled = True
    End If
    
    Set oJob = Nothing

    On Error GoTo 0
    Exit Sub
'
'********* Code Ends Here **************
'
ErrorProc:
    ErrorRaise "frm_job_maint.EnableButtons"
End Sub

Public Sub Display(Optional JobNo As Variant)
'***************************************
' Module/Form Name   : frm_job_maint
'
' Procedure Name     : Display
'
' Purpose            :
'
' Date Created       : 11/04/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 11/04/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo Display_Error
'
'******** Code Starts Here *************
'
    Dim intCustomerNo       As Integer
    
    job_selected = False
    '
    '   Set up the grid.
    '
    With smgJobs
      .Columns = 11
      .KeyCol = 2
      .Column(2).Header = "Job No."
      .Column(2).Align = flexAlignLeftBottom
''      .Column(2).Header = "Customer"
      .Column(3).Header = "Delivery Address"
      .Column(3).Align = flexAlignLeftBottom
      .Column(4).Header = "Job Reference"
      .Column(4).Align = flexAlignLeftBottom
      .Column(5).Header = "Date Required"
      .Column(5).Align = flexAlignCenterBottom
      .Column(6).Header = "No. of Searches"
      .Column(6).Align = flexAlignRightBottom
      .Column(7).Header = "No. of Photos"
      .Column(7).Align = flexAlignRightBottom
      .Column(8).Header = "No. Confirmed"
      .Column(8).Align = flexAlignRightBottom
      .Column(9).Header = "Web Status"
      .Column(9).Align = flexAlignLeftBottom
      .Column(10).Header = "Date Posted"
      .Column(10).Align = flexAlignCenterBottom
      .Column(11).Header = "DNote"
      .Column(11).Align = flexAlignCenterBottom
    End With
    Me.WindowState = vbMaximized
    Form_Resize
    Me.Show
    DoEvents
    '
    '   Get Filter Settings
    '
    '
    '   Reset the job filter - requested by Jane in email of 31/10/2006
    '
''    mvarShowAll = GetSetting(App.Title, "JobMaintenance", "FilterShowAll", False)
''    mvarShowOpenOnly = GetSetting(App.Title, "JobMaintenance", "FilterShowOpenOnly", True)
''    intCustomerNo = GetSetting(App.Title, "JobMaintenance", "FilterCustomerNo", 0)
    mvarShowAll = False
    mvarShowOpenOnly = True
    intCustomerNo = 0
    If intCustomerNo = 0 Then
        Set mvarCustomer = Nothing
        Set mvarCustomer = New Customer2
    Else
        mvarCustomer.Read intCustomerNo
    End If
    '
    '   Get the jobs.
    '
    Set moJobs = New Jobs
    moJobs.Refresh mvarCustomer.CustomerNo, mvarShowOpenOnly
    If Not IsMissing(JobNo) Then
      moJobs.CurrentJobNo = JobNo
    End If
    DisplayJobs
    DoEvents
    com_position_form Me
    '
    On Error Resume Next
    If moJobs.Count = 0 Then
        cmdNew.SetFocus
    Else
        smgJobs.SetFocus
    End If
    On Error GoTo Display_Error
    '
    '********* Code Ends Here **************
    '
    Exit Sub
        '
Display_Error:
    DisplayError , "frm_job_maint.Display", vbExclamation
End Sub

Private Sub Form_Unload(Cancel As Integer)
''    Resize.Unhook Me.hWnd
    UnlockWindow
    On Error Resume Next
    gcolMaxedWindows.Remove CStr(Me.hWnd)
End Sub

Private Sub smgJobs_DblClick()
    '***************************************
    ' Module/Form Name   : frm_job_maint
    '
    ' Procedure Name     : smgJobs_DblClick
    '
    ' Purpose            :
    '
    ' Date Created       : 14/04/2002
    '
    ' Author             : GARETH SAUNDERS
    '
    ' Amendment History  : Date       Author    Description
    '                    : --------------------------------
    '                    : 14/04/2002 GARETH SAUNDERS
    '
    '***************************************
    '
On Error GoTo smgJobs_DblClick_Error
    '
    '******** Code Starts Here *************
    '
    If mstrCurrentKeyCol <> "0" Then
        EditJob
    End If
    '
    '********* Code Ends Here **************
    '
Exit Sub
    '
smgJobs_DblClick_Error:
    DisplayError , "frm_job_maint.smgJobs_DblClick", vbExclamation
End Sub

Private Sub smgJobs_MouseHover(CurrentKeyCol As String)
    mstrCurrentKeyCol = CurrentKeyCol
End Sub

Private Sub smgJobs_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'***************************************
' Module/Form Name   : frm_job_maint
'
' Procedure Name     : smgJobs_MouseUp
'
' Purpose            :
'
' Date Created       : 07/12/2002
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
On Error GoTo smgJobs_MouseUp_Error
'
'******** Code Starts Here *************
'
'
'   If right mouse button clicked and an item selected, display pop up menu.
'
    If Button <> vbRightButton Then
        Exit Sub
    End If
'
'   Display pop up menu.
'
    With mdi_npls
        Set .fPopUp = Me
        PopupMenu .mnuJobs, _
                  vbPopupMenuRightButton, _
                  , _
                  , _
                  .mnuJobsEdit
    End With
'
'********* Code Ends Here **************
'
   Exit Sub
'
smgJobs_MouseUp_Error:
    DisplayError , "frm_job_maint.smgJobs_MouseUp", vbExclamation
End Sub

Private Sub smgJobs_RowChanged(CurrentRow As String)
    moJobs.CurrentJobNo = smgJobs.Column(smgJobs.KeyCol).Value
    EnableButtons
End Sub

Public Sub ForceResize()
    gblnResizeMaxedWindows = False
    Form_Resize
    gblnResizeMaxedWindows = True
End Sub


Private Sub timRefresh_Timer()
'***************************************
' Module/Form Name   : frm_job_maint
'
' Procedure Name     : timRefresh_Timer
'
' Purpose            :
'
' Date Created       : 07/07/2004
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo timRefresh_Timer_Error
'
'******** Code Starts Here *************
'
    If moJobs Is Nothing Then
        Exit Sub
    End If
    '
    If moJobs.PendingRefresh Then
        If mblnRefreshPrimed Then
            mblnRefreshPrimed = False
            moJobs.PendingRefresh = False
            redisplay pblnRefresh:=True
        Else
            mblnRefreshPrimed = True
        End If
    End If
'
'********* Code Ends Here **************
'
   Exit Sub
'
timRefresh_Timer_Error:
    DisplayError , "frm_job_maint.timRefresh_Timer", vbExclamation
End Sub
