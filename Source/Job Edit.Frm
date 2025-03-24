VERSION 5.00
Begin VB.Form frm_job_edit 
   Caption         =   "Edit Job"
   ClientHeight    =   6150
   ClientLeft      =   3075
   ClientTop       =   1680
   ClientWidth     =   7335
   Icon            =   "Job Edit.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6150
   ScaleWidth      =   7335
   WindowState     =   2  'Maximized
   Begin ImageWhere.SimpleGrid smgSearches 
      Height          =   1335
      Left            =   1320
      TabIndex        =   21
      Top             =   3960
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   2355
      Columns         =   1
      KeyCol          =   0
   End
   Begin VB.Frame fraSearches 
      Caption         =   "Searches"
      ForeColor       =   &H80000001&
      Height          =   3615
      Left            =   0
      TabIndex        =   15
      Top             =   2400
      Width           =   7335
      Begin VB.Timer timRefresh 
         Interval        =   1000
         Left            =   4920
         Top             =   360
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "&View"
         Height          =   675
         Left            =   3720
         Picture         =   "Job Edit.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   675
         Left            =   6480
         Picture         =   "Job Edit.frx":0B04
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   675
         Left            =   2760
         Picture         =   "Job Edit.frx":11C6
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   675
         Left            =   1860
         Picture         =   "Job Edit.frx":1888
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         Height          =   675
         Left            =   960
         Picture         =   "Job Edit.frx":1F4A
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   735
      End
   End
   Begin VB.Frame fraJob 
      Caption         =   "Job"
      ForeColor       =   &H80000001&
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      Begin VB.CommandButton cmdJobNotes 
         Caption         =   "J&ob Notes"
         Height          =   435
         Left            =   5640
         TabIndex        =   12
         Top             =   1680
         Width           =   1515
      End
      Begin VB.TextBox txtWebStatus 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000001&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1920
         Width           =   3915
      End
      Begin VB.CommandButton cmdReceiveFeedback 
         Caption         =   "&Receive Feedback"
         Height          =   435
         Left            =   5640
         TabIndex        =   11
         Top             =   1200
         Width           =   1515
      End
      Begin VB.CommandButton cmdPostToWeb 
         Caption         =   "&Post to Web"
         Height          =   435
         Left            =   5640
         TabIndex        =   7
         Top             =   720
         Width           =   1515
      End
      Begin VB.TextBox txtJobNo 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000001&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   3915
      End
      Begin VB.CommandButton com_job_details 
         Caption         =   "&Job Details..."
         Height          =   435
         Left            =   5640
         TabIndex        =   3
         Top             =   240
         Width           =   1515
      End
      Begin VB.Label lblWebStatus 
         Caption         =   "Web Status:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label lblJobNo 
         Caption         =   "Job No:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lab_customer_name_lab 
         Caption         =   "Customer Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   540
         Width           =   1215
      End
      Begin VB.Label lab_customer_name 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000001&
         Height          =   255
         Left            =   1560
         TabIndex        =   5
         Top             =   540
         Width           =   3915
      End
      Begin VB.Label lab_address_line_1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000001&
         Height          =   255
         Left            =   1560
         TabIndex        =   8
         Top             =   840
         Width           =   3915
      End
      Begin VB.Label lab_job_reference 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000001&
         Height          =   615
         Left            =   1560
         TabIndex        =   10
         Top             =   1140
         Width           =   3915
      End
      Begin VB.Label lab_job_reference_lab 
         Caption         =   "Job Reference:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1140
         Width           =   1215
      End
      Begin VB.Label lab_address_line_1_lab 
         Caption         =   "Address Line 1:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   1155
      End
   End
End
Attribute VB_Name = "frm_job_edit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SQL As String
Private mod_customer_no As Integer
Private mod_job_no As Long
Private mod_current_search As Long

Public snap_customer As DAO.Recordset
Public snap_job As DAO.Recordset
Private moCustomer As Customer2
Private moJob As Job2
Private moCurrentSearch As Search2
Private moContact As Contact
'
'   Variables to control Event firing.
'
Private mblnAllowResize As Boolean
'
'   Used to control the visibility of controls.
'
Public colInvisibleControls As Collection
Private WithEvents fSearchDetails As frm_search_details
Attribute fSearchDetails.VB_VarHelpID = -1
Private mblnRefreshPrimed       As Boolean
    
Public Function display_job(customer_no As Integer, _
                            job_no As Long, _
                            poJob As Job2, _
                            Optional search_no As Variant, _
                            Optional Initialize As Boolean, _
                            Optional blnSetFocus As Boolean = True) As Boolean
'***************************************
' Module/Form Name   : frm_job_edit
'
' Procedure Name     : display_job
'
' Purpose            :
'
' Date Created       : 14/04/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : customer_no - Integer
'                    : job_no - Long
'                    : search_no - Variant
'                    : Initialize - Boolean
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 14/04/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo display_job_Error
'
'******** Code Starts Here *************
'
    Dim counter As Integer
    Dim oSearches As Searches
    Dim oSearch As Search2
    Dim intNoPhotographs    As Integer
    
    '    Me.Show
    display_job = False
    mod_customer_no = customer_no
    mod_job_no = job_no
    
    If Not IsMissing(Initialize) Then
        If Initialize Then
            mod_current_search = -1
        End If
    End If
    '
    '   Retrieve Job.
    '
''    Set moJob = Nothing
''    Set moJob = New Job2
''    On Error Resume Next
''    moJob.Read mod_job_no
''    If Err.Number = vbObjectError + 10 Then
''        MsgBox "Job " & CStr(mod_job_no) & " does not exist. It may have been deleted by another user.", vbExclamation, App.Title
''        Exit Function
''    ElseIf Err.Number = 0 Then
''        On Error GoTo display_job_Error
''    Else
''        ErrorSave
''        On Error GoTo display_job_Error
''        ErrorRestore
''    End If
    Set moJob = poJob
    '
    '   Retrieve Customer.
    '
    Set moCustomer = New Customer2
    moCustomer.Read customer_no
    '
    '   Set up form title for this screen.
    '
    Me.Caption = "Edit Job '" & moJob.reference & "'"
    '
    '   Determine which buttons to show.
    '
    EnableButtons
    '
    '   Display fields
    '
    txtJobNo.Text = moJob.JobNo
    lab_customer_name.Caption = Replace(moCustomer.CustomerName, "&", "&&")
    lab_address_line_1.Caption = Replace(moCustomer.Address1, "&", "&&")
    
    With snap_job
        lab_job_reference.Caption = Replace(moJob.reference, "&", "&&")
    End With
    '
    '   Web Status.
    '
    If moJob.DeliveryNoteNo = 0 Then
        If Not moContact Is Nothing Then
            If moContact.WebUser Then
                If moJob.WebStatus = "P" Then
                    txtWebStatus.Text = "Posted To Web - " & Format(moJob.WebDatePosted, "DD/MM/YYYY")
                Else
                    If moJob.WebDatePosted = 0 Then
                        txtWebStatus.Text = "Waiting to be posted to Web"
                    Else
                        txtWebStatus.Text = "Web feedback received from customer"
                    End If
                End If
            End If
        Else
            txtWebStatus.Text = ""
        End If
    End If
    
    With smgSearches
        .Redraw = False
        .Clear
        For Each oSearch In moJob.Searches
            If Initialize Then
                intNoPhotographs = oSearch.NoPhotographs
            Else
                intNoPhotographs = oSearch.SearchResults.Count
            End If
            .AddRow False, _
                    CStr(oSearch.SearchNo), _
                    Trim(oSearch.DisplayDescription), _
                    CStr(intNoPhotographs), _
                    oSearch.PhotoTypeDescription
        Next oSearch
        '
        '   Sometime produces a 3009 - invalid row error!
        '
        On Error Resume Next
        .ResizeRows
        On Error GoTo display_job_Error
        '
        '   If the search_no is passed to this routine, then locate
        '   the current search in the listview.
        '
        If Not IsMissing(search_no) Then
            smgSearches.GetKeyRow CStr(search_no)
        Else
            smgSearches.CurrentRow = 1
        End If
        .Redraw = True
    End With
    '
    If Me.Visible = False Then
'        frm_job_maint.DontResize = True
        MakeAllControlsInvisible Me
        Me.Show
        If goSystemConfig.BasicImageWhere Then
            Set colInvisibleControls = New Collection
            colInvisibleControls.Add cmdPostToWeb
            colInvisibleControls.Add cmdReceiveFeedback
        Else
            Set colInvisibleControls = New Collection
        End If

        MakeAllControlsVisible Me, colInvisibleControls
'        frm_job_maint.DontResize = False
    End If
    DoEvents
    If blnSetFocus Then
        If mdi_npls.ActiveForm.Name <> Me.Name Then
            On Error Resume Next
            Me.SetFocus
            On Error GoTo display_job_Error
        End If
    End If
    display_job = True
    '
    '********* Code Ends Here **************
    '
    Exit Function
    '
display_job_Error:
    ErrorRaise "frm_job_edit.display_job"
End Function

Public Sub RefreshJob(Optional blnSetFocus As Boolean = True)
'***************************************
' Module/Form Name   : frm_job_edit
'
' Procedure Name     : RefreshJob
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
On Error GoTo RefreshJob_Error
'
'******** Code Starts Here *************
'
    display_job mod_customer_no, mod_job_no, moJob, search_no:=mod_current_search, blnSetFocus:=blnSetFocus
'
'********* Code Ends Here **************
'
   Exit Sub
'
RefreshJob_Error:
    ErrorRaise "frm_job_edit.RefreshJob"
End Sub

Private Sub cmdDelete_Click()
'***************************************
' Module/Form Name   : frm_job_edit
'
' Procedure Name     : cmdDelete_Click
'
' Purpose            :
'
' Date Created       : 23/04/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 23/04/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo cmdDelete_Click_Error
'
'******** Code Starts Here *************
'
    If smgSearches.CurrentRow = 0 Then
        MsgBox "Select a Search to Delete", vbExclamation, App.Title
        Exit Sub
    End If

    If moJob.Searches.Item(CStr(smgSearches.Column(1).Value)).SearchResults.Count <> 0 Then
        MsgBox "Search has photographs!" & vbCrLf & "These must be removed first before deleting the Search", vbExclamation, "Delete Search"
        Exit Sub
    End If

    If MsgBox("Are you sure you wish to delete Search '" & smgSearches.Column(2).Value & "'", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
        Exit Sub
    End If

    On Error Resume Next
    moJob.Searches.Delete smgSearches.Column(1).Value
    If Err.Number = vbObjectError + 1 Then
        MsgBox Err.Description, vbExclamation, App.Title
        moJob.Searches.Refresh
    ElseIf Err.Number <> 0 Then
        ErrorSave
        On Error GoTo cmdDelete_Click_Error
        ErrorRestore
    Else
        On Error GoTo cmdDelete_Click_Error
        If smgSearches.CurrentRow > 1 Then
            LockWindow Me.hWnd
            smgSearches.CurrentRow = smgSearches.CurrentRow - 1
            mod_current_search = smgSearches.Column(1).Value
        End If
    End If
    '
    update_loaded_screens
    '
    '   Update screen to show that search has been deleted, having first decremented the
    '   current search indicator.
    '
    LockWindow Me.hWnd
    frm_job_edit.display_job moJob.CustomerNo, moJob.JobNo, moJob, mod_current_search
    UnlockWindow
'
'********* Code Ends Here **************
'
    Exit Sub
    '
cmdDelete_Click_Error:
    DisplayError , "frm_job_edit.cmdDelete_Click", vbExclamation
End Sub

Private Sub cmdEdit_Click()
'***************************************
' Module/Form Name   : frm_job_edit
'
' Procedure Name     : cmdEdit_Click
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
On Error GoTo cmdEdit_Click_Error
'
'******** Code Starts Here *************
'
    EditSearch
'
'********* Code Ends Here **************
'
   Exit Sub
'
cmdEdit_Click_Error:
    DisplayError , "frm_job_edit.cmdEdit_Click", vbExclamation
End Sub

Private Sub EditSearch()
'***************************************
' Module/Form Name   : frm_job_edit
'
' Procedure Name     : EditSearch
'
' Purpose            :
'
' Date Created       : 09/06/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 09/06/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo EditSearch_Error
'
'******** Code Starts Here *************
'
    Dim fSearch As frm_search

    If smgSearches.CurrentRow = 0 Then
        MsgBox "Select a Search to Edit", vbExclamation, App.Title
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    Set fSearch = New frm_search
    fSearch.display_search_form Edit, _
                                moJob, _
                                moJob.Searches.Item(smgSearches.Column(1).Value).SearchNo

    Set fSearch = Nothing
    Screen.MousePointer = vbDefault
    '
    '********* Code Ends Here **************
    '
    Exit Sub
    '
EditSearch_Error:
    ErrorRaise "frm_job_edit.EditSearch"
End Sub

Private Sub cmdJobNotes_Click()
    Dim fJobNotes      As frmJobNotes

    Set fJobNotes = New frmJobNotes
    fJobNotes.Display moJob
    Set fJobNotes = Nothing
End Sub

Private Sub cmdPostToWeb_Click()
    '***************************************
    ' Module/Form Name   : frm_job_edit
    '
    ' Procedure Name     : cmdPostToWeb_Click
    '
    ' Purpose            :
    '
    ' Date Created       : 18/04/2002
    '
    ' Author             : GARETH SAUNDERS
    '
    ' Amendment History  : Date       Author    Description
    '                    : --------------------------------
    '                    : 18/04/2002 GARETH SAUNDERS
    '
    '***************************************
    '
    On Error GoTo cmdPostToWeb_Click_Error
    '
    '******** Code Starts Here *************
    '
    Dim strEmail As String
    Dim oOutlook As OutlookClass
    Dim oSearch As Search2
    Dim strMessage As String
    Dim oSearchResult As SearchResult
    '
    '   Simplifed Web Searching - ensure that there are no more than six searches as
    '   the website cannot show more than six tabs.
    '
    If moJob.Searches.Count > 6 Then
        MsgBox "A maximum of six Searches may be posted for each Web Job." & vbCrLf & _
              "If you wish to post more, a new job must be created for the customer.", vbExclamation, App.Title
        Exit Sub
    End If
    '
    '   Ensure there are no duplicate entries in any of the searches.
    '
    For Each oSearch In moJob.Searches
        oSearch.DuplicateSearchResults.DuplicateEntries
        If oSearch.DuplicateSearchResults.Count > 0 Then
            strMessage = "The following photographs are duplicated in search '" & oSearch.Description & "':" & vbCrLf & vbCrLf
            For Each oSearchResult In oSearch.DuplicateSearchResults
                strMessage = strMessage & CStr(oSearchResult.BatchNo) & " - " & oSearchResult.PhotoDesc & vbCrLf & vbCrLf
            Next oSearchResult
            strMessage = strMessage & "Duplicates within a Search must be removed before posting to the web."
            MsgBox strMessage, vbExclamation, App.Title
            Exit Sub
        End If
    Next oSearch
    '
    '   Check at least one Image has been selected and none have been confirmed.
    '
    If moJob.NoPhotos = 0 Then
        MsgBox "At least one Image must exist on the Job", vbExclamation, App.Title
        Exit Sub
    End If
    If moJob.NoConfirmed > 0 Then
        MsgBox "There must be no confirmed Images on the Job if you wish to Post it to the Web", vbExclamation, App.Title
        Exit Sub
    End If
    '
    '   Confirm that the user wants to Post to the Web.
    '
    If MsgBox("Are you sure you wish to Post the Job to the Web?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
        Exit Sub
    End If
    '
    '   Check that the Web Address to be posted to has been set.
    '
    If Trim(goSystemConfig.PostWebAddress) = "" Then
        MsgBox "The Web Address to be posted to must be set." & vbCrLf & _
               "Please go to Tools/Options and select the FTP Settings Tab.", vbExclamation
        Exit Sub
    End If
    '
    DoEvents
    Screen.MousePointer = vbHourglass
    '
    '   Post Job to the Web.
    '
''    gdbADO.BeginTrans
    begin_trans
        On Error Resume Next
        moJob.PostToWeb
        If Err.Number = 0 Then
            On Error GoTo cmdPostToWeb_Click_Error
        ElseIf Err.Number = vbObjectError + 11 Then
            MsgBox Err.Description, vbExclamation, App.Title
            Screen.MousePointer = vbDefault
''            gdbADO.RollbackTrans
            roll_back
            Exit Sub
        ElseIf Err.Number = 12007 Then
            MsgBox "Cannot connect to FTP Server. Please try again later.", vbExclamation, App.Title
            Screen.MousePointer = vbDefault
            'gdbADO.RollbackTrans
            roll_back
            Exit Sub
        Else
            ErrorSave
            On Error GoTo cmdPostToWeb_Click_Error
            ErrorRestore
        End If
        '
        '   Send an Email to the Contact.
        '
        If goSystemConfig.WebSearchTestEmail = "" Then
            strEmail = moContact.Email
        Else
            strEmail = goSystemConfig.WebSearchTestEmail
        End If
        Set oOutlook = New OutlookClass
        oOutlook.HTMLEmail strEmail, _
                           moCustomer.CustomerName & "  Job: " & moJob.reference & " (" & moJob.JobNo & ")", _
                           moJob.PostedHTMLEmail, _
                           goSystemConfig.PostedImages
        If Not oOutlook.MailSent Then
            MsgBox "Email not sent. Job has not been posted.", vbExclamation, App.Title
''            gdbADO.RollbackTrans
            roll_back
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        Set oOutlook = Nothing
''    gdbADO.CommitTrans
    commit_trans
    '
    cmdPostToWeb.Enabled = False
    If Not goSystemConfig.BasicImageWhere Then
        cmdReceiveFeedback.Enabled = True
    End If
    '
    update_loaded_screens
    '
    Screen.MousePointer = vbDefault
    MsgBox "Job has been Posted to the Web", vbInformation, App.Title
    RefreshJob
    '
    '********* Code Ends Here **************
    '
    Exit Sub
    '
cmdPostToWeb_Click_Error:
    DisplayError , "frm_job_edit.cmdPostToWeb_Click", vbExclamation
    MsgBox "Job has not been posted.", vbExclamation, App.Title
End Sub

Private Sub cmdReceiveFeedback_Click()
    '***************************************
    ' Module/Form Name   : frm_job_edit
    '
    ' Procedure Name     : cmdReceiveFeedback_Click
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
    On Error GoTo cmdReceiveFeedback_Click_Error
    '
    '******** Code Starts Here *************
    '
    Dim fReceiveFeedback As frmJobReceiveFeedBack

    If MsgBox("Are you sure you wish to Record Customers Web Search Feedback?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
        Exit Sub
    End If

    Set fReceiveFeedback = New frmJobReceiveFeedBack
    fReceiveFeedback.Display moJob

    RefreshJob
    '
    update_loaded_screens
    '
    '********* Code Ends Here **************
    '
Exit Sub
    '
cmdReceiveFeedback_Click_Error:
    DisplayError , "frm_job_edit.cmdReceiveFeedback_Click", vbExclamation
End Sub

Private Sub cmdView_Click()
'***************************************
' Module/Form Name   : frm_job_edit
'
' Procedure Name     : cmdView_Click
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
On Error GoTo cmdView_Click_Error
'
'******** Code Starts Here *************
'
    ViewSearch
'
'********* Code Ends Here **************
'
    Exit Sub
    '
cmdView_Click_Error:
    DisplayError , "frm_job_edit.cmdView_Click", vbExclamation
End Sub

Private Sub ViewSearch()
'***************************************
' Module/Form Name   : frm_job_edit
'
' Procedure Name     : ViewSearch
'
' Purpose            :
'
' Date Created       : 29/06/2004
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo ViewSearch_Error
'
'******** Code Starts Here *************
'
    Dim fSearch As frm_search
    
    If smgSearches.CurrentRow = 0 Then
        MsgBox "Select a Search to View", vbExclamation, App.Title
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    Set fSearch = New frm_search
    fSearch.display_search_form view, _
                                moJob, _
                                moJob.Searches.Item(smgSearches.Column(1).Value).SearchNo
    Set fSearch = Nothing
    Screen.MousePointer = vbDefault
'
'********* Code Ends Here **************
'
   Exit Sub
'
ViewSearch_Error:
    ErrorRaise "frm_job_edit.ViewSearch"
End Sub

Private Sub com_job_details_Click()
    '***************************************
    ' Module/Form Name   : frm_job_edit
    '
    ' Procedure Name     : com_job_details_Click
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
On Error GoTo com_job_details_Click_Error
    '
    '******** Code Starts Here *************
    '
    If is_form_loaded("frm_job_details") Then
        On Error Resume Next
        frm_job_details.SetFocus
    Else
        frm_job_details.input_details "E", moJob
    End If
    '
    '********* Code Ends Here **************
    '
Exit Sub
    '
com_job_details_Click_Error:
DisplayError , "frm_job_edit.com_job_details_Click", vbExclamation
End Sub

Private Sub cmdNew_Click()
'***************************************
' Module/Form Name   : frm_job_edit
'
' Procedure Name     : cmdNew_Click
'
' Purpose            :
'
' Date Created       : 16/10/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 16/10/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo cmdNew_Click_Error
'
'******** Code Starts Here *************
'
    If goSystemConfig.SageLink Then
        If Not moCustomer.TermsAndConditionsApproved Then
            MsgBox "Terms & Conditions have not yet been agreed for this customer." & vbCrLf & _
                   "You cannot therefore create a new Search.", vbExclamation
            Exit Sub
        End If
        If moCustomer.OnHold Then
            MsgBox "This customer is on hold." & vbCrLf & _
                   "You cannot therefore create a new Search.", vbExclamation
            Exit Sub
        End If
    End If
    
    Set fSearchDetails = New frm_search_details
    fSearchDetails.load_search_title moJob, Add, moCurrentSearch
    Set fSearchDetails = Nothing
    '
    '********* Code Ends Here **************
    '
    Exit Sub
        '
cmdNew_Click_Error:
    DisplayError , "frm_job_edit.cmdNew_Click", vbExclamation
End Sub

Private Sub Form_Initialize()
    mblnAllowResize = True
    mblnRefreshPrimed = False
End Sub

Private Sub Form_Load()
'***************************************
' Module/Form Name   : frm_job_edit
'
' Procedure Name     : Form_Load
'
' Purpose            :
'
' Date Created       : 28/04/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 28/04/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo Form_Load_Error
'
'******** Code Starts Here *************
'
    If goSystemConfig.BasicImageWhere Then
        cmdPostToWeb.Visible = False
        cmdReceiveFeedback.Visible = False
    End If

''    Resize.Hook Me.hWnd, 500, 400
    Set smgSearches.Container = fraSearches
    With smgSearches
        .Height = 2535
        .Left = 120
        .Top = 960
        .Width = 7095
    End With
    
    With smgSearches
      .Columns = 4
      .KeyCol = 1
      .Column(2).Header = "Description"
      .Column(2).Align = flexAlignLeftTop
      .Column(3).Header = "No. of Photographs"
      .Column(3).Align = flexAlignRightTop
      .Column(4).Header = "Photo Type"
      .Column(4).Align = flexAlignLeftTop
    End With
    
    timRefresh.Interval = 500
    
    com_position_form Me

'
'********* Code Ends Here **************
'
    Exit Sub
    '
Form_Load_Error:
    DisplayError , "frm_job_edit.Form_Load", vbExclamation
End Sub

Private Sub Form_Resize()
'***************************************
' Module/Form Name   : frm_job_edit
'
' Procedure Name     : Form_Resize
'
' Purpose            :
'
' Date Created       : 28/04/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 16/10/2002 GARETH SAUNDERS
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
''    '
''    '   Resizing Rules for MID Children.
''    '
''    If Not ChildWindowsMaxed Then
''        For Each f In gcolMaxedWindows
''            If f.hWnd <> Me.hWnd Then
''                f.ForceResize
''            End If
''        Next f
''    End If
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
    '   Resize only if required.
    '
    If Not mblnAllowResize Then Exit Sub
    '
'    MakeAllControlsInvisible Me
    '
    MIN_WIDTH = 7500
    MIN_HEIGHT = 5500
    header_height = 440

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
    On Error Resume Next
    '
    '   Size the frames.
    '
    fraJob.Move 50, _
                50, _
                Me.Width - 200
    '
    fraSearches.Move 50, _
                   fraJob.Top + fraJob.Height + 200, _
                   fraJob.Width, _
                   Me.Height - goSystemConfig.TitleBarHeight - (fraJob.Top + fraJob.Height + 200) - 100
    '
    '   Position the Buttons.
    '
    com_job_details.Move fraJob.Width - com_job_details.Width - 100, lblJobNo.Top
    
    If goSystemConfig.BasicImageWhere Then
        cmdJobNotes.Move com_job_details.Left, com_job_details.Top + com_job_details.Height + 80
    Else
        cmdPostToWeb.Move com_job_details.Left, com_job_details.Top + com_job_details.Height + 80
        cmdReceiveFeedback.Move cmdPostToWeb.Left, cmdPostToWeb.Top + cmdPostToWeb.Height + 80
        cmdJobNotes.Move cmdPostToWeb.Left, cmdReceiveFeedback.Top + cmdReceiveFeedback.Height + 80
    End If
    
    '
    cmdNew.Move 100, 300
    cmdEdit.Move cmdNew.Left + cmdNew.Width + 100, cmdNew.Top
    cmdDelete.Move cmdEdit.Left + cmdEdit.Width + 100, cmdNew.Top
    cmdView.Move cmdDelete.Left + cmdDelete.Width + 100, cmdNew.Top
    cmdExit.Left = fraSearches.Width - cmdExit.Width - 100
    '
    '   Size the Grid.
    '
    With smgSearches
        .Redraw = False
        .Move 50, _
              cmdNew.Top + cmdNew.Height + 100, _
              fraSearches.Width - 100, _
              fraSearches.Height - (cmdNew.Top + cmdNew.Height + 100) - 100
        .Column(1).Width = 0
        .Column(3).Width = 1000
        .Column(4).Width = 1000
        .Column(2).Width = (.Width - scroll_bar_width - _
                                     .Column(1).Width - _
                                     .Column(3).Width) - _
                                     .Column(4).Width
        .Redraw = True
    End With
    smgSearches.ResizeRows
'    MakeAllControlsVisible Me
    '
    '********* Code Ends Here **************
    '
    Exit Sub
    '
Form_Resize_Error:
    DisplayError , "frm_job_edit.Form_Resize"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set moCustomer = Nothing
    Set moJob = Nothing
''    Resize.Unhook Me.hWnd
    On Error Resume Next
    gcolMaxedWindows.Remove CStr(Me.hWnd)
End Sub

Private Sub cmdExit_Click()
'***************************************
' Module/Form Name   : frm_job_edit
'
' Procedure Name     : cmdExit_Click
'
' Purpose            :
'
' Date Created       : 23/04/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 16/10/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo cmdExit_Click_Error
'
'******** Code Starts Here *************
'
    If Not is_form_loaded("frm_job_maint") Then
        Load frm_job_maint
        frm_job_maint.Display mod_job_no
    Else
        frm_job_maint.redisplay
    End If
    '
    Unload Me
    Set frm_job_edit = Nothing
    DoEvents
    '
    '********* Code Ends Here **************
    '
    Exit Sub
    '
cmdExit_Click_Error:
    DisplayError , "frm_job_edit.cmdExit_Click", vbExclamation
End Sub

Private Sub fSearchDetails_TitleChanged()
    '
    '   Need to refresh the job to let it know that the search details have changed.
    '
    moJob.Searches.Refresh
End Sub

Private Sub fSearchDetails_TitleCreated(SearchNo As Long)
    mod_current_search = SearchNo
End Sub

Private Sub smgSearches_Click()
    mod_current_search = IIf(smgSearches.Column(1).Value = "", 0, smgSearches.Column(1).Value)
End Sub

Private Sub EnableButtons()
'***************************************
' Module/Form Name   : frm_job_edit
'
' Procedure Name     : EnableButtons
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
On Error GoTo EnableButtons_Error
'
'******** Code Starts Here *************
'
    Dim blnFound    As Boolean
    
    On Error Resume Next
    Set moContact = moJob.Contact
    Select Case Err.Number
    Case Is = 0, vbObjectError + 5, vbObjectError + 9
        On Error GoTo EnableButtons_Error
    Case Else
        ErrorSave
        On Error GoTo EnableButtons_Error
        ErrorRestore
    End Select
    '
''
''    If moJob.RequestedByID <> 0 Then
''        Set moContact = moCustomer.Contacts.Item(CStr(moJob.RequestedByID))
''    Else
''        blnFound = False
''        For Each oContact In moCustomer.Contacts
''            If oContact.Name = moJob.RequestedBy Then
''                blnFound = True
''                Exit For
''            End If
''        Next oContact
''        '
''        If blnFound Then
''            Set moContact = oContact
''        End If
''    End If
''    '
''    If Err.Number <> 0 And Err.Number <> 5 And Err.Number <> 9 Then
''      ErrorSave
''      On Error GoTo EnableButtons_Error
''      ErrorRestore
''    Else
''      On Error GoTo EnableButtons_Error
''    End If
''        '
    If moJob.DeliveryNoteNo > 0 Then
        cmdNew.Enabled = False
        cmdEdit.Enabled = False
        cmdDelete.Enabled = False
        com_job_details.Enabled = False
        cmdPostToWeb.Enabled = False
        cmdReceiveFeedback.Enabled = False
    Else
        '
        If Not moContact Is Nothing Then
          If moContact.WebUser = True Then
              If moJob.WebStatus = "P" Then
                  cmdPostToWeb.Enabled = False
                  If Not goSystemConfig.BasicImageWhere Then
                      cmdReceiveFeedback.Enabled = True
                  End If
              ElseIf moJob.WebStatus = "R" Then
                  cmdPostToWeb.Enabled = False
                  cmdReceiveFeedback.Enabled = False
              Else
                  If Not goSystemConfig.BasicImageWhere Then
                      cmdPostToWeb.Enabled = True
                  End If
                  cmdReceiveFeedback.Enabled = False
              End If
          Else
              cmdPostToWeb.Enabled = False
              cmdReceiveFeedback.Enabled = False
          End If
        Else
          cmdPostToWeb.Enabled = False
          cmdReceiveFeedback.Enabled = False
        End If
        '
        Select Case moJob.WebStatus
          Case Is = "P"
              cmdNew.Enabled = False
              cmdEdit.Enabled = False
              cmdDelete.Enabled = False
          Case Is = "R"
              cmdNew.Enabled = True
              cmdEdit.Enabled = True
              cmdDelete.Enabled = True
          Case Else
              cmdNew.Enabled = True
              cmdEdit.Enabled = True
              cmdDelete.Enabled = True
        End Select
    End If
'
'********* Code Ends Here **************
'
Exit Sub
    '
EnableButtons_Error:
    ErrorRaise "frm_job_edit.EnableButtons"
End Sub

Private Sub smgSearches_DblClick()
    '***************************************
    ' Module/Form Name   : frm_job_edit
    '
    ' Procedure Name     : smgSearches_DblClick
    '
    ' Purpose            :
    '
    ' Date Created       : 09/06/2002
    '
    ' Author             : GARETH SAUNDERS
    '
    ' Amendment History  : Date       Author    Description
    '                    : --------------------------------
    '                    : 09/06/2002 GARETH SAUNDERS
    '
    '***************************************
    '
On Error GoTo smgSearches_DblClick_Error
    '
    '******** Code Starts Here *************
    '
If cmdEdit.Enabled Then
  EditSearch
Else
  ViewSearch
End If
    '
    '********* Code Ends Here **************
    '
Exit Sub
    '
smgSearches_DblClick_Error:
    DisplayError , "frm_job_edit.smgSearches_DblClick", vbExclamation
End Sub

Public Sub ForceResize()
    gblnResizeMaxedWindows = False
    Form_Resize
    gblnResizeMaxedWindows = True
End Sub

Private Sub update_loaded_screens()
'***************************************
' Module/Form Name   : frm_job_edit
'
' Procedure Name     : update_loaded_screens
'
' Purpose            :
'
' Date Created       : 30/06/2004
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo update_loaded_screens_Error
'
'******** Code Starts Here *************
'
    LockWindow Me.hWnd
    '
    If is_form_loaded("frm_job_maint") Then
        frm_job_maint.PendingRedisplay
    End If
    
    Me.SetFocus
    UnlockWindow
'
'********* Code Ends Here **************
'
   Exit Sub
'
update_loaded_screens_Error:
    ErrorRaise "frm_job_edit.update_loaded_screens"
End Sub

Private Sub timRefresh_Timer()
'***************************************
' Module/Form Name   : frm_job_edit
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
    If moJob Is Nothing Then
        Exit Sub
    End If
    
    If moJob.Searches.PendingRefresh Then
        If mblnRefreshPrimed Then
            mblnRefreshPrimed = False
            moJob.Searches.PendingRefresh = False
            moJob.Searches.Refresh
            moJob.Read moJob.JobNo
            RefreshJob False
        Else
            DBEngine.Idle dbRefreshCache
            mblnRefreshPrimed = True
        End If
    End If
'
'********* Code Ends Here **************
'
   Exit Sub
'
timRefresh_Timer_Error:
    DisplayError , "frm_job_edit.timRefresh_Timer", vbExclamation
End Sub

Public Sub PendingRedisplay()
    moJob.Searches.PendingRefresh = True
    mblnRefreshPrimed = False
End Sub


