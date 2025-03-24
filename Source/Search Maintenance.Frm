VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_search 
   AutoRedraw      =   -1  'True
   ClientHeight    =   7260
   ClientLeft      =   3270
   ClientTop       =   1950
   ClientWidth     =   9345
   ForeColor       =   &H00000000&
   Icon            =   "Search Maintenance.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7260
   ScaleWidth      =   9345
   WindowState     =   2  'Maximized
   Begin VB.Frame fraBatches 
      Caption         =   "Batches"
      ForeColor       =   &H80000001&
      Height          =   4095
      Left            =   0
      TabIndex        =   13
      Top             =   2880
      Width           =   8895
      Begin VB.CommandButton cmdConfirmDigitalImages 
         Caption         =   "C&onfirm"
         Height          =   675
         Left            =   6720
         Picture         =   "Search Maintenance.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   360
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdRefine 
         Caption         =   "&Refine"
         Height          =   675
         Left            =   5880
         Picture         =   "Search Maintenance.frx":0B04
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin MSComctlLib.ImageList iml_ticks 
         Left            =   6840
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   8
         ImageHeight     =   8
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Search Maintenance.frx":0C4E
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lsv_batches 
         Height          =   2655
         Left            =   240
         TabIndex        =   23
         Top             =   1200
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   4683
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Frame fraAdd 
         Height          =   975
         Left            =   2160
         TabIndex        =   16
         Top             =   140
         Width           =   2895
         Begin VB.CommandButton cmdConfirm 
            Caption         =   "C&onfirm"
            Height          =   675
            Left            =   120
            Picture         =   "Search Maintenance.frx":0E60
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   240
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            Height          =   675
            Left            =   960
            Picture         =   "Search Maintenance.frx":1522
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   200
            UseMaskColor    =   -1  'True
            Width           =   735
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Remove"
            Height          =   675
            Left            =   1860
            Picture         =   "Search Maintenance.frx":1BE4
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   300
            UseMaskColor    =   -1  'True
            Width           =   735
         End
         Begin VB.TextBox txt_batch_no 
            Enabled         =   0   'False
            Height          =   315
            Left            =   100
            MaxLength       =   8
            TabIndex        =   17
            Top             =   240
            Width           =   795
         End
      End
      Begin VB.OptionButton opt_search_maint 
         Caption         =   "S&earch Maintenance"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.OptionButton opt_bar_code 
         Caption         =   "&Bar Code Confirmation"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   1935
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   675
         Left            =   7560
         Picture         =   "Search Maintenance.frx":22A6
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "F&ind"
         Height          =   675
         Left            =   5040
         Picture         =   "Search Maintenance.frx":2968
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   735
      End
   End
   Begin VB.Frame fraSearch 
      Caption         =   "Search"
      ForeColor       =   &H80000001&
      Height          =   1095
      Left            =   0
      TabIndex        =   9
      Top             =   1680
      Width           =   7335
      Begin VB.CommandButton com_change 
         Caption         =   "&CHANGE ..."
         Height          =   435
         Left            =   6180
         TabIndex        =   12
         Top             =   360
         Width           =   1035
      End
      Begin VB.TextBox txt_search_description 
         BackColor       =   &H8000000F&
         ForeColor       =   &H80000001&
         Height          =   495
         Left            =   1440
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   360
         Width           =   4575
      End
      Begin VB.Label lab_description_lab 
         Caption         =   "Description:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1035
      End
   End
   Begin VB.Frame fraJob 
      Caption         =   "Job"
      ForeColor       =   &H80000001&
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      Begin VB.TextBox txt_customer_name 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000001&
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   600
         Width           =   5775
      End
      Begin VB.TextBox txt_address_line_1 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000001&
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   900
         Width           =   5775
      End
      Begin VB.TextBox txt_job_reference 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000001&
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1200
         Width           =   5775
      End
      Begin VB.TextBox txtJobNo 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000001&
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   300
         Width           =   3915
      End
      Begin VB.Label lab_job_reference_lab 
         Caption         =   "Job Reference:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   1155
      End
      Begin VB.Label lab_customer_name_lab 
         Caption         =   "Customer Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lab_address_line_1_lab 
         Caption         =   "Address Line 1:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   900
         Width           =   1215
      End
      Begin VB.Label lblJobNo 
         Caption         =   "Job No:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   1215
      End
   End
   Begin MSComctlLib.StatusBar stb_search 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   24
      Top             =   7005
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5292
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5292
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5292
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SQL As String

Private snap_customer As DAO.Recordset
Private snap_search_result As DAO.Recordset
Private snap_batch As DAO.Recordset
Private snap_batch_search As DAO.Recordset


Private dyna_photograph As DAO.Recordset
Private dyna_search As DAO.Recordset
Private dyna_search_result As DAO.Recordset
Private dyna_job As DAO.Recordset

Dim table_search_result As DAO.Recordset

Public mod_action                       As UpdateMode
Public mstrSearchString                 As String
Public mod_customer_no                  As Integer
Public mod_job_no                       As Long
Private mod_search_no                   As Long
Private mod_search_result_key           As Long
Private mbolDontResize                  As Boolean
Private WithEvents fSearchDetails       As frm_search_details
Attribute fSearchDetails.VB_VarHelpID = -1
Private WithEvents fSearchPhotograph    As frm_search_photograph
Attribute fSearchPhotograph.VB_VarHelpID = -1
'
'   Objects
'
Private moCustomer As Customer2
Private moJob As Job2
Private moSearch As Search2

Public Sub edit_photo()
'***************************************
' Module/Form Name   : frm_search
'
' Procedure Name     : edit_photo
'
' Purpose            :
'
' Date Created       : 17/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo edit_photo_Error
'
'******** Code Starts Here *************
'
    '
    '   Used for invoking the edit photograph from from search maintenance.
    '   Find the search result in the snapshot using the mod_search_result_key set up
    '   in the mouseup event.
    '
        If mod_search_result_key = 0 Then
            MsgBox "Please select an image", vbExclamation, "Search Maintenance"
            Exit Sub
        End If
    '
        Set snap_search_result = search_result.Read(mod_search_result_key, dbOpenSnapshot)
    '
    '   Invoke the edit photograph from.
    '
        frm_photograph_edit.load_batch LTrim(snap_search_result!batch_no)
        frm_photograph_edit.SetFocus
'
'********* Code Ends Here **************
'
   Exit Sub
'
edit_photo_Error:
    ErrorRaise "frm_search.edit_photo"
End Sub

Private Sub display_totals()
'***************************************
' Module/Form Name   : frm_search
'
' Procedure Name     : display_totals
'
' Purpose            :
'
' Date Created       : 18/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo display_totals_Error
'
'******** Code Starts Here *************
'
    Dim pnlX As Panel
    Dim number_confirmed As Integer
    Dim i As Integer

    number_confirmed = 0
    For i = 1 To lsv_batches.ListItems.Count
        If lsv_batches.ListItems(i).SmallIcon = 1 Then
            number_confirmed = number_confirmed + 1
        End If
    Next i

    Set pnlX = stb_search.Panels.Item(1)
    pnlX.Text = "Unconfirmed: " + Str(lsv_batches.ListItems.Count - number_confirmed)

    Set pnlX = stb_search.Panels.Item(2)
    pnlX.Text = "Confirmed: " + Str(number_confirmed)

    Set pnlX = stb_search.Panels.Item(3)
    pnlX.Text = "Total: " + Str(lsv_batches.ListItems.Count)
'
'********* Code Ends Here **************
'
   Exit Sub
'
display_totals_Error:
    ErrorRaise "frm_search.display_totals"
End Sub

Public Sub redisplay_search_details()
'***************************************
' Module/Form Name   : frm_search
'
' Procedure Name     : redisplay_search_details
'
' Purpose            :
'
' Date Created       : 18/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo redisplay_search_details_Error
'
'******** Code Starts Here *************
'
    moSearch.SearchResults.Refresh
    display_search_details mod_search_no
'
'********* Code Ends Here **************
'
   Exit Sub
'
redisplay_search_details_Error:
    ErrorRaise "frm_search.redisplay_search_details"
End Sub

Private Sub update_loaded_screens()
'***************************************
' Module/Form Name   : frm_search
'
' Procedure Name     : update_loaded_screens
'
' Purpose            :
'
' Date Created       : 18/12/2002
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
    Dim fSearch As frm_search
    Dim colFrmSearch As New Collection

    LockWindow Me.hWnd
    If is_form_loaded("frm_search", Me, colFrmSearch) Then
        For Each fSearch In colFrmSearch
            fSearch.redisplay_search_details
        Next fSearch
    End If
    '
    If is_form_loaded("frm_photograph_view") Then
        frm_photograph_view.RedisplayTransparencies
    End If
    '
    If is_form_loaded("frm_photograph_edit") Then
        frm_photograph_edit.redisplay_images
    End If
    '
    If is_form_loaded("frm_job_edit") Then
        moJob.Searches.PendingRefresh = True
''        frm_job_edit.RefreshJob False
    End If
    '
    If is_form_loaded("frm_job_maint") Then
        frm_job_maint.PendingRedisplay
    End If
    '
    Set colFrmSearch = Nothing
    On Error Resume Next
    Me.SetFocus
    UnlockWindow
'
'********* Code Ends Here **************
'
   Exit Sub
'
update_loaded_screens_Error:
    ErrorRaise "frm_search.update_loaded_screens"
End Sub

Private Function valid_input() As Boolean
'***************************************
' Module/Form Name   : frm_search
'
' Procedure Name     : valid_input
'
' Purpose            :
'
' Date Created       : 18/12/2002
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
    Dim oSearchResult       As SearchResult
    
    valid_input = False

    If Len(RTrim(txt_batch_no.Text)) = 0 Then
        MsgBox "Please enter a photograph number", vbExclamation
        txt_batch_no.SetFocus
        Exit Function
    End If

    If (InStr(txt_batch_no, ".") <> 0 And (last_character(txt_batch_no) = "." Or InStr(txt_batch_no, ".") = 1)) Or _
       Not IsNumeric(txt_batch_no) Or _
       (InStr(txt_batch_no, ".") = 0 And opt_bar_code.Value = True) Then
            MsgBox "Please enter a valid image number, for example: 12345.2", vbExclamation
            txt_batch_no.SetFocus
            Exit Function
    End If
    '
    '   If this is a digital Search then only accept a Photograph number without an image number.
    '
    If moSearch.PhotoType = "D" Then
        If InStr(txt_batch_no.Text, ".") <> 0 Then
            MsgBox "Please enter a Photograph number without an Image Number for Digital Searches.", vbExclamation
            txt_batch_no.SetFocus
            Exit Function
        End If
    End If
    '
    '   Is the photograph already on a Digital Search?
    '
    If moSearch.PhotoType = "D" Then
        For Each oSearchResult In moSearch.SearchResults
            If oSearchResult.BatchNo = CLng(txt_batch_no.Text) Then
                MsgBox "Digital Photograph '" & CLng(txt_batch_no.Text) & "' has already been selected for this Search.", vbExclamation
                txt_batch_no.SetFocus
                Exit Function
            End If
        Next oSearchResult
    End If
    '
    valid_input = True
'
'********* Code Ends Here **************
'
   Exit Function
'
valid_input_Error:
    ErrorRaise "frm_search.valid_input"
End Function

Public Sub display_title()
'***************************************
' Module/Form Name   : frm_search
'
' Procedure Name     : display_title
'
' Purpose            :
'
' Date Created       : 18/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo display_title_Error
'
'******** Code Starts Here *************
'
'
'   Set up Form Title and Job Reference.
'

    Me.Caption = IIf(mod_action = Edit Or mod_action = Add, "Maintain", "View") & " Search '" _
                & moSearch.DisplayDescription _
                & "' for Job '" _
                & moJob.reference & "'"

    txt_job_reference = moJob.reference

'
'********* Code Ends Here **************
'
   Exit Sub
'
display_title_Error:
    ErrorRaise "frm_search.display_title"
End Sub

Private Sub cmdConfirmDigitalImages_Click()
'***************************************
' Module/Form Name   : frm_search
'
' Procedure Name     : cmdConfirmDigitalImages_Click
'
' Purpose            :
'
' Date Created       : 25/06/2004
'
' Author             : ADMINISTRATOR
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo cmdConfirmDigitalImages_Click_Error
'
'******** Code Starts Here *************
'
    Dim oActivity       As Activity
    Dim oFSO            As Scripting.FileSystemObject
    Dim oSearchResult   As SearchResult
    Dim strImageFolder  As String
    Dim oDigitalImage   As DigitalImage
    
    If moSearch.SearchResults.Count = 0 Then
        MsgBox "There are no Pictures to confirm", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Are you sure you wish to confirm all the digital images on this Search?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
        Exit Sub
    End If
    '
    begin_trans
        moSearch.SearchResults.ConfirmAll
        '
        For Each oActivity In moSearch.Activities
            oActivity.Status = "3"
            oActivity.EndDate = Now()
            oActivity.update
        Next oActivity
        moSearch.Activities.Refresh
        redisplay_search_details
''        '
''        '   Create a 'Burn to CD' activity.
''        '
''        CreateDigitialActivity moSearch, "BURN"
    commit_trans
    '
    update_loaded_screens
    '
    EnableButtons
'
'********* Code Ends Here **************
'
   Exit Sub
'
cmdConfirmDigitalImages_Click_Error:
    DisplayError , "frm_search.cmdConfirmDigitalImages_Click", vbExclamation
End Sub

Private Sub EnableButtons()
'***************************************
' Module/Form Name   : frm_search
'
' Procedure Name     : EnableButtons
'
' Purpose            :
'
' Date Created       : 25/06/2004
'
' Author             : ADMINISTRATOR
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo EnableButtons_Error
'
'******** Code Starts Here *************
'
    If moSearch.PhotoType = "D" Then
        If mod_action = view Then
            cmdConfirmDigitalImages.Enabled = False
        Else
            With moSearch.Activities
                .ActivityType = "REFS"
                .Refresh
                If .Count = 0 Then
                    cmdConfirmDigitalImages.Enabled = False
                Else
                    cmdConfirmDigitalImages.Enabled = True
                End If
            End With
        End If
    End If
'
'********* Code Ends Here **************
'
   Exit Sub
'
EnableButtons_Error:
    ErrorRaise "frm_search.EnableButtons"
End Sub

Private Sub cmdRefine_Click()
'***************************************
' Module/Form Name   : frm_search
'
' Procedure Name     : cmdRefine_Click
'
' Purpose            :
'
' Date Created       : 28/06/2004
'
' Author             : ADMINISTRATOR
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo cmdRefine_Click_Error
'
'******** Code Starts Here *************
'
    If Not CheckUnconfirmPictures Then
        Exit Sub
    End If
    '
    If moSearch.SearchResults.Count = 0 Then
        MsgBox "There are no Pictures to Refine", vbExclamation
        Exit Sub
    End If
    '
    EnableButtons
    '
    Screen.MousePointer = vbHourglass
    Set fSearchPhotograph = Nothing
    Set fSearchPhotograph = New frm_search_photograph
    fSearchPhotograph.display_screen IIf(mod_action = view, srreview, srRefine), moSearch, moJob
    redisplay_search_details
    Set fSearchPhotograph = Nothing
    On Error Resume Next
    Me.SetFocus
    Screen.MousePointer = vbDefault
'
'********* Code Ends Here **************
'
   Exit Sub
'
cmdRefine_Click_Error:
    DisplayError , "frm_search.cmdRefine_Click", vbExclamation
End Sub

Private Function CheckUnconfirmPictures(Optional pblnRedisplay As Boolean = True, _
                                        Optional pblnAddRefine As Boolean = True) As Boolean
'***************************************
' Module/Form Name   : frm_search
'
' Procedure Name     : CheckUnconfirmPictures
'
' Purpose            :
'
' Date Created       : 28/06/2004
'
' Author             : ADMINISTRATOR
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo CheckUnconfirmPictures_Error
'
'******** Code Starts Here *************
'
    Dim oActivity       As Activity
    
    CheckUnconfirmPictures = False
    
    If moSearch.PhotoType = "D" Then
        '
        '   If there are no outstanding Refine Search activities and there are pictures on the
        '   Search then another Refine Activity will need to be created and all pictures will be set to unconfirmed.
        '
        moSearch.Activities.ActivityType = "REFS"
        moSearch.Activities.Refresh
        If moSearch.Activities.Count = 0 And moSearch.SearchResults.Count <> 0 And mod_action <> view Then
            If MsgBox("This will unconfirm all pictures." & vbCrLf & _
                      "Are you sure you wish to continue?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                Exit Function
            End If
            '
            '   Unconfirm all pictures on the Search.
            '
            begin_trans
                moSearch.SearchResults.UnConfirmAll
                If pblnAddRefine Then
                    CreateDigitialActivity moSearch, "REFS"
                End If
                '
                '   Remove any 'Burn to CD' activities.
                '
                moSearch.Activities.ActivityType = "BURN"
                moSearch.Activities.Refresh
                For Each oActivity In moSearch.Activities
                    oActivity.EndDate = Now()
                    oActivity.Status = "3"
                    oActivity.update
                Next oActivity
            commit_trans
            If pblnRedisplay Then
                redisplay_search_details
                update_loaded_screens
            End If
        ElseIf moSearch.SearchResults.Count = 0 Then
            '
            '   There are no Search results left so close all activities.
            '
            begin_trans
                moSearch.Activities.ActivityType = ""
                moSearch.Activities.Refresh
                For Each oActivity In moSearch.Activities
                    oActivity.EndDate = Now()
                    oActivity.Status = "3"
                    oActivity.update
                Next oActivity
            commit_trans
        End If
    End If
    
    CheckUnconfirmPictures = True
'
'********* Code Ends Here **************
'
   Exit Function
'
CheckUnconfirmPictures_Error:
    ErrorRaise "frm_search.CheckUnconfirmPictures"
End Function

Private Sub CreateDigitialActivity(ByRef poSearch As Search2, _
                                   ByVal pstrActivityType As String)
'***************************************
' Module/Form Name   : frm_search
'
' Procedure Name     : CreateRefineActivity
'
' Purpose            :
'
' Date Created       : 28/06/2004
'
' Author             : ADMINISTRATOR
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo CreateRefineActivity_Error
'
'******** Code Starts Here *************
'
    Dim oActivity       As New Activity
    
    If poSearch.PhotoType = "D" Then
        poSearch.Activities.ActivityType = pstrActivityType
        poSearch.Activities.Refresh
        If poSearch.Activities.Count = 0 Then
            With oActivity
                .ActivityType = pstrActivityType
                .StartDate = Now()
                .SearchNo = poSearch.SearchNo
                .CustomerNo = moJob.CustomerNo
                .JobNo = moJob.JobNo
                .Status = "0"
                .Create
            End With
        End If
    End If
    poSearch.Activities.Refresh
'
'********* Code Ends Here **************
'
   Exit Sub
'
CreateRefineActivity_Error:
    ErrorRaise "frm_search.CreateRefineActivity"
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormMDIForm Then
        mdi_npls.mnuSearch.Visible = False
    End If
End Sub

Private Sub Form_Resize()
'***************************************
' Module/Form Name   : frm_search
'
' Procedure Name     : Form_Resize
'
' Purpose            :
'
' Date Created       : 18/12/2002
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
    Dim MIN_WIDTH As Integer
    Dim MIN_HEIGHT As Integer

    MIN_WIDTH = 8805
    MIN_HEIGHT = 6000

    If WindowState = 1 Then
        Exit Sub
    End If

    If Not ResizeForm(Me) Then
        Exit Sub
    End If
    
    LockWindow Me.hWnd
    If WindowState = 0 Then
        If Me.Width < MIN_WIDTH Then
            Me.Width = MIN_WIDTH
            Exit Sub
        End If
        If Height < MIN_HEIGHT Then
            Height = MIN_HEIGHT
            Exit Sub
        End If
    End If
'
'   position the frames.
'
    fraJob.Left = 50
    fraJob.Width = Me.Width - (2 * fraJob.Left) - 100
    fraSearch.Left = fraJob.Left
    fraSearch.Width = fraJob.Width
    fraBatches.Left = fraJob.Left
    fraBatches.Width = fraJob.Width
    fraBatches.Height = Me.Height - fraBatches.Top - 500 - stb_search.Height
'
'   Position the control buttons.
'
    fraAdd.Top = 120
    cmdConfirm.Top = cmdAdd.Top
    cmdConfirm.Left = cmdAdd.Left
    cmdDelete.Move cmdAdd.Left + cmdAdd.Width + 100, cmdAdd.Top
    cmdFind.Top = fraAdd.Top + cmdAdd.Top
    cmdFind.Left = fraAdd.Left + fraAdd.Width + 100
    cmdRefine.Move cmdFind.Left + cmdFind.Width + 100, cmdFind.Top
    cmdConfirmDigitalImages.Move cmdRefine.Left + cmdRefine.Width + 100, cmdRefine.Top
    cmdExit.Top = cmdFind.Top
    cmdExit.Left = fraBatches.Width - cmdExit.Width - 200
    com_change.Left = fraSearch.Width - com_change.Width - 200
'
'   Position the list box & panel headings.
'
    With lsv_batches
        .Visible = False
        .Move 100, _
              fraAdd.Top + fraAdd.Height + 100, _
              fraBatches.Width - 200, _
              fraBatches.Height - (fraAdd.Top + fraAdd.Height + 100) - 100
        .ColumnHeaders(1).Width = 300
        .ColumnHeaders(2).Width = 900
        .ColumnHeaders(3).Width = lsv_batches.Width - _
                                  .ColumnHeaders(1).Width - _
                                  .ColumnHeaders(2).Width - scroll_bar_width
        .Visible = True
    End With
'
'   position the text output fields.
'
    txt_customer_name.Width = fraJob.Width - txt_customer_name.Left - 100
    txt_address_line_1.Width = fraJob.Width - txt_address_line_1.Left - 100
    txt_job_reference.Width = fraJob.Width - txt_job_reference.Left - 100
    txt_search_description.Width = com_change.Left - txt_search_description.Left - 100

    UnlockWindow
'
'********* Code Ends Here **************
'
   Exit Sub
'
Form_Resize_Error:
    DisplayError , "frm_search.Form_Resize", vbExclamation
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting App.Title, "SearchMaintenance", "Width", Me.Width
    SaveSetting App.Title, "SearchMaintenance", "Height", Me.Height
    SaveSetting App.Title, "SearchMaintenance", "WindowState", Me.WindowState
    SaveSetting App.Title, "SearchMaintenance", "SortKey", lsv_batches.SortKey
    SaveSetting App.Title, "SearchMaintenance", "SortOrder", lsv_batches.SortOrder
    On Error Resume Next
    gcolMaxedWindows.Remove CStr(Me.hWnd)
End Sub

Private Sub fSearchDetails_TitleChanged()
'***************************************
' Module/Form Name   : frm_search
'
' Procedure Name     : fSearchDetails_TitleChanged
'
' Purpose            :
'
' Date Created       : 18/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo fSearchDetails_TitleChanged_Error
'
'******** Code Starts Here *************
'
    redisplay_search_details
'
'********* Code Ends Here **************
'
   Exit Sub
'
fSearchDetails_TitleChanged_Error:
    DisplayError , "frm_search.fSearchDetails_TitleChanged", vbExclamation
End Sub

Private Sub fSearchDetails_TitleCreated(SearchNo As Long)
'***************************************
' Module/Form Name   : frm_search
'
' Procedure Name     : fSearchDetails_TitleCreated
'
' Purpose            :
'
' Date Created       : 18/12/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : SearchNo - Long
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo fSearchDetails_TitleCreated_Error
'
'******** Code Starts Here *************
'
    mod_action = Edit
    display_search_details SearchNo
'
'********* Code Ends Here **************
'
   Exit Sub
'
fSearchDetails_TitleCreated_Error:
    DisplayError , "frm_search.fSearchDetails_TitleCreated", vbExclamation
End Sub

Private Sub fSearchPhotograph_PhotographSelected(poSearch As Search2, poDigitalImage As DigitalImage)
'***************************************
' Module/Form Name   : frm_search
'
' Procedure Name     : fSearchPhotograph_PhotographSelected
'
' Purpose            :
'
' Date Created       : 18/12/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : BatchNo - Long
'                    : SearchText - String
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo fSearchPhotograph_PhotographSelected_Error
'
'******** Code Starts Here *************
'
''    search_string = SearchText
''    add_photographs CStr(BatchNo), False, False
    mstrSearchString = poSearch.Description
    '
    '   Is the Photograph to be removed or added.
    '
    If poDigitalImage.Pending = pdgDelete Then
        Call RemovePhotograph(plngBatchNo:=poDigitalImage.BatchNo)
    ElseIf poDigitalImage.Pending <> pdgClean Then
        photograph.mblnMultipleAddition = True
        Call add_photographs(poDigitalImage.BatchNo, False, False, poSearch:=poSearch)
    End If
'
'********* Code Ends Here **************
'
   Exit Sub
'
fSearchPhotograph_PhotographSelected_Error:
    DisplayError , "frm_search.fSearchPhotograph_PhotographSelected", vbExclamation
End Sub

Private Sub fSearchPhotograph_PhotographSelectionComplete()
    update_loaded_screens
End Sub

Private Sub lsv_batches_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lsv_batches
        If ColumnHeader.Index = 2 Then
            SortImages IIf(.SortOrder = lvwAscending, lvwDescending, lvwAscending)
        Else
            .SortKey = ColumnHeader.Index - 1
            .SortOrder = IIf(.SortOrder = lvwAscending, lvwDescending, lvwAscending)
            .Sorted = True
        End If
    End With
End Sub

Private Sub SortImages(SortOrder As Integer)
    Dim intCount As Integer
    
    With lsv_batches
        '
        '   Add a temporary column.
        '
        .ColumnHeaders.Add , , , 1000
        For intCount = 1 To .ListItems.Count
            If moSearch.PhotoType = "T" Then
                .ListItems(intCount).SubItems(3) = Format(Mid(.ListItems(intCount).SubItems(1), 1, InStr(.ListItems(intCount).SubItems(1), ".") - 1), "00000000") & Format(Mid(.ListItems(intCount).SubItems(1), InStr(.ListItems(intCount).SubItems(1), ".") + 1), "000")
            Else
                .ListItems(intCount).SubItems(3) = Format(.ListItems(intCount).SubItems(1), "00000000")
            End If
        Next intCount
        .SortKey = 3
        .SortOrder = SortOrder
        .Sorted = True
        '
        '   Remove temporary column.
        '
        .ColumnHeaders.Remove 4
        .Sorted = False
        .SortKey = 1
    End With
End Sub

Private Sub lsv_batches_DblClick()
'***************************************
' Module/Form Name   : frm_search
'
' Procedure Name     : lsv_batches_DblClick
'
' Purpose            :
'
' Date Created       : 18/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo lsv_batches_DblClick_Error
'
'******** Code Starts Here *************
'
    edit_photo
'
'********* Code Ends Here **************
'
   Exit Sub
'
lsv_batches_DblClick_Error:
    DisplayError , "frm_search.lsv_batches_DblClick", vbExclamation
End Sub

Private Sub lsv_batches_ItemClick(ByVal Item As MSComctlLib.ListItem)
'***************************************
' Module/Form Name   : frm_search
'
' Procedure Name     : lsv_batches_ItemClick
'
' Purpose            :
'
' Date Created       : 18/12/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : Item - MSComctlLib.ListItem
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo lsv_batches_ItemClick_Error
'
'******** Code Starts Here *************
'
    Dim i As Integer
    i = 1
    Do
        If i > lsv_batches.ListItems.Count Then
            Exit Sub
        End If
        If lsv_batches.ListItems(i).Selected Then
            Exit Do
        End If
        i = i + 1
    Loop
'
'   Set the key ready for any manipulation necessary.
'
    mod_search_result_key = lsv_batches.ListItems(i).Tag

'
'********* Code Ends Here **************
'
   Exit Sub
'
lsv_batches_ItemClick_Error:
    DisplayError , "frm_search.lsv_batches_ItemClick", vbExclamation
End Sub

Private Sub lsv_batches_LostFocus()
    mod_search_result_key = 0
End Sub

Private Sub lsv_batches_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'***************************************
' Module/Form Name   : frm_search
'
' Procedure Name     : lsv_batches_MouseUp
'
' Purpose            :
'
' Date Created       : 18/12/2002
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
On Error GoTo lsv_batches_MouseUp_Error
'
'******** Code Starts Here *************
'

    Dim i
'
'   If right mouse button clicked and an item selected, display pop up menu.
'
    If Button <> vbRightButton Then
        Exit Sub
    End If
'
'   Display pop up menu.
'
    If opt_bar_code Then
        mdi_npls.mnuRemoveItem.Enabled = False
    Else
        mdi_npls.mnuRemoveItem.Enabled = True
    End If
    Set mdi_npls.fSearch = Me
    PopupMenu mdi_npls.mnuSearch, vbPopupMenuRightButton, , , mdi_npls.mnuSearchEditPhoto
'
'********* Code Ends Here **************
'
   Exit Sub
'
lsv_batches_MouseUp_Error:
    DisplayError , "frm_search.lsv_batches_MouseUp", vbExclamation
End Sub

Private Sub opt_bar_code_Click()
'***************************************
' Module/Form Name   : frm_search
'
' Procedure Name     : opt_bar_code_Click
'
' Purpose            :
'
' Date Created       : 18/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo opt_bar_code_Click_Error
'
'******** Code Starts Here *************
'
    If opt_bar_code.Value = True Then
        cmdAdd.Visible = False
        cmdAdd.Default = False
        cmdDelete.Visible = False
        cmdFind.Visible = False
        cmdConfirm.Visible = True
        cmdConfirm.Default = True
        cmdRefine.Visible = False
        fraAdd.Width = cmdConfirm.Left + cmdConfirm.Width + 100
        txt_batch_no.SetFocus
    End If
'
'********* Code Ends Here **************
'
   Exit Sub
'
opt_bar_code_Click_Error:
    DisplayError , "frm_search.opt_bar_code_Click", vbExclamation
End Sub

Private Sub opt_search_maint_Click()
'***************************************
' Module/Form Name   : frm_search
'
' Procedure Name     : opt_search_maint_Click
'
' Purpose            :
'
' Date Created       : 18/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo opt_search_maint_Click_Error
'
'******** Code Starts Here *************
'
    If opt_search_maint.Value = True Then
        fraAdd.Width = cmdDelete.Left + cmdDelete.Width + 100
        cmdConfirm.Visible = False
        cmdConfirm.Default = False
        cmdAdd.Visible = True
        cmdAdd.Default = True
        cmdDelete.Visible = True
        cmdFind.Visible = True
        cmdRefine.Visible = True
        txt_batch_no.SetFocus
    End If
'
'********* Code Ends Here **************
'
   Exit Sub
'
opt_search_maint_Click_Error:
    DisplayError , "frm_search.opt_search_maint_Click", vbExclamation
End Sub

Private Sub cmdAdd_Click()
'***************************************
' Module/Form Name   : frm_search
'
' Procedure Name     : cmdAdd_Click
'
' Purpose            :
'
' Date Created       : 23/04/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo cmdAdd_Click_Error
'
'******** Code Starts Here *************
'
    Dim oDigitalImage       As DigitalImage
    Dim oSearchToUpdate     As Search2
    
    '
    '   Put in a Doevents to ensure that any correction of a number by the Keypress event is completed.
    '
    DoEvents
    '
    If goSystemConfig.SageLink Then
        If Not moCustomer.TermsAndConditionsApproved Then
            MsgBox "Terms & Conditions have not yet been agreed for '" & moCustomer.CustomerName & "'." & vbCrLf & _
                   "You cannot therefore add any more Images.", vbExclamation
            Exit Sub
        End If
        If moCustomer.OnHold Then
            MsgBox "'" & moCustomer.CustomerName & "' is on hold." & vbCrLf & _
                   "You cannot therefore add any more Images.", vbExclamation
            Exit Sub
        End If
    End If
    
    If Not valid_input Then
        Exit Sub
    End If
    '
    If Not CheckUnconfirmPictures(pblnRedisplay:=True) Then
        Exit Sub
    End If
    
''    begin_trans
        If characters_after_decimal(txt_batch_no.Text) = "0" Then
            Set oDigitalImage = New DigitalImage
            oDigitalImage.PhotoTypeRequired = moSearch.PhotoType
            oDigitalImage.BatchNo = CLng(characters_before_decimal(txt_batch_no.Text))
            If Not ImageAvailable(oSearchToUpdate, oDigitalImage) Then
''                roll_back
                Exit Sub
            End If
        End If
        photograph.mblnMultipleAddition = False
        add_photographs txt_batch_no.Text, False, False, oSearchToUpdate
        Set oDigitalImage = Nothing
        redisplay_search_details
        '
''    commit_trans
'
'   Highlight image number.
'
    txt_batch_no.SetFocus
    txt_batch_no.SelStart = 0
    txt_batch_no.SelLength = Len(txt_batch_no.Text)

    update_loaded_screens
'
'********* Code Ends Here **************
'
   Exit Sub
'
cmdAdd_Click_Error:
    DisplayError , "frm_search.cmdAdd_Click", vbExclamation
End Sub

Private Function ImageAvailable(ByRef poSearchToUpdate As Search2, _
                                ByRef poDigitalImage As DigitalImage) As Boolean
'***************************************
' Module/Form Name   : frm_search
'
' Procedure Name     : ImageAvailable
'
' Purpose            :
'
' Date Created       : 29/06/2004
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
On Error GoTo ImageAvailable_Error
'
'******** Code Starts Here *************
'
    Dim blnAvailable        As Boolean
    Dim oDigitalSearch      As Search2
    Dim oSearchResult       As SearchResult
    
    blnAvailable = False
    '
    '   If Search Photo Type is Transparency then check if transparencies in stock.
    '
    If moSearch.PhotoType = "T" Then
        If Not photograph.photos_in_stock(poDigitalImage.BatchNo) Then
            '
            '   If no transparencies in stock, ask user if a digital image is required.
            '
            If MsgBox("There are no transparencies in stock." & vbCrLf & "Do you wish to send a digital image instead?", vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
                poDigitalImage.PhotoTypeRequired = "D"
                If poDigitalImage.FileLocation("A4") = "" Then
                    MsgBox "There is no 'A4' image of photograph '" & poDigitalImage.BatchNo & "'.", vbExclamation
                Else
                    blnAvailable = True
                End If
            End If
        Else
            blnAvailable = True
        End If
    Else
        If poDigitalImage.FileLocation("A4") = "" Then
            MsgBox "There is no 'A4' image of photograph '" & poDigitalImage.BatchNo & "'.", vbExclamation
        Else
            blnAvailable = True
        End If
    End If
    '
    If Not blnAvailable Then
        Exit Function
    End If
    '
    If moSearch.PhotoType <> poDigitalImage.PhotoTypeRequired Then
        '
        '   Search result needs to go on a different Search.
        '
        If moSearch.PhotoType = "T" Then
            '
            '   Has a Digital Search already been created for this Search?
            '
            Set oDigitalSearch = New Search2
            begin_trans
                '
                '   Create a new Search to hold the digital images.
                '
                With oDigitalSearch
                    .PhotoType = "D"
                    On Error Resume Next
                    .AddDAO moSearch.JobNo, moSearch.Description, moSearch.SearchDate
                    If Err.Number = vbObjectError + 1 Then
                        '
                        '   Search already exists, so use this one.
                        '
                        On Error GoTo ImageAvailable_Error
                        .Read plngJobNo:=moSearch.JobNo, _
                              pstrDescription:=moSearch.Description, _
                              pstrPhotoType:="D"
                    ElseIf Err.Number <> 0 Then
                        ErrorSave
                        On Error GoTo ImageAvailable_Error
                        ErrorRestore
                    Else
                        On Error GoTo ImageAvailable_Error
                    End If
                End With
                '
                '   Is the photograph already on the Digital Search?
                '
                For Each oSearchResult In oDigitalSearch.SearchResults
                    If oSearchResult.BatchNo = poDigitalImage.BatchNo Then
                        MsgBox "Digital Photograph '" & poDigitalImage.BatchNo & "' has already been selected for Search '" & oDigitalSearch.DisplayDescription & "'", vbExclamation
                        Exit Function
                    End If
                Next oSearchResult
                '
                moSearch.DigitalSearchNo = oDigitalSearch.SearchNo
                moSearch.update
            commit_trans
            Set poSearchToUpdate = oDigitalSearch
            '
        Else
            '
            '   Something similar may need to go in here for when a digital image isn't available but a Transparency is.
            '
            MsgBox "The scenario whereby a transparency is required on a Digital Search is not currently catered for.", vbExclamation
        End If
    Else
        Set poSearchToUpdate = moSearch
    End If
    '
    ImageAvailable = blnAvailable
'
'********* Code Ends Here **************
'
   Exit Function
'
ImageAvailable_Error:
    ErrorRaise "frm_search.ImageAvailable"
End Function

Public Sub add_photographs(ByVal pstrBatchNo As String, _
                           ByVal pblnConfirmed As Boolean, _
                           Optional ByVal blnUpdateOtherScreens As Boolean = True, _
                           Optional ByRef poSearch As Search2)
'***************************************
' Module/Form Name   : frm_search
'
' Procedure Name     : add_photographs
'
' Purpose            :
'
' Date Created       : 18/12/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : pstrBatchNo - String
'                    : Confirmed - Boolean
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo add_photographs_Error
'
'******** Code Starts Here *************
'
    
    Dim Cancel                  As Boolean
    Dim var_search_result_key   As Long
    Dim lngBatchNo              As Long
    Dim lngPhotographNo         As Long
    Dim lngPhotographKey        As Long
    Dim blnScanned              As Boolean
    Dim itmX                    As MSComctlLib.ListItem
    Dim oSearchResult           As SearchResult
    Dim blnTransparency         As Boolean
    Dim oSearchToAddTo          As Search2
    Dim strSQL                  As String
    Dim lngPhotographSwapNo     As Long
    Dim oSearchResultSwap       As SearchResult
    
    If poSearch Is Nothing Then
        blnTransparency = True
    Else
        If poSearch.PhotoType = "T" Then
            blnTransparency = True
        Else
            blnTransparency = False
        End If
    End If
    '
    If poSearch Is Nothing Then
        Set oSearchToAddTo = moSearch
    Else
        Set oSearchToAddTo = poSearch
    End If
    '
    '   Extract the photograph and image number from the entry.
    '
    lngBatchNo = CLng(characters_before_decimal(pstrBatchNo))
    lngPhotographNo = CLng(characters_after_decimal(pstrBatchNo))
    
    If lngPhotographNo = 0 Then
        blnScanned = False
    Else
        blnScanned = True
    End If
    '
    begin_trans
        '
        '   Get the batch record.
        '
        SQL = "SELECT batch_no, description " _
             & "From batch " _
             & "Where batch_no = " & CStr(lngBatchNo)
        Set snap_batch_search = Nothing
        Set snap_batch_search = db.OpenRecordset(SQL, dbOpenSnapshot)
        
        If snap_batch_search.RecordCount = 0 Then
            MsgBox "Batch does not exist", vbExclamation
            close_batch_recordsets
            roll_back
            Exit Sub
        End If
        '
        '   Set up the Search Result to populate.
        '
        Set oSearchResult = New SearchResult
        With oSearchResult
            .SearchNo = oSearchToAddTo.SearchNo
            .JobDescription = moJob.reference
            .BatchNo = lngBatchNo
            .PhotographNo = lngPhotographNo
            .CustomerNo = mod_customer_no
            .Confirmed = pblnConfirmed
            .PhotoType = IIf(blnTransparency, "T", "D")
            .PhotoDesc = snap_batch_search!Description
        End With
        snap_batch_search.Close
        Set snap_batch_search = Nothing
        '
        '   Photograph: if an image number has not been entered then get the first
        '   in stock else attempt to get the precise image.
        '
        Call photograph.make_pending(oSearchResult, Cancel, lngPhotographSwapNo)
        If Cancel Then
            close_batch_recordsets
            roll_back
            Exit Sub
        End If
        '
        '   Search Result
        '
        On Error Resume Next
        oSearchResult.CreateDAO
        If Err.Number = vbObjectError + 3 Then
            MsgBox "Photograph '" & CStr(oSearchResult.BatchNo) & "' already exists on Search '" & oSearchToAddTo.DisplayDescription & "'", vbInformation
            roll_back
            Exit Sub
        ElseIf Err.Number = vbObjectError + 2 Then
            MsgBox Err.Description, vbExclamation
            Exit Sub
        ElseIf Err.Number <> 0 Then
            ErrorSave
            On Error GoTo add_photographs_Error
            ErrorRestore
        Else
            On Error GoTo add_photographs_Error
        End If
        '
        '   Now add to the Search Results collection of the Search object.
        '   Ignore any errors as thsi will have ben caused by an automatic refresh placing the
        '   object on the collection already in the case that the SearchResults collection hasn't been
        '   initialised.
        '
        On Error Resume Next
        With oSearchToAddTo.SearchResults
            '
            '   If an image has been swapped then find it in the collection etc.
            '
            If lngPhotographSwapNo <> 0 Then
                For Each oSearchResultSwap In oSearchToAddTo.SearchResults
                    If oSearchResultSwap.BatchNo = oSearchResult.BatchNo And _
                       oSearchResultSwap.PhotographNo = oSearchResult.PhotographNo Then
                        oSearchResultSwap.PhotographNo = lngPhotographSwapNo
                        Exit For
                    End If
                Next oSearchResultSwap
            End If
            .AddObject oSearchResult
            .Sort
        End With
        On Error GoTo add_photographs_Error
        '
        '   If there is no Refine Activity for this Search yet then create one.
        '
        oSearchToAddTo.Activities.ActivityType = "REFS"
        oSearchToAddTo.Activities.Refresh
        If oSearchToAddTo.Activities.Count = 0 Then
            CreateDigitialActivity oSearchToAddTo, "REFS"
        End If
    commit_trans
        
    close_batch_recordsets
    
    If blnUpdateOtherScreens Then
        update_loaded_screens
    End If
'
'********* Code Ends Here **************
'
   Exit Sub
'
add_photographs_Error:
    ErrorRaise "frm_search.add_photographs"
End Sub

Private Sub close_batch_recordsets()
'***************************************
' Module/Form Name   : frm_search
'
' Procedure Name     : close_batch_recordsets
'
' Purpose            :
'
' Date Created       : 18/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo close_batch_recordsets_Error
'
'******** Code Starts Here *************
'
    snap_batch_search.Close
    dyna_photograph.Close
    table_search_result.Close
    dyna_search_result.Close
    dyna_search.Close
    dyna_job.Close

'
'********* Code Ends Here **************
'
   Exit Sub
'
close_batch_recordsets_Error:
    Select Case Err.Number
        Case Is = 91
            Resume Next
        Case Is = 3420
            Resume Next
        Case Else
            DisplayError , "frm_search.close_batch_recordsets", vbExclamation
    End Select
End Sub

Private Sub get_snap_customer()
'***************************************
' Module/Form Name   : frm_search
'
' Procedure Name     : get_snap_customer
'
' Purpose            :
'
' Date Created       : 18/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo get_snap_customer_Error
'
'******** Code Starts Here *************
'
    SQL = "select * from customer " _
    & "where customer_no = " + Str(mod_customer_no) + ";"

    Set snap_customer = Nothing
    Set snap_customer = db.OpenRecordset(SQL, dbOpenSnapshot)

'
'********* Code Ends Here **************
'
   Exit Sub
'
get_snap_customer_Error:
    ErrorRaise "frm_search.get_snap_customer"
End Sub

Public Sub display_search_details(search_no As Long)
    '***************************************
    ' Module/Form Name   : frm_search
    '
    ' Procedure Name     : display_search_details
    '
    ' Purpose            :
    '
    ' Date Created       : 01/12/2002
    '
    ' Author             : GARETH SAUNDERS
    '
    ' Parameters         : search_no - Long
    '
    ' Amendment History  : Date       Author    Description
    '                    : --------------------------------
    '                    : 01/12/2002 GARETH SAUNDERS
    '
    '***************************************
    '
On Error GoTo display_search_details_Error
    '
    '******** Code Starts Here *************
    '
    Dim counter As Integer
    Dim itmX As MSComctlLib.ListItem
    Dim oSearchResult           As SearchResult
    
    Screen.MousePointer = vbHourglass

    LockWindow Me.hWnd
    mod_search_no = search_no

    DoEvents
    '
    '   Set up Form Title.
    '
    display_title
        '
''    get_snap_search_result
    '
    '   Use the search results collection instead now.
    '
    txt_search_description = moSearch.DisplayDescription

    DoEvents

    lsv_batches.ListItems.Clear
    lsv_batches.SmallIcons = iml_ticks
    lsv_batches.view = lvwReport

    For Each oSearchResult In moSearch.SearchResults
        With oSearchResult
            Set itmX = lsv_batches.ListItems.Add()
            '
            '   If Search is a Digital one then don't display the Photograph Number.
            '
            If moSearch.PhotoType = "D" Then
                itmX.SubItems(1) = CStr(.BatchNo)
            Else
                itmX.SubItems(1) = CStr(.BatchNo) & "." & CStr(.PhotographNo)
            End If
            itmX.SubItems(2) = .PhotoDesc
            itmX.Tag = CStr(.SearchResultKey)
            If .Confirmed = True Then
                itmX.SmallIcon = 1
            End If
        End With
    Next oSearchResult
    With lsv_batches
        If .SortKey = 1 Then
            SortImages lsv_batches.SortOrder
        Else
            .Sorted = True
        End If
    End With
    '
    '   Make sure the first one is not selected. It seems to be in the code but not on the screen!
    '
    If lsv_batches.ListItems.Count <> 0 Then
        lsv_batches.ListItems(1).Selected = False
    End If
    '
    display_totals
    '
    EnableButtons
    Screen.MousePointer = vbDefault
    UnlockWindow
'
'********* Code Ends Here **************
'
    Exit Sub
          '
display_search_details_Error:
    ErrorRaise "frm_search.display_search_details"
End Sub

Public Sub display_search_form(Action As UpdateMode, _
                               poJob As Job2, _
                               plngSearchNo As Long)
'***************************************
' Module/Form Name   : frm_search
'
' Procedure Name     : display_search_form
'
' Purpose            :
'
' Date Created       : 27/05/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : Action - String
'                    : poSearch - Search2
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 30/11/2002 GARETH SAUNDERS
'                    : 18/12/2002 GARETH SAUNDERS
'
'***************************************
'
    On Error GoTo display_search_form_Error
'
'******** Code Starts Here *************
'
'    LockWindow Me.hWnd
    Dim f As Form
    
    '
    '   Set up job details.
    '
''    GetJob
    Set moJob = poJob
    Set moSearch = moJob.Searches.Item(CStr(plngSearchNo))
    '
    mod_action = Action
    mod_customer_no = moJob.CustomerNo
    mod_job_no = moJob.JobNo
    mod_search_no = moSearch.SearchNo
''    GetSearch
    '
    If moSearch.PhotoType = "D" Then
        opt_bar_code.Enabled = False
    End If
    '
    '   Set up customer details.
    '
    get_snap_customer
    Set moCustomer = Nothing
    Set moCustomer = New Customer2
    moCustomer.Read mod_customer_no
    txt_customer_name = moCustomer.CustomerName
    txt_address_line_1 = moCustomer.Address1
    '
    txtJobNo.Text = moJob.JobNo

    If moSearch.PhotoType = "D" Then
        cmdConfirmDigitalImages.Visible = True
        With moSearch.Activities
            .ActivityType = "REFS"
            .SearchNo = moSearch.SearchNo
            .Refresh
        End With
    End If
    
    Select Case mod_action
        Case Is = Add, Edit
            '
            '   On edit set up label for key description field
            '   and display current search details
            '
            display_search_details (mod_search_no)
            Form_Resize
            DoEvents
            com_position_form Me
            For Each f In Forms
                If f.Name <> "mdi_npls" Then
                    If f.WindowState = vbMaximized Then
                        Me.WindowState = vbMaximized
                        Exit For
                    End If
                End If
            Next f
            Me.Show
            txt_batch_no.SetFocus
        Case Is = view
            opt_bar_code.Enabled = False
            opt_search_maint.Enabled = False
            cmdAdd.Enabled = False
            cmdDelete.Enabled = False
            cmdFind.Enabled = False
            com_change.Enabled = False
            cmdRefine.Enabled = True
            cmdRefine.Caption = "&Review"
            display_search_details (mod_search_no)
            com_position_form Me
            Me.Show
    End Select
    '
''    EnableButtons
'    UnlockWindow
    '
    '********* Code Ends Here **************
    '
Exit Sub
    '
display_search_form_Error:
    ErrorRaise "frm_search.display_search_form"
End Sub

Private Sub GetJob()
    '***************************************
    ' Module/Form Name   : frm_search
    '
    ' Procedure Name     : GetJob
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
On Error GoTo GetJob_Error
    '
    '******** Code Starts Here *************
    '
    Set moJob = Nothing
    Set moJob = New Job2

    moJob.Read moSearch.JobNo
    '
    '********* Code Ends Here **************
    '
    Exit Sub
        '
GetJob_Error:
    ErrorRaise "frm_search.GetJob"
End Sub
''
''Private Sub GetSearch()
''          '***************************************
''          ' Module/Form Name   : frm_search
''          '
''          ' Procedure Name     : GetSearch
''          '
''          ' Purpose            :
''          '
''          ' Date Created       : 01/12/2002
''          '
''          ' Author             : GARETH SAUNDERS
''          '
''          ' Amendment History  : Date       Author    Description
''          '                    : --------------------------------
''          '                    : 01/12/2002 GARETH SAUNDERS
''          '
''          '***************************************
''          '
''      On Error GoTo GetSearch_Error
''          '
''          '******** Code Starts Here *************
''          '
''
''      Set moSearch = Nothing
''      Set moSearch = New Search2
''
''      moSearch.Read mod_search_no
''
''          '
''          '********* Code Ends Here **************
''          '
''      Exit Sub
''          '
''GetSearch_Error:
''    ErrorRaise "frm_search.GetSearch"
''End Sub
''
''Private Sub get_snap_search_result()
'''***************************************
''' Module/Form Name   : frm_search
'''
''' Procedure Name     : get_snap_search_result
'''
''' Purpose            :
'''
''' Date Created       : 18/12/2002
'''
''' Author             : GARETH SAUNDERS
'''
''' Amendment History  : Date       Author    Description
'''                    : --------------------------------
'''
'''***************************************
'''
''On Error GoTo get_snap_search_result_Error
'''
'''******** Code Starts Here *************
'''
''    SQL = "SELECT search_result.*, batch.description " _
''        & "FROM search_result, photograph, batch " _
''        & "WHERE search_no = " + Str(mod_search_no) + "and " _
''        & "photograph.photograph_key = search_result.photograph_key " _
''        & "and batch.batch_no = photograph.batch_no and photograph.deleted =false "
''    SQL = SQL & "ORDER BY SEARCH_RESULT.BATCH_NO, SEARCH_RESULT.PHOTOGRAPH_NO"
''    Set snap_search_result = db.OpenRecordset(SQL, dbOpenSnapshot)
''
'''
'''********* Code Ends Here **************
'''
''   Exit Sub
'''
''get_snap_search_result_Error:
''    ErrorRaise "frm_search.get_snap_search_result"
''End Sub

Private Sub com_change_Click()
'***************************************
' Module/Form Name   : frm_search
'
' Procedure Name     : com_change_Click
'
' Purpose            :
'
' Date Created       : 18/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo com_change_Click_Error
'
'******** Code Starts Here *************
'
    Set fSearchDetails = New frm_search_details
    If mod_action = Edit Then
        On Error Resume Next
        fSearchDetails.load_search_title moJob, mod_action, moSearch
        If Err.Number = vbObjectError + 2 Then
            '
            '   Search has been deleted.
            '
            MsgBox Err.Description, vbExclamation, App.Title
            Unload Me
            Exit Sub
        ElseIf Err.Number = 0 Then
            On Error GoTo com_change_Click_Error
        Else
            ErrorSave
            On Error GoTo com_change_Click_Error
            ErrorRestore
        End If
    Else
        fSearchDetails.load_search_title moJob, mod_action, moSearch
    End If
    Set fSearchDetails = Nothing
    txt_batch_no.SetFocus
'
'********* Code Ends Here **************
'
   Exit Sub
'
com_change_Click_Error:
    DisplayError , "frm_search.com_change_Click", vbExclamation
End Sub

Private Sub Form_Load()
'***************************************
' Module/Form Name   : frm_search
'
' Procedure Name     : Form_Load
'
' Purpose            :
'
' Date Created       : 18/12/2002
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
''    Resize.Hook Me.hWnd, 600, 400
    mbolDontResize = True
'
    With lsv_batches
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , ""
        .ColumnHeaders.Add , , "Photo", , lvwColumnRight
        .ColumnHeaders.Add , , "Description", , lvwColumnLeft
        .SortKey = GetSetting(App.Title, "SearchMaintenance", "SortKey", 2)
        .SortOrder = GetSetting(App.Title, "SearchMaintenance", "SortOrder", lvwDescending)
        .Sorted = True
    End With
'
'   Set the width & height of the form.
'
    Me.WindowState = GetSetting(App.Title, "SearchMaintenance", "WindowState", 0)
    If GetSetting(App.Title, "SearchMaintenance", "WindowState", 0) <> 2 Then
        Me.Width = GetSetting(App.Title, "SearchMaintenance", "Width", 8025)
        Me.Height = GetSetting(App.Title, "SearchMaintenance", "Height", 6000)
    End If
    '
    '   Allow resizing to take place now as this is the last modification.
    '
    mbolDontResize = False
    mod_search_result_key = 0
''    com_position_form Me
''    DoEvents
''    Me.Show
'
'********* Code Ends Here **************
'
   Exit Sub
'
Form_Load_Error:
    DisplayError , "frm_search.Form_Load", vbExclamation
End Sub

Private Sub cmdconfirm_Click()
'***************************************
' Module/Form Name   : frm_search
'
' Procedure Name     : cmdconfirm_Click
'
' Purpose            :
'
' Date Created       : 18/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo cmdconfirm_Click_Error
'
'******** Code Starts Here *************
'
    Dim Confirmed                   As Boolean
    Dim possible_swap_image         As Long
    Dim possible_swap_image_index   As Integer
    Dim oSearchResult               As SearchResult
    Dim itmX                        As ListItem
    
    '
    '   Put in a Doevents to ensure that any correction of a number by the Keypress event is completed.
    '
    DoEvents
    '
    If goSystemConfig.SageLink Then
        If Not moCustomer.TermsAndConditionsApproved Then
            MsgBox "Terms & Conditions have not yet been agreed for '" & moCustomer.CustomerName & "'." & vbCrLf & _
                   "You cannot therefore confirm any more Images.", vbExclamation
            Exit Sub
        End If
        If moCustomer.OnHold Then
            MsgBox "'" & moCustomer.CustomerName & "' is on hold." & vbCrLf & _
                   "You cannot therefore confirm any more Images.", vbExclamation
            Exit Sub
        End If
    End If
    '
    If Not valid_input Then
        Exit Sub
    End If
'
'   Is the confirm image in the list box.
'
    Dim i
    Dim search_photograph_no As String
    Dim search_photograph_image_no As String
    Dim txt_photograph_no As String
    Dim txt_image_no As String
    Dim txt_description As String

    txt_photograph_no = characters_before_decimal(txt_batch_no)
    txt_image_no = RTrim(characters_after_decimal(txt_batch_no))
'
'   Find exact match.
'
    possible_swap_image = 0
    With lsv_batches
        For i = 1 To .ListItems.Count
            search_photograph_no = characters_before_decimal(LTrim(.ListItems(i).SubItems(1)))
            txt_description = LTrim(characters_after_space(LTrim(.ListItems(i).SubItems(1))))
            search_photograph_image_no = characters_before_space(LTrim(.ListItems(i).SubItems(1)))
            If search_photograph_no = txt_photograph_no Then
                If search_photograph_image_no <> LTrim(txt_batch_no.Text) Then
'
'                   Right Photograph, wrong image. Mark it for possible swapping later.
'
                    If Not search_result.Confirmed(.ListItems(i).Tag) Then
                        If possible_swap_image = 0 Then
                            possible_swap_image = .ListItems(i).Tag
                            possible_swap_image_index = i
                        End If
                    End If
                Else
                    With .ListItems(i)
                        Set oSearchResult = moSearch.SearchResults.Item(.Tag)
                        With oSearchResult
                            .Confirmed = Not .Confirmed
                            '
                            '   If the display is out of sync with the database
                            '   then refresh the collection.
                            '
                            On Error Resume Next
                            .UpdateDAO
                            If Err.Number - vbObjectError = 1 Then
                                MsgBox "Unable to find Photograph on Search. The Search may have been updated since the last display." & vbCrLf & _
                                       "The screen will now be refreshed", vbInformation
                                moSearch.SearchResults.Refresh
                                redisplay_search_details
                                Exit Sub
                            ElseIf Err.Number <> 0 Then
                                ErrorSave
                                On Error GoTo cmdconfirm_Click_Error
                                ErrorRestore
                            Else
                                On Error GoTo cmdconfirm_Click_Error
                            End If
                        End With
                        moJob.Searches.PendingRefresh = True
''                        search_result.confirm .Tag, Confirmed
''                        On Error Resume Next
''                        moSearch.SearchResults.Item(.Tag).Confirmed = Confirmed
''                        On Error GoTo cmdconfirm_Click_Error
                        '
                        '   Clear all previously selected Images.
                        '
                        For Each itmX In lsv_batches.ListItems
                            itmX.Selected = False
                        Next itmX
                        If oSearchResult.Confirmed = True Then
                            .SmallIcon = 1
                            .EnsureVisible
                            .Selected = True
                        Else
                            .SmallIcon = 0
                            .EnsureVisible
                            .Selected = True
                        End If
                        update_loaded_screens
                        GoTo highlight_image_no
                    End With
                End If
            End If
        Next i
'
'   Exact match not found. If possible swap found, swap it.
'
        If possible_swap_image <> 0 Then
            Set oSearchResult = moSearch.SearchResults.Item(CStr(possible_swap_image))
            photo_image.change_image_number possible_swap_image, _
                                            LTrim(txt_batch_no.Text), _
                                            mod_search_no, _
                                            oSearchResult, _
                                            abort
            If abort Then
                On Error Resume Next
                txt_batch_no.SetFocus
                txt_batch_no.SelStart = 0
                txt_batch_no.SelLength = Len(txt_batch_no.Text)
                Exit Sub
            End If
            search_result.confirm possible_swap_image, Confirmed
            oSearchResult.Confirmed = Not oSearchResult.Confirmed
            '
            '   Deselect all lines first.
            '
            For i = 1 To .ListItems.Count
                .ListItems.Item(i).Selected = False
            Next i
            '
''            .ListItems(possible_swap_image_index).SubItems(1) = _
''                    txt_photograph_no & "." & txt_image_no
''            .ListItems(possible_swap_image_index).SubItems(2) = photograph.Read(txt_photograph_no, dbOpenSnapshot)!Description
''            .ListItems(possible_swap_image_index).SmallIcon = 1
''            .ListItems(possible_swap_image_index).EnsureVisible
''            .ListItems(possible_swap_image_index).Selected = True
            update_loaded_screens
            GoTo highlight_image_no
        End If
    End With
'
'   Image not on list. Add to list in a confirmed state.
'
    photograph.mblnMultipleAddition = False
    add_photographs txt_batch_no, True, , moSearch
'
highlight_image_no:
    redisplay_search_details
    txt_batch_no.SetFocus
    txt_batch_no.SelStart = 0
    txt_batch_no.SelLength = Len(txt_batch_no.Text)
    '
''    update_loaded_screens
    '
    display_totals
    '
    '   Repaint.
    '
    Me.Refresh
    lsv_batches.Refresh
'
'********* Code Ends Here **************
'
   Exit Sub
'
cmdconfirm_Click_Error:
    DisplayError , "frm_search.cmdconfirm_Click", vbExclamation
End Sub

Private Sub cmdExit_Click()
'***************************************
' Module/Form Name   : frm_search
'
' Procedure Name     : cmdExit_Click
'
' Purpose            :
'
' Date Created       : 18/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo cmdExit_Click_Error
'
'******** Code Starts Here *************
'
    On Error Resume Next

''    If is_form_loaded("frm_job_maint") Then
''        frm_job_maint.redisplay
''    End If
    
    If is_form_loaded("frm_job_edit") Then
        frm_job_edit.display_job snap_customer!customer_no, _
                                 moJob.JobNo, _
                                 moJob, _
                                 mod_search_no
    Else
        frm_job_edit.display_job snap_customer!customer_no, _
                                 moJob.JobNo, _
                                 moJob, _
                                 mod_search_no
    End If
    
    Unload Me
    DoEvents
'
'********* Code Ends Here **************
'
   Exit Sub
'
cmdExit_Click_Error:
    DisplayError , "frm_search.cmdExit_Click", vbExclamation
End Sub

Private Sub cmdDelete_Click()
'***************************************
' Module/Form Name   : frm_search
'
' Procedure Name     : cmdDelete_Click
'
' Purpose            :
'
' Date Created       : 18/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo cmdDelete_Click_Error
'
'******** Code Starts Here *************
'
    Dim i
    Dim blnNoneSelected As Boolean
    Dim oActivity       As Activity
    
    i = lsv_batches.ListItems.Count
    blnNoneSelected = True
    Do While i <> 0
        If lsv_batches.ListItems(i).Selected Then
            blnNoneSelected = False
            Exit Do
        End If
        i = i - 1
    Loop

    If blnNoneSelected = True Then
        MsgBox "Select an Image to remove", vbExclamation
        Exit Sub
    End If
    If MsgBox("Are you sure you wish to remove the selected Images?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
        Exit Sub
    End If
    '
    begin_trans
        CheckUnconfirmPictures pblnRedisplay:=False, pblnAddRefine:=False
        '
        Do While i <> 0
            If lsv_batches.ListItems(i).Selected Then
                If moSearch.SearchResults.Item(lsv_batches.ListItems(i).Tag).Confirmed Then
    ''            If lsv_batches.ListItems(i).SmallIcon = 1 Then
                    MsgBox "You cannot Remove Confirmed Photographs.", vbExclamation
                    Exit Do
                End If
                RemovePhotograph (lsv_batches.ListItems(i).Tag)
                '
                '   Remove the batch from the listbox.
                '
                lsv_batches.ListItems.Remove i
            End If
            i = i - 1
        Loop
        '
        If moSearch.SearchResults.Count <> 0 Then
            CreateDigitialActivity moSearch, "REFS"
        Else
            '
            '   Close all activities.
            '
            moSearch.Activities.ActivityType = ""
            moSearch.Activities.Refresh
            For Each oActivity In moSearch.Activities
                With oActivity
                    .Description = "All Search Results removed from Search"
                    .EndDate = Now()
                    .Status = "3"
                    .update
                End With
            Next oActivity
        End If
    commit_trans
    '
    redisplay_search_details
    '
    display_totals

    update_loaded_screens
    
    EnableButtons
'
'********* Code Ends Here **************
'
   Exit Sub
'
cmdDelete_Click_Error:
    DisplayError , "frm_search.cmdDelete_Click", vbExclamation
End Sub

Private Sub cmdFind_Click()
'***************************************
' Module/Form Name   : frm_search
'
' Procedure Name     : cmdFind_Click
'
' Purpose            :
'
' Date Created       : 18/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo cmdFind_Click_Error
'
'******** Code Starts Here *************
'
    If goSystemConfig.SageLink Then
        If Not moCustomer.TermsAndConditionsApproved Then
            MsgBox "Terms & Conditions have not yet been agreed for '" & moCustomer.CustomerName & "'." & vbCrLf & _
                   "You cannot therefore find any Images.", vbExclamation
            Exit Sub
        End If
        If moCustomer.OnHold Then
            MsgBox "'" & moCustomer.CustomerName & "' is on hold." & vbCrLf & _
                   "You cannot therefore find any Images.", vbExclamation
            Exit Sub
        End If
    End If
    '
    Screen.MousePointer = vbHourglass
    Set fSearchPhotograph = Nothing
    Set fSearchPhotograph = New frm_search_photograph
    fSearchPhotograph.display_screen srSsearch, moSearch, moJob
    redisplay_search_details
    Set fSearchPhotograph = Nothing
    On Error Resume Next
    Me.SetFocus
    On Error GoTo cmdFind_Click_Error
    Screen.MousePointer = vbDefault
'
'********* Code Ends Here **************
'
   Exit Sub
'
cmdFind_Click_Error:
    DisplayError , "frm_search.cmdFind_Click", vbExclamation
End Sub

Private Sub txt_batch_no_GotFocus()
'***************************************
' Module/Form Name   : frm_search
'
' Procedure Name     : txt_batch_no_GotFocus
'
' Purpose            :
'
' Date Created       : 18/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo txt_batch_no_GotFocus_Error
'
'******** Code Starts Here *************
'
    If opt_search_maint.Value = True Then
        cmdAdd.Default = True
    Else
        cmdConfirm.Default = True
    End If

    txt_batch_no.SelStart = 0
    txt_batch_no.SelLength = Len(txt_batch_no.Text)
'
'********* Code Ends Here **************
'
   Exit Sub
'
txt_batch_no_GotFocus_Error:
    DisplayError , "frm_search.txt_batch_no_GotFocus", vbExclamation
End Sub

Private Sub txt_batch_no_KeyPress(KeyAscii As Integer)
          If (KeyAscii < 48 Or KeyAscii > 57) And _
              Not KeyAscii = 8 And _
              Not KeyAscii = 46 Then
              KeyAscii = 0
              Beep
          End If
End Sub


Private Sub txt_search_description_Change()
'***************************************
' Module/Form Name   : frm_search
'
' Procedure Name     : txt_search_description_Change
'
' Purpose            :
'
' Date Created       : 18/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo txt_search_description_Change_Error
'
'******** Code Starts Here *************
'
    If mod_action = view Then Exit Sub

    If Len(RTrim(txt_search_description)) = 0 Then
        cmdAdd.Enabled = False
        cmdDelete.Enabled = False
        cmdFind.Enabled = False
        txt_batch_no.Enabled = False
    Else
        cmdAdd.Enabled = True
        cmdDelete.Enabled = True
        cmdFind.Enabled = True
        txt_batch_no.Enabled = True
    End If
'
'********* Code Ends Here **************
'
   Exit Sub
'
txt_search_description_Change_Error:
    DisplayError , "frm_search.txt_search_description_Change", vbExclamation
End Sub

Private Sub RemovePhotograph(Optional ByVal plngSearchResultKey As Long = 0, _
                             Optional ByVal plngBatchNo As Long = 0)
'***************************************
' Module/Form Name   : frm_search
'
' Procedure Name     : RemovePhotograph
'
' Purpose            :
'
' Date Created       : 18/12/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : lngSearchResultKey - Long
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo RemovePhotograph_Error
'
'******** Code Starts Here *************
'
    Dim oSearchResult               As SearchResult
    Dim oPhotograph                 As Photograph2
    Dim intIdx                      As Integer
    Dim strMessage                  As String
    
    begin_trans
''    gdbADO.BeginTrans
        '
        '   Find Search Results.
        '
        For intIdx = moSearch.SearchResults.Count To 1 Step -1
            Set oSearchResult = moSearch.SearchResults.Item(intIdx)
            If (plngSearchResultKey <> 0 And oSearchResult.SearchResultKey = plngSearchResultKey) Or _
               (plngBatchNo <> 0 And oSearchResult.BatchNo = plngBatchNo) Then
                '
                '   Photograph
                '
                Set oPhotograph = New Photograph2
                oPhotograph.Read oSearchResult.BatchNo, oSearchResult.PhotographNo
                If oPhotograph.BatchNo = 0 Then
                    MsgBox "Database error: contact your supervisor", vbCritical
                    roll_back
''                    gdbADO.RollbackTrans
                    Exit Sub
                End If
                '
                With oPhotograph
                    '
                    '   Only check status if a transparency. That is, the photograph number is non zero.
                    '
                    If .PhotographNo <> 0 Then
                        If .Status = "O" Then
                            MsgBox "Photograph has been delivered. You cannot remove it from this Search", vbExclamation
                            roll_back
''                            gdbADO.RollbackTrans
                            Exit Sub
                        Else
                            If .Status = "I" Then
                                If goSystemConfig.SupportUser Then
                                    '
                                    '   Allow the correction of this sort of thing if a support user.
                                    '
                                    strMessage = "Unexpected error: photograph " & .BatchNo & "." & .PhotographNo & " is not recorded as Pending" & vbCrLf & _
                                                 "As an Administrative user, do you wish to correct the data?"
                                    If MsgBox(strMessage, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                                        roll_back
''                                        gdbADO.RollbackTrans
                                        Exit Sub
                                    End If
                                Else
                                    Err.Raise 513, "frm_Search: RemovePhotograph", "Unexpected error: photograph " & .BatchNo & "." & .PhotographNo & " is not recorded as Pending"
                                    MsgBox "Unexpected error: photograph " & .BatchNo & "." & .PhotographNo & " is not recorded as Pending", vbCritical
                                    roll_back
''                                    gdbADO.RollbackTrans
                                    Exit Sub
                                End If
                            End If
                        End If
                        .Status = "I"
                        .update
                    End If
                End With
                '
                '   Delete Search Result
                '
                oSearchResult.Delete
                moSearch.SearchResults.Remove intIdx
            End If
        Next intIdx
        commit_trans
''    gdbADO.CommitTrans

'
'********* Code Ends Here **************
'
   Exit Sub
'
RemovePhotograph_Error:
    ErrorRaise "frm_search.RemovePhotograph"
End Sub

Public Sub ForceResize()
    gblnResizeMaxedWindows = False
    Form_Resize
    gblnResizeMaxedWindows = True
End Sub

