VERSION 5.00
Begin VB.Form frm_search_details 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Details"
   ClientHeight    =   1800
   ClientLeft      =   1815
   ClientTop       =   2580
   ClientWidth     =   6390
   Icon            =   "Search Details.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1800
   ScaleWidth      =   6390
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   555
      Left            =   4560
      Picture         =   "Search Details.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      UseMaskColor    =   -1  'True
      Width           =   795
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   555
      Left            =   5460
      Picture         =   "Search Details.frx":09BC
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      UseMaskColor    =   -1  'True
      Width           =   795
   End
   Begin VB.TextBox txt_search_description 
      Height          =   795
      Left            =   1680
      MaxLength       =   40
      MultiLine       =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "Enter the search title e.g. 'Chapter One'."
      Top             =   120
      Width           =   4635
   End
   Begin VB.Label lblSearchTitle 
      Caption         =   "Search Title:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frm_search_details"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private SQL As String
Private retry_update As Boolean

Private snap_search As DAO.Recordset
Private dyna_search As DAO.Recordset
Private dyna_job As DAO.Recordset

Private mod_search_no As Long
Private mod_job_no As Long
Private mod_action As UpdateMode
Public Event TitleChanged()
Public Event TitleCreated(SearchNo As Long)
Private moJob           As Job2
Private moSearch        As Search2

Private Sub close_search_recordsets()
'***************************************
' Module/Form Name   : frm_search_details
'
' Procedure Name     : close_search_recordsets
'
' Purpose            :
'
' Date Created       : 18/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 18/12/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo close_search_recordsets_Error
'
'******** Code Starts Here *************
'

          dyna_search.Close
          dyna_job.Close

'
'********* Code Ends Here **************
'
   Exit Sub
'
close_search_recordsets_Error:
    Select Case Err.Number
        Case Is = 91
            Resume Next
        Case Is = 3420
            Resume Next
        Case Else
            ErrorRaise "frm_search_details.close_search_recordsets"
    End Select
End Sub

Private Sub create_search()
'***************************************
' Module/Form Name   : frm_search_details
'
' Procedure Name     : create_search
'
' Purpose            :
'
' Date Created       : 30/11/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 30/11/2002 GARETH SAUNDERS
'                    : 18/12/2002 GARETH SAUNDERS
'
'***************************************
'
    On Error GoTo create_search_Error
    '
    '******** Code Starts Here *************
    '
    retry_update = True

    Set moSearch = New Search2
    moSearch.PhotoType = moJob.PreferredPhotoType
    moSearch.Add mod_job_no, _
                 txt_search_description.Text, _
                 Date
    moJob.Searches.AddObject moSearch
    mod_search_no = moSearch.SearchNo

    retry_update = False
    '
    '********* Code Ends Here **************
    '
    Exit Sub
    '
create_search_Error:
    Select Case Err.Number
        Case Is = vbObjectError + 1
            MsgBox "Search '" & txt_search_description.Text & "' already exists for this Job", vbExclamation
            close_search_recordsets
            roll_back
        Case Else
            ErrorRaise "frm_search_details.create_search"
            close_search_recordsets
    End Select
End Sub

Private Function valid_input() As Boolean
'***************************************
' Module/Form Name   : frm_search_details
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
    valid_input = True
    If Len(LTrim(txt_search_description)) = 0 Then
        MsgBox "Please enter a Search Description", vbExclamation
        txt_search_description.SetFocus
        valid_input = False
    End If
'
'********* Code Ends Here **************
'
   Exit Function
'
valid_input_Error:
    ErrorRaise "frm_search_details.valid_input"
End Function

Private Function search_changed() As Boolean
'***************************************
' Module/Form Name   : frm_search_details
'
' Procedure Name     : search_changed
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
On Error GoTo search_changed_Error
'
'******** Code Starts Here *************
'
    Dim i As Integer
    search_changed = False
    For i = 0 To snap_search.Fields.Count - 1
        If Not snap_search.Fields(i).Value = dyna_search.Fields(i).Value Then
            search_changed = True
            Exit For
        End If
    Next i
'
'********* Code Ends Here **************
'
   Exit Function
'
search_changed_Error:
    ErrorRaise "frm_search_details.search_changed"
End Function

Private Sub get_dyna_search()
'***************************************
' Module/Form Name   : frm_search_details
'
' Procedure Name     : get_dyna_search
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
On Error GoTo get_dyna_search_Error
'
'******** Code Starts Here *************
'
    SQL = "select search_description, search_no, PhotoType, job_no from search " _
    & "where search_no = " + Str(snap_search!search_no) + ";"

    Set dyna_search = db.OpenRecordset(SQL, dbOpenDynaset)
'
'********* Code Ends Here **************
'
   Exit Sub
'
get_dyna_search_Error:
    ErrorRaise "frm_search_details.get_dyna_search"
End Sub


Private Sub get_snap_search()
'***************************************
' Module/Form Name   : frm_search_details
'
' Procedure Name     : get_snap_search
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
On Error GoTo get_snap_search_Error
'
'******** Code Starts Here *************
'
    SQL = "select search_description, search_no, PhotoType from search " _
    & "where search_no = " + Str(mod_search_no) + ";"

    Set snap_search = db.OpenRecordset(SQL, dbOpenSnapshot)
'
'********* Code Ends Here **************
'
   Exit Sub
'
get_snap_search_Error:
    ErrorRaise "frm_search_details.get_snap_search"
End Sub

Private Sub update_search()
'***************************************
' Module/Form Name   : frm_search_details
'
' Procedure Name     : update_search
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
On Error GoTo update_search_Error
'
'******** Code Starts Here *************
'
    retry_update = True

    begin_trans
        get_dyna_search
        dyna_search.Edit
        If search_changed Then
            MsgBox "Search has been changed by another user." + vbCr + _
            "Press OK to see the current details", vbExclamation
            roll_back
''            load_search_title snap_search!search_no, snap_search!job_no, mod_action
            load_search_title moJob, mod_action, moSearch
            Exit Sub
        End If
'
'   Update the search record.
'
        With dyna_search
            !Search_Description = txt_search_description.Text
''            !PhotoType = IIf(cboPhotoType.ListIndex = 0, "S", "D")
            .Update
            moSearch.Description = txt_search_description.Text
        End With

        mod_search_no = dyna_search!search_no
        dyna_search.Close
        retry_update = False

    commit_trans
'
'********* Code Ends Here **************
'
    Exit Sub
'
update_search_Error:
    Select Case Err.Number
        Case Is = 3022
            MsgBox "Search '" + txt_search_description.Text + "' already exists for Customer: ", vbExclamation
            close_search_recordsets
            On Error Resume Next
            roll_back
        Case Else
            ErrorRaise "frm_search_details.update_search"
    End Select
End Sub

Public Sub load_search_title(ByRef poJob As Job2, _
                             Action As UpdateMode, _
                             ByRef poSearch As Search2)
'***************************************
' Module/Form Name   : frm_search_details
'
' Procedure Name     : load_search_title
'
' Purpose            :
'
' Date Created       : 18/12/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : search_no - Integer
'                    : job_no - As
'                    : UpdateMode -
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo load_search_title_Error
'
'******** Code Starts Here *************
'
    If poSearch Is Nothing Then
        mod_search_no = 0
        Set moSearch = New Search2
    Else
        mod_search_no = poSearch.SearchNo
        Set moSearch = poSearch
    End If
    mod_job_no = poJob.JobNo
    mod_action = Action
    Set moJob = poJob
    '
    If mod_action = Edit Then
        get_snap_search
        If snap_search.EOF Then
            '
            '   Search has been deleted by another user or window.
            '
            Err.Raise vbObjectError + 2, , "Search has been deleted."
        Else
            txt_search_description = snap_search!Search_Description
        End If
    End If

    Me.Show 1
'
'********* Code Ends Here **************
'
   Exit Sub
'
load_search_title_Error:
    ErrorRaise "frm_search_details.load_search_title"
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    Set frm_search_details = Nothing
End Sub

Private Sub cmdOK_Click()
'***************************************
' Module/Form Name   : frm_search_details
'
' Procedure Name     : cmdOK_Click
'
' Purpose            :
'
' Date Created       : 27/11/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 30/11/2002 GARETH SAUNDERS
'                    : 18/12/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo cmdOK_Click_Error
    '
    '******** Code Starts Here *************
    '
    Dim fSearch As frm_search

    If Not valid_input Then
        Exit Sub
    End If
    If mod_action = Add Then
        create_search
        '
        '   Refresh the Job's Searches recordset.
        '
        moJob.Searches.Refresh
        '
        If Not retry_update Then
            RaiseEvent TitleCreated(mod_search_no)
            Set frm_search_details = Nothing
        Else
            txt_search_description.SetFocus
            Exit Sub
        End If
    Else
        update_search
        If Not retry_update Then
            RaiseEvent TitleChanged
            Set frm_search_details = Nothing
        Else
            txt_search_description.SetFocus
            Exit Sub
        End If
    End If
    '
    '   Update Screens.
    '
    update_loaded_screens
    '
    '   Invoke the Search Maintenance Screen if adding.
    '
    If mod_action = Add Then
        Set fSearch = New frm_search
        Unload Me
        DoEvents
        fSearch.display_search_form Edit, moJob, moSearch.SearchNo
    End If
    '
    Unload Me
    DoEvents
    '
    '********* Code Ends Here **************
    '
    Exit Sub
    '
cmdOK_Click_Error:
    DisplayError , "frm_search_details.cmdOk_Click", vbExclamation
End Sub

Private Sub update_loaded_screens()
    LockWindow Me.hwnd
    '
    If is_form_loaded("frm_job_maint") Then
        frm_job_maint.PendingRedisplay
    End If
    '
    If is_form_loaded("frm_job_edit") Then
''        frm_job_edit.display_job moJob.CustomerNo, moJob.JobNo, moJob, mod_search_no
        frm_job_edit.PendingRedisplay
    End If
    '
    Me.SetFocus
    UnlockWindow
End Sub


