VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm_delivery_note_return_photos 
   Caption         =   "Delivery Note - Return Photographs"
   ClientHeight    =   6060
   ClientLeft      =   2115
   ClientTop       =   1860
   ClientWidth     =   8715
   Icon            =   "Delivery Note Return Photographs.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6060
   ScaleWidth      =   8715
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   675
      Left            =   6720
      Picture         =   "Delivery Note Return Photographs.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1800
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "&Return/Reject"
      Default         =   -1  'True
      Height          =   675
      Left            =   5280
      Picture         =   "Delivery Note Return Photographs.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1800
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.Frame fraDeliveryNote 
      Caption         =   "Delivery Note"
      ForeColor       =   &H80000001&
      Height          =   1635
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   7335
      Begin VB.TextBox txt_customer_name 
         BackColor       =   &H8000000F&
         ForeColor       =   &H80000001&
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   5715
      End
      Begin VB.TextBox txt_address_line_1 
         BackColor       =   &H8000000F&
         ForeColor       =   &H80000001&
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   660
         Width           =   5715
      End
      Begin VB.TextBox txt_requestor 
         BackColor       =   &H8000000F&
         ForeColor       =   &H80000001&
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   960
         Width           =   1995
      End
      Begin VB.TextBox txt_return_date 
         BackColor       =   &H8000000F&
         ForeColor       =   &H80000001&
         Height          =   285
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   960
         Width           =   1995
      End
      Begin VB.TextBox txt_job_reference 
         BackColor       =   &H8000000F&
         ForeColor       =   &H80000001&
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1260
         Width           =   5715
      End
      Begin VB.Label lab_customer_name_lab 
         Caption         =   "Customer Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lab_address_line_1_lab 
         Caption         =   "Address Line 1:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   660
         Width           =   1155
      End
      Begin VB.Label lab_requestor_lab 
         Caption         =   "Requestor:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1155
      End
      Begin VB.Label lab_job_reference_lab 
         Caption         =   "Job Reference:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1260
         Width           =   1275
      End
      Begin VB.Label lab_return_date_lab 
         Caption         =   "Return Date:"
         Height          =   195
         Left            =   3960
         TabIndex        =   7
         Top             =   960
         Width           =   1035
      End
   End
   Begin VB.TextBox txt_image_no 
      Height          =   315
      Left            =   4140
      MaxLength       =   8
      TabIndex        =   18
      Top             =   2160
      Width           =   795
   End
   Begin VB.TextBox txt_booked_in_by 
      Height          =   315
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   17
      Top             =   2160
      Width           =   2055
   End
   Begin MSComctlLib.StatusBar stb_return 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   20
      Top             =   5805
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4921
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4921
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4921
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lsv_delivery_note 
      Height          =   2895
      Left            =   0
      TabIndex        =   19
      Top             =   2580
      WhatsThisHelpID =   1
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   5106
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Image"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Description"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Returned by"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Return Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Digital Status"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComCtl2.DTPicker txt_date_returned 
      Height          =   285
      Left            =   1440
      TabIndex        =   12
      Top             =   1800
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   63700995
      CurrentDate     =   36844
   End
   Begin VB.Label lab_image_no 
      Caption         =   "Image no:"
      Height          =   195
      Left            =   4200
      TabIndex        =   13
      Top             =   1860
      Width           =   735
   End
   Begin VB.Label Lab_date_returned_lab 
      Caption         =   "Date:"
      Height          =   195
      Left            =   60
      TabIndex        =   11
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label lab_booked_in_by_lab 
      Caption         =   "Booked in by:"
      Height          =   255
      Left            =   60
      TabIndex        =   14
      Top             =   2160
      Width           =   1155
   End
End
Attribute VB_Name = "frm_delivery_note_return_photos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private snap_job                    As DAO.Recordset
Private snap_delivery_note          As DAO.Recordset
Private snap_customer               As DAO.Recordset
Private dyna_search_result          As DAO.Recordset
Private snap_search_result          As DAO.Recordset
Private abort                       As Boolean
Private list_view_selected          As Boolean
Private mod_customer_filter         As Integer
Private mod_photograph_filter       As String
Private mod_delivery_note_no        As Long
Private mblnOpenDeliveryNotesOnly   As Boolean
Private mod_search_result_key       As String
Private mlngCurrentDeliveryNoteNo   As Long

Private Sub display_images(ByVal delivery_note_no As Long, _
                           ByRef abort As Boolean, _
                           Optional ByVal plngCurrentSearchResultKey As Long = 0)
'***************************************
' Module/Form Name   : frm_delivery_note_return_photos
'
' Procedure Name     : display_images
'
' Purpose            :
'
' Date Created       : 07/12/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : delivery_note_no -
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo display_images_Error
'
'******** Code Starts Here *************
'
    Dim intIndex        As Integer
    
    abort = True
'
'   Get all the photographs for the delivery note.
'
    Set dyna_search_result = search_result.read_all_delivery_note_photos(delivery_note_no, dbOpenDynaset)
    
    lsv_delivery_note.ListItems.Clear
    Dim itmX As ListItem
    With dyna_search_result
        While Not .EOF
            Set itmX = lsv_delivery_note.ListItems.Add _
                                (, , format_search_result(!batch_no, _
                                                          !photograph_no, _
                                                          ""))
            itmX.SubItems(1) = !Description
            If Not IsNull(!booked_in_by) Then
                itmX.SubItems(2) = !booked_in_by
                '
                '   3.8: Don't now need this for Digital images.
                '
''            ElseIf NVL(!photograph_no, NVLNumeric) = 0 Then
''                itmX.SubItems(2) = "N/A"
            Else
                itmX.SubItems(2) = ""
            End If
            If Not IsNull(!date_returned) Then
                If !date_returned = 0 Then
                    '
                    '   3.8: Don't now need this for Digital images.
                    '
''                    If NVL(!photograph_no, NVLNumeric) = 0 Then
''                        itmX.SubItems(3) = "N/A"
''                    Else
                        itmX.SubItems(3) = ""
''                    End If
                Else
                    itmX.SubItems(3) = Format(!date_returned, "dd/mm/yyyy")
                End If
            Else
                itmX.SubItems(3) = ""
            End If
            '
            '   Add column to say whether a digital image has been Invoiced.
            '
            itmX.SubItems(4) = IIf(!DigitalStatus = "A", "Invoiced", IIf(!DigitalStatus = "R", "Rejected", ""))
            itmX.Tag = CStr(!search_result_key)
            .MoveNext
        Wend
    End With
'
'   Sort by Photograph number and date returned.
'
    lsv_delivery_note.SortKey = 0
    lsv_delivery_note.Sorted = True
    
    lsv_delivery_note.SortKey = 2
    lsv_delivery_note.Sorted = True

'
'   Determine currently selected item.
'
    With lsv_delivery_note
        '
        '   Remove all current selections.
        '
        If .ListItems.Count > 0 Then
            .SelectedItem.Selected = False
        End If
        If plngCurrentSearchResultKey <> 0 Then
            For intIndex = 1 To .ListItems.Count
                If .ListItems(intIndex).Tag = CStr(plngCurrentSearchResultKey) Then
                    Set .SelectedItem = .ListItems(intIndex)
                    Exit For
                End If
            Next intIndex
        End If
    End With
    abort = False
'
'********* Code Ends Here **************
'
   Exit Sub
'
display_images_Error:
    ErrorRaise "frm_delivery_note_return_photos.display_images"
End Sub

Public Sub edit_photo()
'***************************************
' Module/Form Name   : frm_delivery_note_return_photos
'
' Procedure Name     : edit_photo
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
On Error GoTo edit_photo_Error
'
'******** Code Starts Here *************
'
'
'   Used for invoking the edit photograph from from search maintenance.
'   Find the search result in the snapshot using the mod_search_result_key set up
'   in the lsv_delivery_note_ItemClick event.
'
    If mod_search_result_key = 0 Then
        MsgBox "Please select an image", vbExclamation, "Customer Returns"
        Exit Sub
    End If
    
    Set snap_search_result = search_result.Read(mod_search_result_key, dbOpenSnapshot)
'
'   Invoke the edit photograph form.
'
    frm_photograph_edit.load_batch LTrim(snap_search_result!batch_no)
    frm_photograph_edit.SetFocus
'
'********* Code Ends Here **************
'
   Exit Sub
'
edit_photo_Error:
    ErrorRaise "frm_delivery_note_return_photos.edit_photo"
End Sub

Public Sub redisplay_images()
'***************************************
' Module/Form Name   : frm_delivery_note_return_photos
'
' Procedure Name     : redisplay_images
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
On Error GoTo redisplay_images_Error
'
'******** Code Starts Here *************
'
    display_images mlngCurrentDeliveryNoteNo, abort, Val(mod_search_result_key)
'
'********* Code Ends Here **************
'
   Exit Sub
'
redisplay_images_Error:
    ErrorRaise "frm_delivery_note_return_photos.redisplay_images"
End Sub

Private Sub txt_image_no_SetFocus()
    On Error Resume Next
    txt_image_no.SetFocus
    txt_image_no.SelStart = 0
    txt_image_no.SelLength = Len(txt_image_no.Text)
End Sub

Private Sub display_totals()
'***************************************
' Module/Form Name   : frm_delivery_note_return_photos
'
' Procedure Name     : display_totals
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
On Error GoTo display_totals_Error
'
'******** Code Starts Here *************
'
    Dim pnlX As Panel
    Dim number_outstanding As Integer
    Dim i As Integer
    
    number_outstanding = 0
    For i = 1 To lsv_delivery_note.ListItems.Count
        '
        '   If Digital Status is set, the image is not outstanding.
        '
        If lsv_delivery_note.ListItems(i).SubItems(4) = "" Then
            If Len(lsv_delivery_note.ListItems(i).SubItems(2)) = 0 Then
                number_outstanding = number_outstanding + 1
            End If
        End If
    Next i
    
    Set pnlX = stb_return.Panels.Item(1)
    pnlX.Text = "Returned: " + Str(lsv_delivery_note.ListItems.Count - number_outstanding)
    
    Set pnlX = stb_return.Panels.Item(2)
    pnlX.Text = "Outstanding: " + Str(number_outstanding)
    
    Set pnlX = stb_return.Panels.Item(3)
    pnlX.Text = "Total: " + Str(lsv_delivery_note.ListItems.Count)
'
'********* Code Ends Here **************
'
   Exit Sub
'
display_totals_Error:
    ErrorRaise "frm_delivery_note_return_photos.display_totals"
End Sub

Private Sub TidyDualSelection()
'***************************************
' Module/Form Name   : frm_delivery_note_return_photos
'
' Procedure Name     : TidyDualSelection
'
' Purpose            :
'
' Date Created       : 01/10/2006 17:08
'
' Author             : Gareth
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
    Dim i As Integer

    On Error GoTo ErrorProc

    If Trim(txt_image_no.Text) <> "" Then
        '
        '   Remove any current selections.
        '
        For i = 1 To lsv_delivery_note.ListItems.Count
            If lsv_delivery_note.ListItems(i).Selected = True Then
                lsv_delivery_note.ListItems(i).Selected = False
            End If
        Next i
    End If

    On Error GoTo 0
    Exit Sub
'
'********* Code Ends Here **************
'
ErrorProc:
    ErrorRaise "frm_delivery_note_return_photos.TidyDualSelection"
End Sub

Private Sub update_loaded_screens()
'***************************************
' Module/Form Name   : frm_delivery_note_return_photos
'
' Procedure Name     : update_loaded_screens
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
On Error GoTo update_loaded_screens_Error
'
'******** Code Starts Here *************
'
    If is_form_loaded("frm_delivery_note_maint") Then
        frm_delivery_note_maint.RefreshDeliveryNotes pblnPicturesReturned:=True
    End If
    
    If is_form_loaded("frm_photograph_view") Then
        frm_photograph_view.datPhotograph.Refresh
        frm_photograph_view.RedisplayDigitalImages
    End If
    
    If is_form_loaded("frm_photograph_edit") Then
        frm_photograph_edit.redisplay_images        'datBatch.Refresh
    End If
'
'********* Code Ends Here **************
'
   Exit Sub
'
update_loaded_screens_Error:
    ErrorRaise "frm_delivery_note_return_photos.update_loaded_screens"
End Sub

Private Function valid_input() As Boolean
'***************************************
' Module/Form Name   : frm_delivery_note_return_photos
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
    Dim i As Integer
    
    valid_input = False
    
    If Not IsDate(txt_date_returned) Then
        MsgBox "Please enter a valid date", vbExclamation
        txt_date_returned.SetFocus
        Exit Function
    End If

    If Len(RTrim(txt_booked_in_by)) = 0 Then
        MsgBox "Please enter your Id", vbExclamation
        txt_booked_in_by.SetFocus
        Exit Function
    End If

    list_view_selected = False
    For i = 1 To lsv_delivery_note.ListItems.Count
        If lsv_delivery_note.ListItems(i).Selected = True Then
            list_view_selected = True
            Exit For
        End If
    Next i
    
    If Not list_view_selected Then
        If Len(RTrim(txt_image_no.Text)) = 0 Then
            MsgBox "Please enter an image number or select a image from the list", vbExclamation
            txt_image_no.SetFocus
            Exit Function
        Else
            '
            '   Removed in 3.8 to allow selection of digital images.
            '
''            If (InStr(txt_image_no, ".") = 0 Or last_character(txt_image_no) = ".") Then
''                MsgBox "Please enter a valid image number, for example: 12345.2", vbExclamation
''                txt_image_no.SetFocus
''                Exit Function
''            End If
        End If
    Else
        If Len(RTrim(txt_image_no.Text)) <> 0 Then
            MsgBox "Please select either an image number or select a image from the list, but not both", vbExclamation, App.Title
            TidyDualSelection
            '
            txt_image_no.SetFocus
            Exit Function
        End If
    End If
    
    valid_input = True
'
'********* Code Ends Here **************
'
   Exit Function
'
valid_input_Error:
    ErrorRaise "frm_delivery_note_return_photos.valid_input"
End Function

Public Sub load_delivery_note(ByVal DeliveryNoteNo As Integer, _
                              ByVal filter_delivery_note_no As Long, _
                              ByVal filter_customer_no As Long, _
                              ByVal filter_photograph_no As Long, _
                              ByVal pblnFilterOpenDeliveryNotesOnly As Boolean)
'***************************************
' Module/Form Name   : frm_delivery_note_return_photos
'
' Procedure Name     : load_delivery_note
'
' Purpose            :
'
' Date Created       : 07/12/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : DeliveryNoteNo - Integer
'                    : filter_delivery_note_no - filter_photograph_no
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo load_delivery_note_Error
'
'******** Code Starts Here *************
'
    
    Load frm_delivery_note_return_photos
    Screen.MousePointer = ccHourglass
    DoEvents
    '
    '   Set the current Delivery Note number.
    '
    mlngCurrentDeliveryNoteNo = DeliveryNoteNo
'
'   Set filters for (re)loading of maintenance screen.
'
    mod_customer_filter = filter_customer_no
    mod_photograph_filter = filter_photograph_no
    mod_delivery_note_no = filter_delivery_note_no
    mblnOpenDeliveryNotesOnly = pblnFilterOpenDeliveryNotesOnly
'
'   Set up Return Delivery note window.
'
    If InStr(1, Me.Caption, ":") = 0 Then
        Me.Caption = Me.Caption + ": " + Str(DeliveryNoteNo)
    Else
        Me.Caption = Mid(Me.Caption, 1, InStr(1, Me.Caption, ":")) + Str(DeliveryNoteNo)
    End If
    
    Set snap_delivery_note = delivery_note.Read(DeliveryNoteNo)
    
    Set snap_customer = Customer.Read(snap_delivery_note!customer_no)
    txt_customer_name = snap_customer!customer_name
    txt_address_line_1 = snap_customer!address_line_1
    
    Set snap_job = job.read_for_delivery_note(snap_delivery_note!delivery_note_no)
    
    txt_job_reference = snap_job!job_reference
    If Not IsNull(snap_job!requested_by) Then
        txt_requestor = snap_job!requested_by
    End If
    txt_return_date = Format(snap_delivery_note!return_date, "dd/mm/yyyy")
    
    txt_date_returned = Format(Date, "dd/mm/yyyy")
'
'   Get all the photographs for the delivery note.
'
    display_images DeliveryNoteNo, abort
'
    Screen.MousePointer = ccDefault
'
    snap_delivery_note.Close
    snap_customer.Close
    snap_job.Close
    dyna_search_result.Close
    
    display_totals
    Me.Show
'
'   Position cursor on Name.
'
    txt_booked_in_by.SetFocus
'
'
'********* Code Ends Here **************
'
   Exit Sub
'
load_delivery_note_Error:
    ErrorRaise "frm_delivery_note_return_photos.load_delivery_note"
End Sub

Private Sub cmdExit_Click()
    '***************************************
    ' Module/Form Name   : frm_delivery_note_return_photos
    '
    ' Procedure Name     : cmdExit_Click
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
On Error GoTo cmdExit_Click_Error
    '
    '******** Code Starts Here *************
    '

frm_delivery_note_maint.list_delivery_notes mod_delivery_note_no, _
                                          mod_customer_filter, _
                                          mod_photograph_filter, _
                                          abort, _
                                          True, _
                                          mblnOpenDeliveryNotesOnly

Unload Me
Set frm_delivery_note_return_photos = Nothing

    '
    '********* Code Ends Here **************
    '
Exit Sub
    '
cmdExit_Click_Error:
    DisplayError , "frm_delivery_note_return_photos.cmdExit_Click", vbExclamation
End Sub


Private Sub Form_Load()
    
    Width = 7470
    Height = 6165
    mod_search_result_key = 0
    com_position_form Me
'
'   Make visible the search menu.
'
    mdi_npls.mnuReturn.Visible = True
'
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormMDIForm Then
        mdi_npls.mnuReturn.Visible = False
    End If
End Sub


Private Sub Form_Resize()
'***************************************
' Module/Form Name   : frm_delivery_note_return_photos
'
' Procedure Name     : Form_Resize
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
On Error GoTo Form_Resize_Error
'
'******** Code Starts Here *************
'
    Dim MIN_WIDTH As Integer
    Dim MIN_HEIGHT As Integer
    
    MIN_WIDTH = 7470
    MIN_HEIGHT = 6165
    
    If Not ResizeForm(Me) Then
        Exit Sub
    End If
    
    If WindowState = 1 Then
        Exit Sub
    End If
  
    If WindowState = 0 Then
        If Width < MIN_WIDTH Then Width = MIN_WIDTH
        If Height < MIN_HEIGHT Then Height = MIN_HEIGHT
    End If
    
    With fraDeliveryNote
        .Width = frm_delivery_note_return_photos.Width - 100
        txt_customer_name.Width = .Width - txt_customer_name.Left - 100
        txt_address_line_1.Width = .Width - txt_address_line_1.Left - 100
        txt_job_reference.Width = .Width - txt_job_reference.Left - 100
    End With
    
    cmdExit.Left = frm_delivery_note_return_photos.Width - cmdExit.Width - 200
    cmdReturn.Left = cmdExit.Left - cmdReturn.Width - 100
    
    With lsv_delivery_note
        .Visible = False
        .Left = 0
        .Width = frm_delivery_note_return_photos.Width - 100
        .Height = frm_delivery_note_return_photos.Height - .Top - stb_return.Height - 500
        .ColumnHeaders(1).Width = 900
        .ColumnHeaders(3).Width = 1200
        .ColumnHeaders(4).Width = 1200
        .ColumnHeaders(2).Width = .Width - .ColumnHeaders(1).Width - _
                                           .ColumnHeaders(3).Width - _
                                           .ColumnHeaders(4).Width - 1500
        .Visible = True
    End With
    
    txt_return_date.Left = fraDeliveryNote.Width - txt_return_date.Width - 100
    lab_return_date_lab.Left = txt_return_date.Left - lab_return_date_lab.Width - 200

'
'********* Code Ends Here **************
'
   Exit Sub
'
Form_Resize_Error:
    DisplayError , "frm_delivery_note_return_photos.Form_Resize", vbExclamation
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    gcolMaxedWindows.Remove CStr(Me.hWnd)
End Sub

Private Sub lsv_delivery_note_ColumnClick(ByVal ColumnHeader As ColumnHeader)
    ' When a ColumnHeader object is clicked, the ListView control is
    ' sorted by the subitems of that column.
    ' Set the SortKey to the Index of the ColumnHeader - 1
    lsv_delivery_note.SortKey = ColumnHeader.Index - 1
    ' Set Sorted to True to sort the list.
    lsv_delivery_note.Sorted = True
End Sub

Private Sub lsv_delivery_note_DblClick()
'***************************************
' Module/Form Name   : frm_delivery_note_return_photos
'
' Procedure Name     : lsv_delivery_note_DblClick
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
On Error GoTo lsv_delivery_note_DblClick_Error
'
'******** Code Starts Here *************
'
    If mod_search_result_key <> 0 Then
        edit_photo
    End If
'
'********* Code Ends Here **************
'
   Exit Sub
'
lsv_delivery_note_DblClick_Error:
    DisplayError , "frm_delivery_note_return_photos.lsv_delivery_note_DblClick", vbExclamation
End Sub

Private Sub lsv_delivery_note_ItemClick(ByVal Item As ListItem)
    Dim i As Integer
    i = 1
    Do
        If i > lsv_delivery_note.ListItems.Count Then
            Exit Sub
        End If
        If lsv_delivery_note.ListItems(i).Selected Then
            Exit Do
        End If
        i = i + 1
    Loop
'
'   Set the key ready for any manipulation necessary.
'
    mod_search_result_key = lsv_delivery_note.ListItems(i).Tag '

End Sub

Private Sub lsv_delivery_note_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'***************************************
' Module/Form Name   : frm_delivery_note_return_photos
'
' Procedure Name     : lsv_delivery_note_MouseUp
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
On Error GoTo lsv_delivery_note_MouseUp_Error
'
'******** Code Starts Here *************
'
    Dim i
'
'   If right mouse button clicked and an item selected, display pop up menu.
'
    If Button <> vbRightButton Then
'       mod_search_result_key <> 0 Then
        Exit Sub
    End If
'
'   Display pop up menu.
'
    PopupMenu mdi_npls.mnuReturn, vbPopupMenuRightButton, , , mdi_npls.mnuReturnEditPhoto
'
'********* Code Ends Here **************
'
   Exit Sub
'
lsv_delivery_note_MouseUp_Error:
    DisplayError , "frm_delivery_note_return_photos.lsv_delivery_note_MouseUp", vbExclamation
End Sub

Private Sub cmdReturn_Click()
'***************************************
' Module/Form Name   : frm_delivery_note_return_photos
'
' Procedure Name     : cmdReturn_Click
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
On Error GoTo cmdReturn_Click_Error
'
'******** Code Starts Here *************
'

    Dim selected_items                      As Integer
    Dim i                                   As Integer
    Dim intImagesReturned                   As Integer
    Dim lsiLastImage                        As ListItem
    Dim blnMultipleDigitalImagesSelected    As Boolean
    Dim intNoDigitalImages                  As Integer
    Dim blnIgnoreDigitalImages              As Boolean
    
    If Not valid_input Then
        Exit Sub
    End If

    Screen.MousePointer = vbHourglass
    begin_trans
'
'   If items selected from the list view then process them, otherwise process
'   that entered into the text box.
'
        If Not list_view_selected Then
'
'           Find the item in the list view using the text box entry.
'
            Dim item_found As Boolean
            For i = 1 To lsv_delivery_note.ListItems.Count
                If characters_before_space(LTrim(lsv_delivery_note.ListItems(i).Text)) = _
                   LTrim(RTrim(txt_image_no.Text)) Then
                    item_found = True
                    mdi_npls.MousePointer = ccDefault
                    Exit For
                End If
            Next i
            '   If no ListItem is found, then inform user and exit.  If a
            '   ListItem is found, scroll the control using the EnsureVisible
            '   method, and select the ListItem.
            If item_found Then
                lsv_delivery_note.ListItems(i).EnsureVisible ' Scroll ListView to show found ListItem.
                lsv_delivery_note.ListItems(i).Selected = True   ' Select the ListItem.
                '   Return focus to the control to see selection.
            Else
                MsgBox "No match found", vbExclamation
                roll_back
                txt_image_no_SetFocus
                mdi_npls.MousePointer = ccDefault
                Exit Sub
            End If
        End If
        '
        '   Determine how many digital images have been selected.
        '
        intNoDigitalImages = 0
        blnMultipleDigitalImagesSelected = False
        blnIgnoreDigitalImages = False
        For i = 1 To lsv_delivery_note.ListItems.Count
            If lsv_delivery_note.ListItems(i).Selected = True Then
                If InStr(lsv_delivery_note.ListItems(i), ".") = 0 Then
                    intNoDigitalImages = intNoDigitalImages + 1
                    If intNoDigitalImages > 1 Then
                        blnMultipleDigitalImagesSelected = True
                        If MsgBox("You have selected a number of digital pictures. Has the client rejected these?", vbQuestion + vbYesNo) = vbNo Then
                            blnIgnoreDigitalImages = True
                        End If
                        Exit For
                    End If
                End If
            End If
        Next i
        '
        selected_items = 0
        intImagesReturned = 0
        For i = 1 To lsv_delivery_note.ListItems.Count
            If lsv_delivery_note.ListItems(i).Selected = True Then
                selected_items = selected_items + 1
                Set lsiLastImage = lsv_delivery_note.ListItems(i)
                If InStr(lsv_delivery_note.ListItems(i), ".") = 0 Then
                    If blnMultipleDigitalImagesSelected Then
                        If Not blnIgnoreDigitalImages Then
                            RejectImage lsv_delivery_note.ListItems(i)
                            intImagesReturned = intImagesReturned + 1
                        End If
                    Else
                        '
                        '   Digital picture.
                        '
                        If RejectImage(lsv_delivery_note.ListItems(i)) Then
                            intImagesReturned = intImagesReturned + 1
                        End If
                    End If
                Else
                    If Not Len(lsv_delivery_note.ListItems(i).SubItems(2)) = 0 Then
                        MsgBox "Photograph: " + Str(lsv_delivery_note.ListItems(i).Text) + _
                               " has already been booked back in! Ignoring request.", vbInformation
                    Else
                        search_result.return_photo CLng(lsv_delivery_note.ListItems(i).Tag), _
                                                   Format(txt_date_returned, "dd/mm/yyyy"), _
                                                   txt_booked_in_by, _
                                                   abort
                        If abort Then
                            roll_back
                            txt_image_no_SetFocus
                            mdi_npls.MousePointer = ccDefault
                            Exit Sub
                        End If
                        intImagesReturned = intImagesReturned + 1
                        '
                        '   Update the List View to show the image has been returned.
                        '
                        lsv_delivery_note.ListItems(i).SubItems(2) = txt_booked_in_by.Text
                        lsv_delivery_note.ListItems(i).SubItems(3) = Format(txt_date_returned, "dd/mm/yyyy")
                    End If
                End If
            End If
        Next i

        If selected_items = 0 Or intImagesReturned = 0 Then
            roll_back
            If selected_items = 0 Then
                MsgBox "No photographs have been selected", vbExclamation
            End If
            txt_image_no_SetFocus
            Screen.MousePointer = ccDefault
            TidyDualSelection
            Exit Sub
        End If
        '
        TidyDualSelection
        '
        '   Sort so that returned images appear at the bottom.
        '
        lsv_delivery_note.SortKey = 0
        lsv_delivery_note.Sorted = True

        lsv_delivery_note.SortKey = 2
        lsv_delivery_note.Sorted = True
        '
        ' Ensure that the last image returned is visible.
        '
        lsiLastImage.EnsureVisible
          
        If abort Then
            roll_back
            txt_image_no_SetFocus
            mdi_npls.MousePointer = ccDefault
            Exit Sub
        End If

    commit_trans
'
'   Update the screen totals.
'
    display_totals
'
'   Remove all selections from the listview.
'
    For i = 1 To lsv_delivery_note.ListItems.Count
        lsv_delivery_note.ListItems.Item(i).Selected = False
    Next i
'
'   Refresh the Return Photographs Screen.
'
    update_loaded_screens

    txt_image_no_SetFocus

    Screen.MousePointer = vbDefault
'
'********* Code Ends Here **************
'
   Exit Sub
'
cmdReturn_Click_Error:
    DisplayError , "frm_delivery_note_return_photos.cmdReturn_Click", vbExclamation
End Sub

Private Function RejectImage(ByRef pImage As MSComctlLib.ListItem) As Boolean
'***************************************
' Module/Form Name   : frm_delivery_note_return_photos
'
' Procedure Name     : RejectImage
'
' Purpose            :
'
' Date Created       : 20/04/2006 13:50
'
' Author             : Gareth
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
    Dim abort           As Boolean
    Dim oSearchResult   As SearchResult
    
    On Error GoTo ErrorProc
    RejectImage = False
    Set oSearchResult = New SearchResult
    oSearchResult.Read CLng(pImage.Tag)
    Select Case oSearchResult.DigitalStatus
        Case Is = "R"
            MsgBox "Digital Image '" & CStr(oSearchResult.BatchNo) & "' has already been rejected. Please select another image.", vbExclamation
            Exit Function
        Case Is = "A"
            MsgBox "Digital Image '" & CStr(oSearchResult.BatchNo) & "' has already been invoiced. Please select another image.", vbExclamation
            Exit Function
        Case Else
            If MsgBox("You have selected a digital image '" & CStr(oSearchResult.BatchNo) & "'. Has the client rejected this?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                Exit Function
            End If
    End Select
    '
    '   Reject image.
    '
    oSearchResult.DigitalStatus = "R"
    oSearchResult.BookedInBy = txt_booked_in_by.Text
    oSearchResult.DateReturned = Now()
    oSearchResult.UpdateDAO
    Set oSearchResult = Nothing
    '
    '   Update the List View to show the image has been returned.
    '
    pImage.SubItems(2) = txt_booked_in_by.Text
    pImage.SubItems(3) = Format(txt_date_returned, "dd/mm/yyyy")
    pImage.SubItems(4) = "Rejected"

    RejectImage = True
    On Error GoTo 0
    Exit Function
'
'********* Code Ends Here **************
'
ErrorProc:
    ErrorRaise "frm_delivery_note_return_photos.RejectImage"
End Function

Private Sub txt_booked_in_by_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_image_no_GotFocus()
    txt_image_no.SelStart = 0
    txt_image_no.SelLength = Len(txt_image_no.Text)
End Sub

Private Sub txt_image_no_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And _
        Not KeyAscii = 8 And _
        Not KeyAscii = 46 Then
        KeyAscii = 0
        Beep
    Else
        If KeyAscii <> vbKeyBack Then
            '
            '   If there is an entry in here then deselect all images.
            '
            DeselectAllImages
        End If
    End If
End Sub

Private Sub txt_image_no_KeyUp(keycode As Integer, Shift As Integer)
   
   Static image_no_save As String
    
    If Len(LTrim(txt_image_no)) = 0 Then
         image_no_save = ""
    Else
        If Not IsNumeric(txt_image_no.Text) Then
            txt_image_no.Text = image_no_save
            txt_image_no.SelStart = Len(txt_image_no.Text)
            Beep
        Else
            image_no_save = txt_image_no.Text
        End If
    End If
End Sub

Private Sub DeselectAllImages()
    Dim itmX As ListItem
    
    For Each itmX In lsv_delivery_note.ListItems
        itmX.Selected = False
    Next itmX

End Sub

Public Sub ForceResize()
    gblnResizeMaxedWindows = False
    Form_Resize
    gblnResizeMaxedWindows = True
End Sub


