VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CFC13920-9EF4-11D0-B72F-0000C04D4C0A}#6.0#0"; "mswless.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_photograph_edit 
   Caption         =   "Edit Photograph"
   ClientHeight    =   7125
   ClientLeft      =   2865
   ClientTop       =   870
   ClientWidth     =   7905
   Icon            =   "Photograph Edit.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7125
   ScaleWidth      =   7905
   Begin VB.PictureBox pic_batch 
      Height          =   5115
      Index           =   0
      Left            =   180
      ScaleHeight     =   5055
      ScaleWidth      =   7455
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   540
      Width           =   7515
      Begin VB.CheckBox chkExclusive 
         Caption         =   "Check1"
         Height          =   195
         Left            =   1920
         TabIndex        =   16
         Top             =   3900
         Width           =   255
      End
      Begin VB.TextBox txtTiffLocation 
         DataField       =   "customername"
         DataSource      =   "dat_customer"
         Height          =   315
         Left            =   1920
         MaxLength       =   50
         TabIndex        =   14
         Top             =   3500
         Width           =   1995
      End
      Begin VB.Frame fraImageFile 
         Caption         =   "Digital Images"
         Height          =   3015
         Left            =   3960
         TabIndex        =   17
         Top             =   1800
         Width           =   3495
         Begin VB.CommandButton cmdScreenRes 
            Height          =   495
            Left            =   2160
            Picture         =   "Photograph Edit.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   1560
            Width           =   495
         End
         Begin VB.CommandButton cmdSetPicture 
            Caption         =   "&Set..."
            Height          =   495
            Left            =   2040
            TabIndex        =   18
            Top             =   240
            Width           =   975
         End
         Begin MSWLess.WLCheck wlcQ9 
            Height          =   255
            Left            =   2160
            TabIndex        =   21
            Top             =   1320
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            _Version        =   393216
            Caption         =   "Q9"
            ForeColor       =   -2147483630
            BackColor       =   -2147483633
            Appearance      =   1
            Enabled         =   0   'False
            Value           =   2
         End
         Begin MSWLess.WLCheck wlcA4 
            Height          =   255
            Left            =   2160
            TabIndex        =   20
            Top             =   1080
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            _Version        =   393216
            Caption         =   "A4"
            ForeColor       =   -2147483630
            BackColor       =   -2147483633
            Appearance      =   1
            Enabled         =   0   'False
            Value           =   2
         End
         Begin MSWLess.WLCheck wlcA3 
            Height          =   255
            Left            =   2160
            TabIndex        =   19
            Top             =   840
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            _Version        =   393216
            Caption         =   "A3"
            ForeColor       =   -2147483630
            BackColor       =   -2147483633
            Appearance      =   1
            Enabled         =   0   'False
            Value           =   2
         End
         Begin VB.Image imgPhotograph 
            Height          =   1740
            Left            =   120
            Stretch         =   -1  'True
            Top             =   240
            Width           =   1860
         End
      End
      Begin VB.ComboBox cmb_month 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1860
         Width           =   735
      End
      Begin VB.ComboBox cmb_year 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   1860
         Width           =   735
      End
      Begin VB.TextBox txt_no_of_images 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   2700
         Width           =   735
      End
      Begin VB.TextBox txt_batch_no 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   0
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   180
         Width           =   735
      End
      Begin VB.ComboBox cmb_format 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   3100
         Width           =   1095
      End
      Begin VB.TextBox txt_batch 
         DataField       =   "customername"
         DataSource      =   "dat_customer"
         Height          =   315
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   8
         Top             =   2300
         Width           =   795
      End
      Begin VB.TextBox txt_description 
         DataField       =   "customername"
         DataSource      =   "dat_customer"
         Height          =   975
         Left            =   1920
         MaxLength       =   180
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   720
         Width           =   4095
      End
      Begin VB.Label lblExclusive 
         Caption         =   "Exclusive:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   3900
         Width           =   975
      End
      Begin VB.Label lblTiffLocation 
         Caption         =   "Tiff Location:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   3500
         Width           =   1455
      End
      Begin VB.Label lab_batch_no_lab 
         Caption         =   "&Photograph Number:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   180
         Width           =   1575
      End
      Begin VB.Label lab_description_lab 
         Caption         =   "&Description:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1155
      End
      Begin VB.Label lab_date_taken_lab 
         Caption         =   "Date &Taken:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1860
         Width           =   1035
      End
      Begin VB.Label lab_batch_lab 
         Caption         =   "&Batch:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   2300
         Width           =   735
      End
      Begin VB.Label lab_no_of_photographs_lab 
         Caption         =   "&No. of images:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   2700
         Width           =   1455
      End
      Begin VB.Label lab_format_lab 
         Caption         =   "Fo&rmat:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   3100
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   555
      Left            =   3120
      Picture         =   "Photograph Edit.frx":0793
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   6120
      UseMaskColor    =   -1  'True
      Width           =   795
   End
   Begin VB.PictureBox pic_batch 
      Height          =   4995
      Index           =   1
      Left            =   240
      ScaleHeight     =   4935
      ScaleWidth      =   7335
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   600
      Width           =   7395
      Begin VB.CommandButton cmdDeleteKeyword 
         Caption         =   "&Delete"
         Height          =   315
         Left            =   6120
         TabIndex        =   48
         Top             =   4080
         Width           =   1095
      End
      Begin VB.CommandButton cmdEditKeyword 
         Caption         =   "&Edit..."
         Height          =   315
         Left            =   4920
         TabIndex        =   47
         Top             =   4080
         Width           =   1095
      End
      Begin VB.CommandButton cmdNewKeyword 
         Caption         =   "&New..."
         Height          =   315
         Left            =   3720
         TabIndex        =   46
         Top             =   4080
         Width           =   1095
      End
      Begin MSFlexGridLib.MSFlexGrid msfKeywords 
         Height          =   1575
         Left            =   1920
         TabIndex        =   45
         Top             =   2400
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   2778
         _Version        =   393216
         FixedCols       =   0
         FocusRect       =   0
         SelectionMode   =   1
      End
      Begin VB.TextBox txt_batch_no 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   1
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   180
         Width           =   735
      End
      Begin VB.TextBox txt_credit 
         DataField       =   "customername"
         DataSource      =   "dat_customer"
         Height          =   675
         Left            =   1920
         MaxLength       =   45
         MultiLine       =   -1  'True
         TabIndex        =   28
         Top             =   1560
         Width           =   5355
      End
      Begin VB.TextBox txt_copyright 
         DataField       =   "customername"
         DataSource      =   "dat_customer"
         Height          =   675
         Left            =   1920
         MaxLength       =   30
         MultiLine       =   -1  'True
         TabIndex        =   26
         Top             =   720
         Width           =   5355
      End
      Begin VB.Label lblKeywords 
         Caption         =   "Keywords:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   60
         TabIndex        =   44
         Top             =   2400
         Width           =   1155
      End
      Begin VB.Label lab_credit_lab 
         Caption         =   "C&redit:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   60
         TabIndex        =   27
         Top             =   1560
         Width           =   1155
      End
      Begin VB.Label lab_copyright_lab 
         Caption         =   "C&opyright:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   60
         TabIndex        =   25
         Top             =   720
         Width           =   1155
      End
      Begin VB.Label lab_batch_no_lab 
         Caption         =   "&Photograph Number:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   23
         Top             =   180
         Width           =   1215
      End
   End
   Begin VB.Data datBatch 
      Caption         =   "datBatch"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6600
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   555
      Left            =   4620
      Picture         =   "Photograph Edit.frx":0D0D
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   6120
      UseMaskColor    =   -1  'True
      Width           =   795
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   555
      Left            =   2160
      Picture         =   "Photograph Edit.frx":129F
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   6120
      UseMaskColor    =   -1  'True
      Width           =   795
   End
   Begin MSComDlg.CommonDialog dlgPhotograph 
      Left            =   6120
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.TabStrip tab_batch 
      Height          =   5115
      Left            =   60
      TabIndex        =   41
      Top             =   60
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   9022
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Credit/Copyright"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Images"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pic_batch 
      Height          =   3495
      Index           =   2
      Left            =   180
      ScaleHeight     =   3435
      ScaleWidth      =   7455
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   600
      Width           =   7515
      Begin VB.CommandButton com_undelete 
         Caption         =   "&Undelete"
         Height          =   315
         Left            =   2580
         TabIndex        =   35
         Top             =   3060
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CheckBox chk_archived_images 
         Caption         =   "Archived Transparencies"
         Height          =   195
         Left            =   5220
         TabIndex        =   31
         Top             =   180
         Width           =   2175
      End
      Begin VB.CommandButton com_job 
         Caption         =   "&Job"
         Height          =   315
         Left            =   6300
         TabIndex        =   37
         Top             =   3060
         Width           =   1095
      End
      Begin VB.CommandButton com_delete 
         Caption         =   "&Delete"
         Height          =   315
         Left            =   1260
         TabIndex        =   34
         Top             =   3060
         Width           =   1095
      End
      Begin VB.CommandButton com_new 
         Caption         =   "&New..."
         Height          =   315
         Left            =   60
         TabIndex        =   33
         Top             =   3060
         Width           =   1095
      End
      Begin VB.CommandButton com_delivery_note 
         Caption         =   "Re&turn"
         Height          =   315
         Left            =   5100
         TabIndex        =   36
         Top             =   3060
         Width           =   1095
      End
      Begin MSFlexGridLib.MSFlexGrid msfImages 
         Bindings        =   "Photograph Edit.frx":1819
         Height          =   2415
         Left            =   0
         TabIndex        =   32
         Top             =   600
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   4260
         _Version        =   393216
         FixedCols       =   0
         WordWrap        =   -1  'True
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         FillStyle       =   1
         ScrollBars      =   2
         SelectionMode   =   1
      End
      Begin VB.TextBox txt_batch_no 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   2
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   180
         Width           =   735
      End
      Begin VB.Label lab_batch_no_lab 
         Caption         =   "Photograph Number:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   60
         TabIndex        =   29
         Top             =   180
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frm_photograph_edit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private snap_batch                  As DAO.Recordset
Private snap_parent_batch           As DAO.Recordset
Private dyna_batch                  As DAO.Recordset

Private SQL                         As String
Private retry_update                As Boolean
Private mod_pic_moved(3)            As String
Private current_tab                 As Integer

Public mod_images_displayed         As Boolean
Private current_image_no(2)         As Integer
Private image_selected(2)           As Boolean
Private mbolDontResize              As Boolean
Private moDigitalImage              As DigitalImage
Private oFSO                        As New Scripting.FileSystemObject
Private moBatch                     As Batch2
Private mintKeywordSortCol          As Integer
Private mintKeywordSortOrder        As Integer
Private mstrCurrentKeyword          As String
Private mintCurrentKeywordCol       As Integer
Private mblnManualKeywordsUpdated   As Boolean
Private Const CONST_DIGITAL As String = "DIGITAL"

Private Sub DeleteImages(abort As Boolean)
'***************************************
' Module/Form Name   : frm_photograph_edit
'
' Procedure Name     : DeleteImages
'
' Purpose            :
'
' Date Created       : 23/04/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : abort - Boolean
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo DeleteImages_Error
'
'******** Code Starts Here *************
'
    Dim iCount As Integer
    Dim CallAbort As Boolean
    
    abort = True
    msg_title = "Delete Image"
    
    With msfImages
        If Abs(.RowSel - .row) < 1 Then
            On Error Resume Next
            photo_image.Delete snap_batch!batch_no, .TextArray(FlexGridIndex(msfImages, .row, 0)), CallAbort
            If Err.Number = vbObjectError + 1 Then  '   Job is Pending!
                MsgBox Err.Description, vbExclamation, App.Title
                Screen.MousePointer = vbDefault
                Exit Sub
            ElseIf Err.Number <> 0 Then
                ErrorSave
                On Error GoTo DeleteImages_Error
                ErrorRestore
            Else
                On Error GoTo DeleteImages_Error
            End If
            If CallAbort Then
                Exit Sub
            End If
        Else
            msg = "Are you certain you wish to delete these Images?"
            style = vbYesNo + vbQuestion + vbDefaultButton2
  
            response = MsgBox(msg, style, msg_title)
  
            If response = vbNo Then
                Exit Sub
            End If

            For iCount = 0 To Abs(.RowSel - .row)
                photo_image.Delete snap_batch!batch_no, .TextArray(FlexGridIndex(msfImages, IIf(.row < .RowSel, .row, .RowSel) + iCount, 0)), CallAbort, True
                If CallAbort Then
                    Exit Sub
                End If
                '
                '   At least one image has been deleted, therefore don't
                '   record an abort.
                '
                abort = False
            Next
 
            MsgBox "Images Deleted", vbInformation, msg_title
        End If
    End With
    
    abort = False
'
'********* Code Ends Here **************
'
   Exit Sub
'
DeleteImages_Error:
    DisplayError , "frm_photograph_edit.DeleteImages", vbExclamation
End Sub
Private Sub UnDeleteImages(abort As Boolean)
'***************************************
' Module/Form Name   : frm_photograph_edit
'
' Procedure Name     : UnDeleteImages
'
' Purpose            :
'
' Date Created       : 07/12/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : abort - Boolean
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo UnDeleteImages_Error
'
'******** Code Starts Here *************
'
    Dim iCount As Integer
    Dim CallAbort As Boolean
    
    abort = True
    msg_title = "Undelete Image"
    
    With msfImages
        If Abs(.RowSel - .row) < 1 Then
            photo_image.undelete snap_batch!batch_no, .TextArray(FlexGridIndex(msfImages, .row, 0)), CallAbort
            If CallAbort Then
                Exit Sub
            End If
        Else
            msg = "Are you certain you wish to undelete these Images?"
            style = vbYesNo + vbQuestion + vbDefaultButton2
  
            response = MsgBox(msg, style, msg_title)
  
            If response = vbNo Then
                Exit Sub
            End If

            For iCount = 0 To Abs(.RowSel - .row)
                photo_image.undelete snap_batch!batch_no, .TextArray(FlexGridIndex(msfImages, IIf(.row < .RowSel, .row, .RowSel) + iCount, 0)), CallAbort, True
                If CallAbort Then
                    Exit Sub
                End If
                '
                '   At least one image has been undeleted, therefore don't
                '   record an abort.
                '
                abort = False
            Next
 
            MsgBox "Images Undeleted", vbInformation, msg_title
        End If
    End With
    
    abort = False
'
'********* Code Ends Here **************
'
   Exit Sub
'
UnDeleteImages_Error:
    ErrorRaise "frm_photograph_edit.UnDeleteImages"
End Sub
Private Sub update_loaded_screens()
'***************************************
' Module/Form Name   : frm_photograph_edit
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
    Dim fSearch As frm_search
    Dim colFrmSearch As New Collection
    Dim fJobEdit As frm_job_edit
    Dim colFrmJobEdit As New Collection
    
    If is_form_loaded("frm_photograph_find_maint") Then
        frm_photograph_find_maint.refresh_list snap_batch!batch_no
    End If
    If is_form_loaded("frm_photograph_view") Then
        frm_photograph_view.RedisplayFields
    End If
    
    If is_form_loaded("frm_delivery_note_maint") Then
        frm_delivery_note_maint.RefreshDeliveryNotes
    End If
    
    If is_form_loaded("frm_photograph_view") Then
        frm_photograph_view.datPhotograph.Refresh
    End If
    
    If is_form_loaded("frm_delivery_note_return_photos") Then
        frm_delivery_note_return_photos.redisplay_images
    End If
    
    If is_form_loaded("frm_search", Me, colFrmSearch) Then
        For Each fSearch In colFrmSearch
            fSearch.redisplay_search_details
        Next fSearch
    End If
    
    If is_form_loaded("frm_job_maint") Then
        frm_job_maint.PendingRedisplay
    End If
    '
    If is_form_loaded("frm_job_edit", Me, colFrmJobEdit) Then
        For Each fJobEdit In colFrmJobEdit
            fJobEdit.PendingRedisplay
        Next fJobEdit
    End If
    
    frm_photograph_edit.SetFocus

'
'********* Code Ends Here **************
'
   Exit Sub
'
update_loaded_screens_Error:
    ErrorRaise "frm_photograph_edit.update_loaded_screens"
End Sub

Public Sub redisplay_images()
'***************************************
' Module/Form Name   : frm_photograph_edit
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
    Dim iCount As Integer
'
'   First refresh the data.
'
    photo_image.Image_Display snap_batch!batch_no, _
                              SQL, _
                              chk_archived_images.Value
    datBatch.RecordSource = SQL
    datBatch.Refresh
    datBatch.UpdateControls
    Me.Refresh
    If datBatch.Recordset.RecordCount = 0 Then
        '
        '   If no images to show then choose a column that isn't shown and select just that.
        '
        msfImages.RowSel = 0
        msfImages.col = 6
        msfImages.ColSel = 6
        '
        '   And then disable the undelete button.
        '
        com_undelete.Enabled = False
        Exit Sub
    Else
        com_undelete.Enabled = True
    End If
'
'   Then find the current image in the Flex Grid.
'
    If chk_archived_images Then
        If current_image_no(2) <> 0 Then
            If current_image_no(2) > datBatch.Recordset.RecordCount Then
                current_image_no(2) = datBatch.Recordset.RecordCount
            End If
            FlexGridHighlightRow msfImages, current_image_no(2)
        End If
    Else
        If current_image_no(1) <> 0 Then
            If current_image_no(1) > datBatch.Recordset.RecordCount Then
                current_image_no(1) = datBatch.Recordset.RecordCount
            End If
            FlexGridHighlightRow msfImages, current_image_no(1)
        Else
            If msfImages.Rows > 1 Then
                '
                '   No image selected before but an image now exists so highlight first row.
                '
                FlexGridHighlightRow msfImages, 1
            Else
                '
                '   Disable the RETURN & JOB buttons as no image is currently selected.
                '
                com_delivery_note.Enabled = False
                com_job.Enabled = False
            End If
        End If
    End If
    
    SetDerivedImageFields
    SetImageManipulationButtons
'
'********* Code Ends Here **************
'
   Exit Sub
'
redisplay_images_Error:
    ErrorRaise "frm_photograph_edit.redisplay_images"
End Sub

Private Sub tab_select(selected_item As Integer)
'***************************************
' Module/Form Name   : frm_photograph_edit
'
' Procedure Name     : tab_select
'
' Purpose            :
'
' Date Created       : 07/12/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : selected_item - Integer
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo tab_select_Error
'
'******** Code Starts Here *************
'
    Dim oCurrentControl As Control
    Dim tab_frame As PictureBox
'
'   Used for positioning pictureboxes on the
'   Batch screens
'
    If mod_pic_moved(selected_item) = "N" Then
        With pic_batch(selected_item - 1)
            .Height = tab_batch.ClientHeight
            .Width = tab_batch.ClientWidth
            .Left = tab_batch.ClientLeft
            .Top = tab_batch.ClientTop
            .Visible = True
            mod_pic_moved(selected_item) = "Y"
        End With
    End If
'
'   Resize tabs.
'
    Select Case selected_item
        Case Is = 1
            txt_description.Width = pic_batch(0).Width - txt_description.Left - 200
        Case Is = 2
            txt_copyright.Width = pic_batch(1).Width - txt_copyright.Left - 200
            txt_credit.Width = pic_batch(1).Width - txt_credit.Left - 200
            '
            '   Resposition Keyword buttons.
            '
            cmdDeleteKeyword.Move pic_batch(1).Width - 200 - cmdDeleteKeyword.Width, pic_batch(1).Height - cmdDeleteKeyword.Height - 200
            cmdEditKeyword.Move cmdDeleteKeyword.Left - 200 - cmdEditKeyword.Width, cmdDeleteKeyword.Top
            cmdNewKeyword.Move cmdEditKeyword.Left - 200 - cmdNewKeyword.Width, cmdDeleteKeyword.Top
            '
            '   Resize Keywords Grid
            '
            With msfKeywords
                .Width = txt_credit.Width
                .Height = cmdEditKeyword.Top - .Top - 100
                .ColWidth(1) = 1000
                .ColWidth(0) = .Width - 364 - .ColWidth(1)
            End With
        Case Is = 3
'
'   Reposition Image manipulation buttons.
'
            com_new.Top = tab_batch.ClientHeight - com_new.Height - 100
            com_delete.Top = com_new.Top
            com_delivery_note.Top = com_new.Top
            com_delivery_note.Left = tab_batch.ClientWidth - (com_delivery_note.Width + _
                                                              com_job.Width + _
                                                              200)
            com_job.Top = com_new.Top
            com_job.Left = com_delivery_note.Left + com_delivery_note.Width + 100
            com_undelete.Top = com_new.Top
            com_undelete.Left = com_new.Left
'
'   Resize database grid.
'
            With msfImages
                .Width = pic_batch(2).Width - .Left - 150
                .Height = Int((com_new.Top - .Top) / .RowHeight(0)) * .RowHeight(0)
                .ColWidth(0) = 720
                .ColWidth(1) = 720
                .ColWidth(2) = 900
                .ColWidth(5) = 850
                .ColWidth(3) = (.Width - 350 - _
                                              .ColWidth(0) - _
                                              .ColWidth(1) - _
                                              .ColWidth(2) - _
                                              .ColWidth(5)) / 2
                .ColWidth(4) = .ColWidth(3)
            End With
    End Select
'
'   Resize the Digital Image.
'
    fraImageFile.Visible = False
'
'   First resize the frame.
'
    fraImageFile.Height = pic_batch(0).Height - fraImageFile.Top - 200
    fraImageFile.Width = pic_batch(0).Width - fraImageFile.Left - 200
'
'   Then move the buttons.
'
    cmdSetPicture.Top = 200
    cmdSetPicture.Left = fraImageFile.Width - 200 - cmdSetPicture.Width
    'cmdRemovePicture.Top = cmdSetPicture.Top + cmdSetPicture.Height + 100
    'cmdRemovePicture.Left = cmdSetPicture.Left
    wlcA3.Move cmdSetPicture.Left, cmdSetPicture.Top + cmdSetPicture.Height + 100
    wlcA4.Move wlcA3.Left, wlcA3.Top + wlcA3.Height + 100
    wlcQ9.Move wlcA4.Left, wlcA4.Top + wlcA4.Height + 100
    cmdScreenRes.Move wlcQ9.Left, wlcQ9.Top + wlcQ9.Height + 100
'
'   Then resize the photograph.
'
     SizePhotograph
'
    fraImageFile.Visible = True
    '
    '   Store the active control away.
    '
    Set oCurrentControl = Me.ActiveControl
    '
    For Each tab_frame In pic_batch
        tab_frame.Enabled = False
        tab_frame.Visible = False
    Next
    pic_batch(selected_item - 1).ZOrder 0
    pic_batch(selected_item - 1).Enabled = True
    pic_batch(selected_item - 1).Visible = True
    '
    '   Put the focus back to the control that had it prior to resizing.
    '
    On Error Resume Next
    oCurrentControl.SetFocus
    If Err.Number = 5 Then
        If selected_item = 1 Then
            txt_description.SetFocus
        ElseIf selected_item = 2 Then
            txt_copyright.SetFocus
        Else
            msfImages.SetFocus
        End If
    End If
'
'********* Code Ends Here **************
'
   Exit Sub
'
tab_select_Error:
    ErrorRaise "frm_photograph_edit.tab_select"
End Sub

Private Sub close_batch_recordset()
    On Error GoTo error_proc
    
    dyna_batch.Close
    
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

Public Sub display_images()
'***************************************
' Module/Form Name   : frm_photograph_edit
'
' Procedure Name     : display_images
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
On Error GoTo display_images_Error
'
'******** Code Starts Here *************
'
    Dim iCount As Integer
'
'   Set up the images SQL.
'
    photo_image.Image_Display snap_batch!batch_no, _
                              SQL, _
                              chk_archived_images.Value
    
    With datBatch
        .DatabaseName = glo_dbname
        .RecordSource = SQL
        .Refresh
    End With
    
    With msfImages
        .Cols = 9         ' Set number of columns.
        .ColAlignment(5) = 2
        .RowHeight(0) = 400
        .ColWidth(6) = 0
        .ColWidth(7) = 0
        .ColWidth(8) = 0
        '
        '   Highlight the first row if it exists
        '
        If .Rows > 1 Then
            .col = 0
            .row = 1
            .ColSel = .Cols - 1
            .RowSel = 1
        Else
            .col = 8
            .ColSel = 8
        End If
        '
        '   Set the derived fields.
        '
        SetDerivedImageFields
        '
    End With
'
'********* Code Ends Here **************
'
   Exit Sub
'
display_images_Error:
    ErrorRaise "frm_photograph_edit.display_images"
End Sub

Private Sub SetDerivedImageFields()
    Dim iCount As Integer
    
'***************************************
' Module/Form Name   : frm_photograph_edit
'
' Procedure Name     : DerivedImageFields
'
' Purpose            :
'
' Date Created       : 10/10/2006 22:22
'
' Author             : Gareth
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
    On Error GoTo ErrorProc

    With msfImages
        For iCount = 1 To .Rows - 1
            '
            '   If the Photograph number is 0 then this is a digital image.
            '
            If .TextMatrix(iCount, 0) = 0 Then
                .TextMatrix(iCount, 0) = "DIGITAL"
                If .TextMatrix(iCount, 1) <> "" Then
                    If .TextMatrix(iCount, 2) = "A" Then
                        .TextMatrix(iCount, 2) = "Inv"
                    ElseIf .TextMatrix(iCount, 2) = "" Then
                        .TextMatrix(iCount, 2) = "O"
                    End If
                End If
            End If
            '
            If .TextMatrix(iCount, 5) = True Then
                .TextMatrix(iCount, 5) = "Yes"
            Else
                .TextMatrix(iCount, 5) = "No"
            End If
        Next
    End With

    On Error GoTo 0
    Exit Sub
'
'********* Code Ends Here **************
'
ErrorProc:
    ErrorRaise "frm_photograph_edit.DerivedImageFields"
End Sub

Private Sub get_dyna_batch()
'***************************************
' Module/Form Name   : frm_photograph_edit
'
' Procedure Name     : get_dyna_batch
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
On Error GoTo get_dyna_batch_Error
'
'******** Code Starts Here *************
'

    Set dyna_batch = photograph.Read(snap_batch!batch_no, dbOpenDynaset)

'
'********* Code Ends Here **************
'
   Exit Sub
'
get_dyna_batch_Error:
    ErrorRaise "frm_photograph_edit.get_dyna_batch"
End Sub

Public Sub load_batch(ByVal Batch_Number As Long)
'***************************************
' Module/Form Name   : frm_photograph_edit
'
' Procedure Name     : load_batch
'
' Purpose            :
'
' Date Created       : 07/12/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : Batch_Number -
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo load_batch_Error
'
'******** Code Starts Here *************
'
    Screen.MousePointer = ccHourglass
    mod_images_displayed = False
    
    Set moDigitalImage = Nothing
    Set moDigitalImage = New DigitalImage
    moDigitalImage.BatchNo = Batch_Number
    Set snap_batch = photograph.Read(Batch_Number, dbOpenSnapshot)
    
    If snap_batch.RecordCount = 0 Then
        MsgBox "Photograph not found", vbExclamation, "Edit Photograph"
        DoEvents
        Set frm_photograph_edit = Nothing
        Screen.MousePointer = ccDefault
        Unload Me
        Exit Sub
    End If
    '
    '   Create the Batch object.
    '
    Set moBatch = New Batch2
    moBatch.Read Batch_Number
    moDigitalImage.Description = moBatch.Description
    '
    Me.Caption = "Edit Photograph: " & CStr(Batch_Number)
    Me.Show
    DoEvents
    display_fields
'
'   If already loaded and images are currently being shown, redisplay.
'
    If tab_batch.SelectedItem.Index = 3 Then
        Screen.MousePointer = vbHourglass
        redisplay_images
        Screen.MousePointer = vbDefault
        mod_images_displayed = True
    End If
    Screen.MousePointer = ccDefault
    Me.SetFocus
    Me.Visible = True
'
'********* Code Ends Here **************
'
   Exit Sub
'
load_batch_Error:
''    Me.Visible = True
    ErrorRaise "frm_photograph_edit.load_batch"
End Sub

Private Function batch_changed() As Boolean
    Dim i As Integer
    batch_changed = False
    For i = 0 To snap_batch.Fields.Count - 1
        If Not snap_batch.Fields(i).Value = dyna_batch.Fields(i).Value Then
            batch_changed = True
            Exit For
        End If
    Next i
End Function

Private Sub update_batch()
'***************************************
' Module/Form Name   : frm_photograph_edit
'
' Procedure Name     : update_batch
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
On Error GoTo update_batch_Error
'
'******** Code Starts Here *************
'
    Dim abort As Boolean
    Dim strDefaultFileLocation As String
    Dim fsoPicture As Scripting.FileSystemObject
    Dim strRelativeFileLocation As String
    
    retry_update = False
    begin_trans
        get_dyna_batch
        dyna_batch.Edit
        If batch_changed Then
            MsgBox "batch has been changed by another user." + vbCr + _
            "Press OK to see the current details", vbExclamation
            roll_back
            retry_update = True
            dyna_batch.Close
            Exit Sub
        End If
  
        With dyna_batch
            '
            '   Apply Pending Updates to Keywords.
            '
            moBatch.Keywords.update
            '
            If !Description <> txt_description.Text Or mblnManualKeywordsUpdated Then
                moBatch.Description = txt_description.Text
                Call UpdateKeywords
            End If
            !Description = txt_description
            !date_taken = CDate(cmb_month.Text + " 01 " + cmb_year.Text)
            If Val(txt_batch) = 0 Then
                !parent_batch = snap_batch!batch_no
            Else
                !parent_batch = Val(txt_batch)
            End If
            !Format = cmb_format.Text
            !CopyRight = txt_copyright.Text
            !Credit = txt_credit.Text
            !TiffLocation = Trim(txtTiffLocation.Text)
            !Exclusive = IIf(chkExclusive.Value = vbChecked, True, False)
            '
            '   Process the digital image.
            '
            With moDigitalImage
                On Error Resume Next
                .Associate
                If Err.Number - vbObjectError = 1 Then
                    DisplayError
                    roll_back
                    close_batch_recordset
                    retry_update = True
                    Exit Sub
                ElseIf Err.Number - vbObjectError = 2 Then
                    If MsgBox(Err.Description & vbCrLf _
                            & "Do you wish to overwrite the existing image with the one that you have chosen?" & vbCrLf _
                            & "The existing Image will be moved to the Holding Area and its file name will have the date and time appended.", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                        roll_back
                        close_batch_recordset
                        retry_update = True
                        Exit Sub
                    Else
                        .Associate True
                    End If
                ElseIf Err.Number <> 0 Then
                    On Error GoTo update_batch_Error
                End If
                On Error GoTo update_batch_Error
            End With
            .update
        End With
  
        dyna_batch.Close
    
    commit_trans
   
'
'********* Code Ends Here **************
'
   Exit Sub
'
update_batch_Error:
    ErrorRaise "frm_photograph_edit.update_batch"
End Sub

Private Sub UpdateKeywords()
'***************************************
' Module/Form Name   : frm_photograph_edit
'
' Procedure Name     : UpdateKeywords
'
' Purpose            :
'
' Date Created       : 18/07/2005
'
' Author             : Gareth Saunders
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                      18/07/2005 Gareth Saunders
'
'***************************************
'
    On Error GoTo ErrorHandler:
'
'******** Code Starts Here *************
'
    If MsgBox("The description and/or manual keywords have been updated. Do you wish to update the automatic keywords.", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
        Exit Sub
    End If
    '
    '   Delete existing Automatic keywords
    '
    DoEvents
    moBatch.DeleteAutomaticKeywords
    '
    '   Add keywords from Caption, ignoring excluded words and adding any clustered ones.
    '
    moBatch.CreateAutomaticKeywords
    mblnManualKeywordsUpdated = False
'
'********* Code Ends Here **************
'
    On Error GoTo 0
    Exit Sub
'
ErrorHandler:
    ErrorRaise "frm_photograph_edit.UpdateKeywords"
End Sub

Private Function valid_input() As Boolean
'***************************************
' Module/Form Name   : frm_photograph_edit
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
    valid_input = False
    
    If txt_description = "" Then
        tab_batch.Tabs(1).Selected = True
        MsgBox "Description must be entered", vbExclamation
        txt_description.SetFocus
        Exit Function
    End If

    If Not Len(RTrim(txt_batch)) = 0 Then
        tab_batch.Tabs(1).Selected = True
        Select Case txt_batch
            Case Is = 0
                MsgBox "Please enter a non-zero number for the Batch", vbExclamation
                txt_batch.SetFocus
                Exit Function
            Case Else
                Set snap_parent_batch = photograph.Read(txt_batch, dbOpenSnapshot)
                If snap_parent_batch.RecordCount = 0 Then
                    MsgBox "The Photograph to be used for grouping the Batch does not exist", vbExclamation
                    txt_batch.SetFocus
                    Exit Function
                End If
                snap_parent_batch.Close
        End Select
    End If

    valid_input = True
'
'********* Code Ends Here **************
'
   Exit Function
'
valid_input_Error:
    ErrorRaise "frm_photograph_edit.valid_input"
End Function

Public Sub display_fields()
'***************************************
' Module/Form Name   : frm_photograph_edit
'
' Procedure Name     : display_fields
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
On Error GoTo display_fields_Error
'
'******** Code Starts Here *************
'
    Dim i               As Integer
    
    If IsNull(snap_batch!batch_no) Then
        txt_batch_no(0) = ""
        txt_batch_no(1) = ""
        txt_batch_no(2) = ""
    Else
        txt_batch_no(0) = snap_batch!batch_no
        txt_batch_no(1) = snap_batch!batch_no
        txt_batch_no(2) = snap_batch!batch_no
    End If
    
    If IsNull(snap_batch!Description) Then
        txt_description = ""
    Else
        txt_description = snap_batch!Description
    End If

    If IsNull(snap_batch!date_taken) Then
        photograph.set_up_date_controls Now, cmb_month, cmb_year
    Else
        photograph.set_up_date_controls snap_batch!date_taken, cmb_month, cmb_year
    End If
    If IsNull(snap_batch!parent_batch) Or _
       snap_batch!parent_batch = 0 Or _
       snap_batch!parent_batch = snap_batch!batch_no Then
        txt_batch = ""
    Else
        txt_batch = snap_batch!parent_batch
    End If
    txt_no_of_images = photograph.count_images(snap_batch!batch_no)
'
'   Set up formats in the combo box.
'
    cmb_format.Clear
  
    cmb_format.AddItem "Digital"
    cmb_format.AddItem "35"
    cmb_format.AddItem "54"
    cmb_format.AddItem "66"
    cmb_format.AddItem "67"
    cmb_format.AddItem "617"
    cmb_format.AddItem "645"
    
    For i = 0 To cmb_format.ListCount - 1
        If Mid(snap_batch!Format + Space(7), 1, 7) = Mid(cmb_format.List(i) + Space(7), 1, 7) Then
            cmb_format.ListIndex = i
        End If
    Next i
    
    txt_copyright = RTrim(snap_batch!CopyRight)
    txt_credit = RTrim(snap_batch!Credit)
    txtTiffLocation = Trim(NVL(snap_batch!TiffLocation, NVLString))
    chkExclusive.Value = IIf(moBatch.Exclusive, vbChecked, vbUnchecked)
    '
    '   Set up ImageFile and store location in memory.
    '
    'DoEvents
    'cmdRemovePicture.Enabled = True
    imgPhotograph.Picture = LoadPicture(moDigitalImage.FileLocation("SC"))
    SizePhotograph
    '
    wlcA3.Value = IIf(moDigitalImage.FileLocation("A3") = "", False, True)
    wlcA4.Value = IIf(moDigitalImage.FileLocation("A4") = "", False, True)
    wlcQ9.Value = IIf(moDigitalImage.FileLocation("Q9") = "", False, True)
    cmdScreenRes.Visible = IIf(moDigitalImage.FileLocation("SC") = "", False, True)
    '
    '   Load the keywords.
    '
    DisplayKeywords
'
'********* Code Ends Here **************
'
   Exit Sub
'
display_fields_Error:
    ErrorRaise "frm_photograph_edit.display_fields"
End Sub

Private Sub DisplayKeywords(Optional pstrCurrentKeyword As String = "")
    Dim i               As Integer
    Dim j               As Integer
    Dim strKeyword      As String
    
    With msfKeywords
        .Clear
        .Rows = 1
        .TextMatrix(0, 0) = "Keyword"
        .TextMatrix(0, 1) = "Type"
    End With
    j = 0
    For i = 1 To moBatch.Keywords.Count
        If moBatch.Keywords.Item(i).Dirty <> pdgDelete Then
            j = j + 1
            strKeyword = moBatch.Keywords.Item(i).Keyword
            msfKeywords.Rows = j + 1
            msfKeywords.TextMatrix(j, 0) = moBatch.Keywords.Item(i).Keyword
            msfKeywords.TextMatrix(j, 1) = moBatch.Keywords.Item(i).KeywordType
            If pstrCurrentKeyword = strKeyword Then
                msfKeywords.row = j
                msfKeywords.col = 0
                msfKeywords.RowSel = j
                msfKeywords.ColSel = 1
            End If
        End If
    Next i
    '
    If pstrCurrentKeyword = "" Then
        mstrCurrentKeyword = msfKeywords.TextMatrix(msfKeywords.Rows - 1, 0)
    Else
        mstrCurrentKeyword = pstrCurrentKeyword
    End If

End Sub

Private Sub chk_archived_images_Click()
'***************************************
' Module/Form Name   : frm_photograph_edit
'
' Procedure Name     : chk_archived_images_Click
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
On Error GoTo chk_archived_images_Click_Error
'
'******** Code Starts Here *************
'
'
'   Hide/unhide manipulation buttons.
'
    If chk_archived_images Then
        com_new.Visible = False
        com_delete.Visible = False
        com_delivery_note.Visible = False
        com_job.Visible = False
        com_undelete.Visible = True
    Else
        com_new.Visible = True
        com_delete.Visible = True
        com_delivery_note.Visible = True
        com_job.Visible = True
        com_undelete.Visible = False
    End If
'''
'''   Update SQL and refresh.
'''
''    photo_image.Image_Display snap_batch!batch_no, _
''                                             SQL, _
''                                             chk_archived_images.Value
    
    datBatch.RecordSource = SQL
    
    redisplay_images

'
'********* Code Ends Here **************
'
   Exit Sub
'
chk_archived_images_Click_Error:
    DisplayError , "frm_photograph_edit.chk_archived_images_Click", vbExclamation
End Sub

Private Sub cmdApply_Click()
'***************************************
' Module/Form Name   : frm_photograph_edit
'
' Procedure Name     : cmdApply_Click
'
' Purpose            :
'
' Date Created       : 19/07/2005
'
' Author             : Gareth Saunders
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                      19/07/2005 Gareth Saunders
'
'***************************************
'
    On Error GoTo ErrorHandler:
'
'******** Code Starts Here *************
'
    If Not valid_input Then
        Exit Sub
    End If
    
    update_batch
    load_batch moBatch.BatchNo
    DoEvents
    MsgBox "All changes have been applied", vbInformation
'
'********* Code Ends Here **************
'
    On Error GoTo 0
    Exit Sub
'
ErrorHandler:
    DisplayError , "frm_photograph_edit.cmdApply_Click", vbExclamation
End Sub

Private Sub cmdDeleteKeyword_Click()
'***************************************
' Module/Form Name   : frm_photograph_edit
'
' Procedure Name     : cmdDeleteKeyword_Click
'
' Purpose            :
'
' Date Created       : 19/07/2005
'
' Author             : Gareth Saunders
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                      19/07/2005 Gareth Saunders
'
'***************************************
'
    On Error GoTo ErrorHandler:
'
'******** Code Starts Here *************
'
    DeleteKeyword
'
'********* Code Ends Here **************
'
    On Error GoTo 0
    Exit Sub
'
ErrorHandler:
    DisplayError , "frm_photograph_edit.cmdDeleteKeyword_Click", vbExclamation
End Sub

Private Sub DeleteKeyword()
    Dim strKeyword       As String
    
'***************************************
' Module/Form Name   : frm_photograph_edit
'
' Procedure Name     : DeleteKeyword
'
' Purpose            :
'
' Date Created       : 19/07/2005
'
' Author             : Gareth Saunders
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                      19/07/2005 Gareth Saunders
'
'***************************************
'
    On Error GoTo ErrorHandler:
'
'******** Code Starts Here *************
'
    strKeyword = msfKeywords.TextMatrix(msfKeywords.row, 0)

    If MsgBox("Are you sure you wish to delete keyword '" & strKeyword & "'", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
        Exit Sub
    End If
    
    moBatch.Keywords.RemovePending strKeyword
    moBatch.Keywords.update
    mblnManualKeywordsUpdated = True
    DisplayKeywords
'
'********* Code Ends Here **************
'
    On Error GoTo 0
    Exit Sub
'
ErrorHandler:
    ErrorRaise "frm_photograph_edit.DeleteKeyword"
End Sub

Private Sub cmdEditKeyword_Click()
'***************************************
' Module/Form Name   : frm_photograph_edit
'
' Procedure Name     : cmdEditKeyword_Click
'
' Purpose            :
'
' Date Created       : 19/07/2005
'
' Author             : Gareth Saunders
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                      19/07/2005 Gareth Saunders
'
'***************************************
'
    On Error GoTo ErrorHandler:
'
'******** Code Starts Here *************
'
    EditKeyword
'
'********* Code Ends Here **************
'
    On Error GoTo 0
    Exit Sub
'
ErrorHandler:
    DisplayError , "frm_photograph_edit.cmdEditKeyword_Click", vbExclamation
End Sub

Private Sub EditKeyword()
    Dim fEditKeyword    As frmEditBatchKeyword
    Dim oKeyword        As BatchKeyword
    Dim strKeyword       As String
    
'***************************************
' Module/Form Name   : frm_photograph_edit
'
' Procedure Name     : EditKeyword
'
' Purpose            :
'
' Date Created       : 19/07/2005
'
' Author             : Gareth Saunders
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                      19/07/2005 Gareth Saunders
'
'***************************************
'
    On Error GoTo ErrorHandler:
'
'******** Code Starts Here *************
'
    Set fEditKeyword = New frmEditBatchKeyword
    strKeyword = msfKeywords.TextMatrix(msfKeywords.row, 0)
    fEditKeyword.Display moBatch, UpdateMode.Edit, strKeyword
    If Not fEditKeyword.Cancel Then
        moBatch.Keywords.RemovePending strKeyword
        moBatch.Keywords.Add moBatch.BatchNo, fEditKeyword.Keyword, "M"
        ApplySynonyms fEditKeyword.Keyword
        DisplayKeywords fEditKeyword.Keyword
    End If
    Set fEditKeyword = Nothing
    mblnManualKeywordsUpdated = True
'
'********* Code Ends Here **************
'
    On Error GoTo 0
    Exit Sub
'
ErrorHandler:
    ErrorRaise "frm_photograph_edit.EditKeyword"
End Sub

Private Sub ApplySynonyms(pstrKeyword As String)

'***************************************
' Module/Form Name   : frm_photograph_edit
'
' Procedure Name     : ApplySynonyms
'
' Purpose            :
'
' Date Created       : 09/10/2005 16:18
'
' Author             : Gareth
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
    Dim oKeywordSynonym     As KeywordSynonym
    Dim oKeywordSynonyms    As KeywordSynonyms
    
    On Error GoTo ErrorProc

    '
    '   Check to see if there are any synonyms.
    '
    Set oKeywordSynonyms = New KeywordSynonyms
    oKeywordSynonyms.Refresh
    For Each oKeywordSynonym In oKeywordSynonyms
        If UCase(pstrKeyword) = UCase(oKeywordSynonym.Keyword) Then
            On Error Resume Next
            moBatch.Keywords.Add moBatch.BatchNo, oKeywordSynonym.Synonym, "S"
            On Error GoTo ErrorProc:
        End If
    Next oKeywordSynonym
    
    On Error GoTo 0
    Exit Sub
'
'********* Code Ends Here **************
'
ErrorProc:
    ErrorRaise "frm_photograph_edit.ApplySynonyms"
End Sub

Private Sub cmdNewKeyword_Click()
'***************************************
' Module/Form Name   : frm_photograph_edit
'
' Procedure Name     : cmdNewKeyword_Click
'
' Purpose            :
'
' Date Created       : 19/07/2005
'
' Author             : Gareth Saunders
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                      19/07/2005 Gareth Saunders
'
'***************************************
'
    On Error GoTo ErrorHandler:
'
'******** Code Starts Here *************
'
    NewKeyword
'
'********* Code Ends Here **************
'
    On Error GoTo 0
    Exit Sub
'
ErrorHandler:
    DisplayError , "frm_photograph_edit.cmdNewKeyword_Click", vbExclamation
End Sub

Private Sub NewKeyword()
    
'***************************************
' Module/Form Name   : frm_photograph_edit
'
' Procedure Name     : NewKeyword
'
' Purpose            :
'
' Date Created       : 19/07/2005
'
' Author             : Gareth Saunders
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                      19/07/2005 Gareth Saunders
'
'***************************************
'
    On Error GoTo ErrorHandler:
'
'******** Code Starts Here *************
'
    Dim fEditKeyword        As frmEditBatchKeyword

    Set fEditKeyword = New frmEditBatchKeyword
    fEditKeyword.Display moBatch, UpdateMode.Add
    If Not fEditKeyword.Cancel Then
        moBatch.Keywords.Add moBatch.BatchNo, fEditKeyword.Keyword, "M"
        '
        ApplySynonyms fEditKeyword.Keyword
        '
        DisplayKeywords fEditKeyword.Keyword
    End If
    Set fEditKeyword = Nothing
'
'********* Code Ends Here **************
'
    On Error GoTo 0
    Exit Sub
'
ErrorHandler:
    ErrorRaise "frm_photograph_edit.NewKeyword"
End Sub

Private Sub cmdScreenRes_Click()
''    frmScreenRes.Display moDigitalImage.FileLocation("SC")
    frmScreenRes.Display moDigitalImage
End Sub

Private Sub cmdSetPicture_Click()
'***************************************
' Module/Form Name   : frm_photograph_edit
'
' Procedure Name     : cmdSetPicture_Click
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
On Error GoTo cmdSetPicture_Click_Error
'
'******** Code Starts Here *************
'
    Dim strFileLocation As String
    Dim strNewFileLocation As String
    
    strFileLocation = moDigitalImage.FileLocation
    strNewFileLocation = moDigitalImage.NewFileLocation
    
    With dlgPhotograph
        .CancelError = True
        .DialogTitle = "Select Picture File"
        If strNewFileLocation = "" Then
            If strFileLocation = "" Then
                '.InitDir = oFSO.BuildPath(App.Path, goSystemConfig.ImageHoldingArea)
                .InitDir = oFSO.BuildPath(goSystemConfig.ServerLocation, goSystemConfig.ImageHoldingArea)
            Else
                .InitDir = StripFile(strFileLocation)
            End If
        Else
            .InitDir = strNewFileLocation
        End If
        .Filter = "Picture Files (*.jpg)|*.jpg"
        If strNewFileLocation = "" Then
            If strFileLocation = "" Then
                .FileName = ""
            Else
                .FileName = CStr(moDigitalImage.BatchNo)
            End If
        Else
            .FileName = StripPath(moDigitalImage.NewFileLocation)
        End If
        .ShowOpen
    End With
    
    Screen.MousePointer = vbHourglass
    SaveSetting App.Title, "PhotographEdit", "ImageFile", StripFile(dlgPhotograph.FileName)
    DoEvents
    imgPhotograph.Picture = LoadPicture(dlgPhotograph.FileName)
    moDigitalImage.NewFileLocation = dlgPhotograph.FileName
    SizePhotograph
    'cmdRemovePicture.Enabled = True
    Screen.MousePointer = vbDefault
'
'********* Code Ends Here **************
'
   Exit Sub
'
cmdSetPicture_Click_Error:
    If Err.Number = cdlCancel Then
        Exit Sub
    Else
        DisplayError , "frm_photograph_edit.cmdSetPicture_Click", vbExclamation
    End If
End Sub

Private Sub com_delete_Click()
'***************************************
' Module/Form Name   : frm_photograph_edit
'
' Procedure Name     : com_delete_Click
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
On Error GoTo com_delete_Click_Error
'
'******** Code Starts Here *************
'
    If datBatch.Recordset.RecordCount = 0 Then
        MsgBox "There are no images to delete", vbExclamation, "Delete Images"
        Exit Sub
    Else
        DeleteImages abort
    End If
'
'   If error or deletion cancelled, exit sub.
'
    If abort Then
        Exit Sub
    End If
    
'    If current_image_no(1) > datBatch.Recordset.RecordCount Then
'        current_image_no(1) = current_image_no(1) - 1
'        If current_image_no(1) = 0 Then
'            Exit Sub
'        End If
'    End If
    
    redisplay_images
    txt_no_of_images = photograph.count_images(snap_batch!batch_no)

    update_loaded_screens
    
'
'********* Code Ends Here **************
'
   Exit Sub
'
com_delete_Click_Error:
    DisplayError , "frm_photograph_edit.com_delete_Click", vbExclamation
End Sub

Private Sub com_delivery_note_Click()
'***************************************
' Module/Form Name   : frm_photograph_edit
'
' Procedure Name     : com_delivery_note_Click
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
On Error GoTo com_delivery_note_Click_Error
'
'******** Code Starts Here *************
'
    Dim lngPhotoNo    As Long
    
    If msfImages.TextMatrix(msfImages.row, 0) = CONST_DIGITAL Then
        lngPhotoNo = CLng(snap_batch!batch_no)
    Else
        lngPhotoNo = CLng(CStr(snap_batch!batch_no) + "." + CStr(msfImages.TextMatrix(msfImages.row, 0)))
    End If

    frm_delivery_note_return_photos.load_delivery_note msfImages.TextMatrix(msfImages.row, 1), _
                                                       0, _
                                                       msfImages.TextMatrix(msfImages.row, 6), _
                                                       lngPhotoNo, _
                                                       False
'
'********* Code Ends Here **************
'
   Exit Sub
'
com_delivery_note_Click_Error:
    DisplayError , "frm_photograph_edit.com_delivery_note_Click", vbExclamation
End Sub

Private Sub com_job_Click()
'***************************************
' Module/Form Name   : frm_photograph_edit
'
' Procedure Name     : com_job_Click
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
On Error GoTo com_job_Click_Error
'
'******** Code Starts Here *************
'
    Dim oSearch     As New Search2
    Dim oJob        As New Job2
    
    With oSearch
        .Read msfImages.TextMatrix(msfImages.row, 8)
        oJob.Read .JobNo
    End With
    '
    frm_search.display_search_form Edit, _
                                   oJob, _
                                   oSearch.SearchNo

    Set oSearch = Nothing
    Set oJob = Nothing
'
'********* Code Ends Here **************
'
   Exit Sub
'
com_job_Click_Error:
    DisplayError , "frm_photograph_edit.com_job_Click", vbExclamation
End Sub

Private Sub com_new_Click()
'***************************************
' Module/Form Name   : frm_photograph_edit
'
' Procedure Name     : com_new_Click
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
On Error GoTo com_new_Click_Error
'
'******** Code Starts Here *************
'
    frm_image_new.load_photograph snap_batch!batch_no
'
'********* Code Ends Here **************
'
   Exit Sub
'
com_new_Click_Error:
    DisplayError , "frm_photograph_edit.com_new_Click", vbExclamation
End Sub

Private Sub com_undelete_Click()
'***************************************
' Module/Form Name   : frm_photograph_edit
'
' Procedure Name     : com_undelete_Click
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
On Error GoTo com_undelete_Click_Error
'
'******** Code Starts Here *************
'
    If datBatch.Recordset.RecordCount = 0 Then
        MsgBox "There are no images to undelete", vbExclamation, "Undelete Images"
        Exit Sub
    Else
        UnDeleteImages abort
    End If
'
'   If error or deletion cancelled, exit sub.
'
    If abort Then
        Exit Sub
    End If
    
'    If current_image_no(1) > datBatch.Recordset.RecordCount Then
'        current_image_no(1) = current_image_no(1) - 1
'        If current_image_no(1) = 0 Then
'            Exit Sub
'        End If
'    End If
    
    redisplay_images
    txt_no_of_images = photograph.count_images(snap_batch!batch_no)
    
    update_loaded_screens
'
'********* Code Ends Here **************
'
   Exit Sub
'
com_undelete_Click_Error:
    DisplayError , "frm_photograph_edit.com_undelete_Click", vbExclamation
End Sub

Private Sub Form_Initialize()
    mintKeywordSortCol = 0
    mintKeywordSortOrder = flexSortGenericAscending
    mstrCurrentKeyword = ""
    mintCurrentKeywordCol = 0
    mblnManualKeywordsUpdated = False
End Sub

Private Sub Form_Load()
'***************************************
' Module/Form Name   : frm_photograph_edit
'
' Procedure Name     : Form_Load
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
On Error GoTo Form_Load_Error
'
'******** Code Starts Here *************
'
    Dim tab_idx As Integer
    mbolDontResize = True
    current_image_no(1) = 0
    current_image_no(2) = 0
    image_selected(1) = False
    image_selected(2) = False
    mod_pic_moved(1) = "N"
    mod_pic_moved(2) = "N"
    mod_pic_moved(3) = "N"
'    dat_batch.DatabaseName = glo_dbname
'
'   Set the width & height of the form.
'
    current_tab = 1
    Me.WindowState = GetSetting(App.Title, "PhotographEdit", "WindowState", 0)
    If GetSetting(App.Title, "PhotographEdit", "WindowState", 0) <> 2 Then
        Me.Width = GetSetting(App.Title, "PhotographEdit", "Width", 8025)
        Me.Height = GetSetting(App.Title, "PhotographEdit", "Height", 6000)
    End If
'
    '
    '   Set up the Keywords Grid.
    '
    With msfKeywords
        .Cols = 2
        .Rows = 1
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignLeftCenter
    End With
    '
    com_position_form Me
    mbolDontResize = False
    Show
    txt_description.SetFocus
'
'********* Code Ends Here **************
'
   Exit Sub
'
Form_Load_Error:
    DisplayError , "frm_photograph_edit.Form_Load", vbExclamation
End Sub

Private Sub Form_Resize()
'***************************************
' Module/Form Name   : frm_photograph_edit
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
    Dim header_height As Integer
    
    MIN_WIDTH = 8025
    MIN_HEIGHT = 6200
    header_height = 440
    
    If Not ResizeForm(Me) Then
        Exit Sub
    End If
    
    If mbolDontResize Then
        Exit Sub
    End If
    
    If WindowState = 1 Then
        Exit Sub
    End If
  
    If WindowState = 0 Then
        If Width < MIN_WIDTH Then
            Width = MIN_WIDTH
        End If
        If Height < MIN_HEIGHT Then
            Height = MIN_HEIGHT
        End If
    End If
'
'   Reposition Done & Cancel buttons.
'
    cmdOK.Top = Height - cmdOK.Height - header_height - 200
    cmdOK.Left = (Width - (cmdOK.Width + cmdApply.Width + cmdCancel.Width + 200)) / 2
    cmdApply.Move cmdOK.Left + cmdOK.Width + 100, cmdOK.Top
    cmdCancel.Move cmdApply.Left + cmdApply.Width + 100, cmdOK.Top
    com_new.Top = pic_batch(2).Height - com_new.Height
'
'   Resize tab.
'
    tab_batch.Width = Width - tab_batch.Left * 2 - 200
    tab_batch.Height = cmdOK.Top - tab_batch.Top - 100
'
'   Reset picture variables to not sized and then call sub to resize current selection.
'
    mod_pic_moved(1) = "N"
    mod_pic_moved(2) = "N"
    mod_pic_moved(3) = "N"
    tab_select current_tab
'
'********* Code Ends Here **************
'
   Exit Sub
'
Form_Resize_Error:
    DisplayError , "frm_photograph_edit.Form_Resize", vbExclamation
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set moBatch = Nothing
    SaveSetting App.Title, "PhotographEdit", "Width", Me.Width
    SaveSetting App.Title, "PhotographEdit", "Height", Me.Height
    SaveSetting App.Title, "PhotographEdit", "WindowState", Me.WindowState
    On Error Resume Next
    gcolMaxedWindows.Remove CStr(Me.hWnd)
End Sub

Private Sub msfImages_Click()
    '
    '   Set the pseudo book marks for the images.
    '
    If chk_archived_images Then
        current_image_no(2) = msfImages.row
    Else
        current_image_no(1) = msfImages.row
    End If
  
End Sub

Private Sub msfImages_SelChange()
'***************************************
' Module/Form Name   : frm_photograph_edit
'
' Procedure Name     : msfImages_SelChange
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
On Error GoTo msfImages_SelChange_Error
'
'******** Code Starts Here *************
'
    SetImageManipulationButtons
'
'********* Code Ends Here **************
'
   Exit Sub
'
msfImages_SelChange_Error:
    DisplayError , "frm_photograph_edit.msfImages_SelChange", vbExclamation
End Sub

Private Sub cmdCancel_Click()
    frm_photograph_edit.Enabled = False
    Unload Me
    Set frm_photograph_edit = Nothing
End Sub

Private Sub cmdOK_Click()
'***************************************
' Module/Form Name   : frm_photograph_edit
'
' Procedure Name     : cmdOK_Click
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
On Error GoTo cmdOK_Click_Error
'
'******** Code Starts Here *************
'
    If Not valid_input Then
        Exit Sub
    End If
    update_batch
    If Not retry_update Then
        update_loaded_screens
        Unload Me
        Set frm_photograph_edit = Nothing
  Else
      load_batch snap_batch!batch_no
    End If
'
'********* Code Ends Here **************
'
   Exit Sub
'
cmdOK_Click_Error:
    DisplayError , "frm_photograph_edit.cmdOK_Click", vbExclamation
End Sub

Private Sub msfKeywords_Click()
    Dim intCol              As Integer
    Dim i                   As Integer
    Dim strOriginalKeyword  As String
    
'***************************************
' Module/Form Name   : frm_photograph_edit
'
' Procedure Name     : msfKeywords_Click
'
' Purpose            :
'
' Date Created       : 19/07/2005
'
' Author             : Gareth Saunders
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                      19/07/2005 Gareth Saunders
'
'***************************************
'
    On Error GoTo ErrorHandler:
'
'******** Code Starts Here *************
'

    msfKeywords.Redraw = False
    If msfKeywords.MouseRow > 0 Then
        mstrCurrentKeyword = msfKeywords.TextMatrix(msfKeywords.RowSel, 0)
        msfKeywords.Redraw = True
        Exit Sub
    End If
    '
    If mintCurrentKeywordCol <> mintKeywordSortCol Then
        mintKeywordSortCol = mintCurrentKeywordCol
        mintKeywordSortOrder = flexSortGenericAscending
    Else
        mintKeywordSortOrder = IIf(mintKeywordSortOrder = flexSortGenericAscending, flexSortGenericDescending, flexSortGenericAscending)
    End If
    '
    With msfKeywords
        .col = mintKeywordSortCol
        .ColSel = mintKeywordSortCol
        .Sort = mintKeywordSortOrder
    End With
    '
    '   Get the original row back.
    '
    For i = 1 To msfKeywords.Rows - 1
        If msfKeywords.TextMatrix(i, 0) = mstrCurrentKeyword Then
            msfKeywords.row = i
            msfKeywords.col = 0
            msfKeywords.RowSel = i
            msfKeywords.ColSel = 1
        End If
    Next i
    msfKeywords.Redraw = True
'
'********* Code Ends Here **************
'
    On Error GoTo 0
    Exit Sub
'
ErrorHandler:
    DisplayError , "frm_photograph_edit.msfKeywords_Click", vbExclamation
End Sub

Private Sub msfKeywords_DblClick()
'***************************************
' Module/Form Name   : frm_photograph_edit
'
' Procedure Name     : msfKeywords_DblClick
'
' Purpose            :
'
' Date Created       : 19/07/2005
'
' Author             : Gareth Saunders
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                      19/07/2005 Gareth Saunders
'
'***************************************
'
    On Error GoTo ErrorHandler:
'
'******** Code Starts Here *************
'

    If msfKeywords.MouseRow > 0 Then
        EditKeyword
    End If
'
'********* Code Ends Here **************
'
    On Error GoTo 0
    Exit Sub
'
ErrorHandler:
    DisplayError , "frm_photograph_edit.msfKeywords_DblClick", vbExclamation
End Sub

Private Sub msfKeywords_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X <= msfKeywords.ColWidth(0) Then
        mintCurrentKeywordCol = 0
    Else
        mintCurrentKeywordCol = 1
    End If
    
    If Button = vbRightButton Then
        Set mdi_npls.fPopUp = Me
        PopupMenu mdi_npls.mnuPhotographEditPopups, vbPopupMenuRightButton, , , mdi_npls.mnuPhotographEditPopupsEdit
    End If

End Sub

Private Sub tab_batch_Click()
'***************************************
' Module/Form Name   : frm_photograph_edit
'
' Procedure Name     : tab_batch_Click
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
On Error GoTo tab_batch_Click_Error
'
'******** Code Starts Here *************
'
    If tab_batch.SelectedItem.Index = 3 Then
        If mod_images_displayed = False Then
            Screen.MousePointer = vbHourglass
            display_images
            SetImageManipulationButtons
            Screen.MousePointer = vbDefault
            mod_images_displayed = True
        End If
    End If
    
    current_tab = tab_batch.SelectedItem.Index
    tab_select tab_batch.SelectedItem.Index
'
'********* Code Ends Here **************
'
   Exit Sub
'
tab_batch_Click_Error:
    DisplayError , "frm_photograph_edit.tab_batch_Click", vbExclamation
End Sub

Private Sub txt_copyright_GotFocus()
'    txt_copyright.SelStart = 0
'    txt_copyright.SelLength = Len(txt_copyright)
End Sub

Private Sub txt_credit_GotFocus()
'    txt_credit.SelStart = 0
'    txt_credit.SelLength = Len(txt_credit)
End Sub

Private Sub txt_description_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub SetImageManipulationButtons()
'***************************************
' Module/Form Name   : frm_photograph_edit
'
' Procedure Name     : SetImageManipulationButtons
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
On Error GoTo SetImageManipulationButtons_Error
'
'******** Code Starts Here *************
'
'   Set the image manipulation buttons according to the selection.
'
    Dim strStatus   As String
    
    If Abs(msfImages.RowSel - msfImages.row) < 1 Then
        strStatus = UCase(msfImages.TextMatrix(msfImages.row, 2))
        Select Case strStatus
            Case Is = "I"
                com_delete.Enabled = True
                com_delivery_note.Enabled = False
                com_job.Enabled = False
            Case Is = "P"
                If UCase(msfImages.TextMatrix(msfImages.row, 0)) = "DIGITAL" Then
                    com_delete.Enabled = False
                Else
                    com_delete.Enabled = True
                End If
                com_delivery_note.Enabled = False
                com_job.Enabled = True
            Case Is = "O"
                If UCase(msfImages.TextMatrix(msfImages.row, 0)) = "DIGITAL" Then
                    com_delete.Enabled = False
                Else
                    com_delete.Enabled = True
                End If
                com_delivery_note.Enabled = True
                com_job.Enabled = False
            Case Is = "R", "INV"
                com_delete.Enabled = False
                com_delivery_note.Enabled = True
                com_job.Enabled = False
            Case Else
                '
                '   Don't allow the deletion of a digital image.
                '
                com_delete.Enabled = False
                com_delivery_note.Enabled = False
                com_job.Enabled = False
        End Select
    Else
        com_delete.Enabled = False
        com_delivery_note.Enabled = False
        com_job.Enabled = False
    End If
    
'
'********* Code Ends Here **************
'
   Exit Sub
'
SetImageManipulationButtons_Error:
    ErrorRaise "frm_photograph_edit.SetImageManipulationButtons"
End Sub

Private Sub SizePhotograph()
'***************************************
' Module/Form Name   : frm_photograph_edit
'
' Procedure Name     : SizePhotograph
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
On Error GoTo SizePhotograph_Error
'
'******** Code Starts Here *************
'
    With moDigitalImage
        .MaxHeight = fraImageFile.Height - 300
        '.MaxWidth = cmdSetPicture.Left - 200
        .MaxWidth = 32767
        Set .Picture = imgPhotograph.Picture
        imgPhotograph.Visible = False
        imgPhotograph.Top = 200
        imgPhotograph.Left = 100
        imgPhotograph.Width = .Width
        imgPhotograph.Height = .Height
        imgPhotograph.Visible = True
    End With
'
'********* Code Ends Here **************
'
   Exit Sub
'
SizePhotograph_Error:
    ErrorRaise "frm_photograph_edit.SizePhotograph"
End Sub

Public Sub ForceResize()
    gblnResizeMaxedWindows = False
    Form_Resize
    gblnResizeMaxedWindows = True
End Sub

