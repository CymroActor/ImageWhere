VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_photograph_view 
   Caption         =   "View Photograph"
   ClientHeight    =   5880
   ClientLeft      =   2145
   ClientTop       =   1635
   ClientWidth     =   8250
   Icon            =   "Photograph View.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5880
   ScaleWidth      =   8250
   Begin ImageWhere.SimpleGrid smgDigitalImages 
      Height          =   1335
      Left            =   120
      TabIndex        =   38
      Top             =   1500
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   2355
      Columns         =   1
      KeyCol          =   0
   End
   Begin ImageWhere.SimpleGrid smgPhotograph 
      Height          =   1335
      Left            =   120
      TabIndex        =   40
      Top             =   960
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   2355
      Columns         =   1
      KeyCol          =   0
   End
   Begin VB.PictureBox pic_batch 
      Height          =   3375
      Index           =   3
      Left            =   600
      ScaleHeight     =   3315
      ScaleWidth      =   7095
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   1320
      Width           =   7155
      Begin VB.TextBox txt_batch_no 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   3
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton cmdDigitalReturn 
         Caption         =   "Re&turn"
         Height          =   315
         Left            =   5400
         TabIndex        =   36
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label lab_batch_no_lab 
         Caption         =   "Batch Number:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   60
         TabIndex        =   39
         Top             =   180
         Width           =   1215
      End
   End
   Begin VB.PictureBox pic_batch 
      Height          =   3615
      Index           =   2
      Left            =   480
      ScaleHeight     =   3555
      ScaleWidth      =   7455
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   1080
      Width           =   7515
      Begin VB.TextBox txt_batch_no 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   2
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   180
         Width           =   735
      End
      Begin VB.CheckBox chk_archived_images 
         Caption         =   "Archived Images"
         Height          =   195
         Left            =   5220
         TabIndex        =   30
         Top             =   180
         Width           =   1695
      End
      Begin VB.CommandButton com_delivery_note 
         Caption         =   "Re&turn"
         Height          =   315
         Left            =   5100
         TabIndex        =   31
         Top             =   3180
         Width           =   1095
      End
      Begin VB.CommandButton com_job 
         Caption         =   "&Job"
         Height          =   315
         Left            =   6300
         TabIndex        =   32
         Top             =   3180
         Width           =   1095
      End
      Begin VB.Label lab_batch_no_lab 
         Caption         =   "Batch Number:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   60
         TabIndex        =   28
         Top             =   180
         Width           =   1215
      End
   End
   Begin VB.PictureBox pic_batch 
      Height          =   4215
      Index           =   0
      Left            =   360
      ScaleHeight     =   4155
      ScaleWidth      =   7455
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   720
      Width           =   7515
      Begin VB.CheckBox chkExclusive 
         Caption         =   "Check1"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1920
         TabIndex        =   16
         Top             =   4080
         Value           =   2  'Grayed
         Width           =   255
      End
      Begin VB.TextBox txtTiffLocation 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   3660
         Width           =   1335
      End
      Begin VB.Frame fraImageFile 
         Caption         =   "Digital Image"
         Height          =   2055
         Left            =   3360
         TabIndex        =   17
         Top             =   1680
         Width           =   2175
         Begin VB.Image imgPhotograph 
            Height          =   1740
            Left            =   120
            Stretch         =   -1  'True
            Top             =   240
            Width           =   1860
         End
      End
      Begin VB.TextBox txt_format 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   3240
         Width           =   495
      End
      Begin VB.TextBox txt_no_of_photographs 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   2820
         Width           =   675
      End
      Begin VB.TextBox txt_batch 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox txt_date_taken 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1860
         Width           =   975
      End
      Begin VB.TextBox txt_description 
         BackColor       =   &H8000000F&
         Height          =   975
         Left            =   1920
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   600
         Width           =   5175
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
      Begin VB.Label lblExclusive 
         Caption         =   "Exclusive:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   60
         TabIndex        =   15
         Top             =   4080
         Width           =   1455
      End
      Begin VB.Label lblTiffLocation 
         Caption         =   "Tiff Location:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   60
         TabIndex        =   13
         Top             =   3660
         Width           =   1455
      End
      Begin VB.Label lab_batch_no_lab 
         Caption         =   "Photograph Number:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   1
         Top             =   180
         Width           =   1575
      End
      Begin VB.Label lab_description_lab 
         Caption         =   "Description:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   60
         TabIndex        =   3
         Top             =   600
         Width           =   1155
      End
      Begin VB.Label lab_date_taken_lab 
         Caption         =   "Date Taken:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   60
         TabIndex        =   5
         Top             =   1860
         Width           =   1035
      End
      Begin VB.Label lab_batch_lab 
         Caption         =   "Batch:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   60
         TabIndex        =   7
         Top             =   2400
         Width           =   675
      End
      Begin VB.Label lab_no_of_photographs_lab 
         Caption         =   "No. of photographs:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   60
         TabIndex        =   9
         Top             =   2820
         Width           =   1455
      End
      Begin VB.Label lab_format_lab 
         Caption         =   "Format:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   60
         TabIndex        =   11
         Top             =   3240
         Width           =   1455
      End
   End
   Begin VB.PictureBox pic_batch 
      Height          =   3495
      Index           =   1
      Left            =   240
      ScaleHeight     =   3435
      ScaleWidth      =   7455
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   540
      Width           =   7515
      Begin VB.TextBox txtKeywords 
         BackColor       =   &H8000000F&
         Height          =   675
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   1000
         MultiLine       =   -1  'True
         TabIndex        =   26
         Top             =   2400
         Width           =   5355
      End
      Begin VB.TextBox txt_batch_no 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   1
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   180
         Width           =   735
      End
      Begin VB.TextBox txt_credit 
         BackColor       =   &H8000000F&
         Height          =   675
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   45
         MultiLine       =   -1  'True
         TabIndex        =   24
         Top             =   1560
         Width           =   5355
      End
      Begin VB.TextBox txt_copyright 
         BackColor       =   &H8000000F&
         Height          =   675
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   45
         MultiLine       =   -1  'True
         TabIndex        =   22
         Top             =   720
         Width           =   5355
      End
      Begin VB.Label lblKeywords 
         Caption         =   "Keywords:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   60
         TabIndex        =   25
         Top             =   2400
         Width           =   1155
      End
      Begin VB.Label lab_batch_no_lab 
         Caption         =   "Photograph Number:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   19
         Top             =   180
         Width           =   1695
      End
      Begin VB.Label lab_credit_lab 
         Caption         =   "Credit:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   60
         TabIndex        =   23
         Top             =   1560
         Width           =   1155
      End
      Begin VB.Label lab_copyright_lab 
         Caption         =   "Copyright:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   60
         TabIndex        =   21
         Top             =   720
         Width           =   1155
      End
   End
   Begin VB.Data datPhotograph 
      Caption         =   "datPhotograph"
      Connect         =   "Access 2000;"
      DatabaseName    =   "E:\Image Where\Development\Database\Library7.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5040
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   675
      Left            =   3420
      Picture         =   "Photograph View.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   5040
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin MSComctlLib.TabStrip tab_batch 
      Height          =   4995
      Left            =   60
      TabIndex        =   34
      Top             =   60
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   8811
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Credit/Copyright"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Transparencies"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Digital Images"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_photograph_view"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private snap_batch As DAO.Recordset

Private SQL                 As String
'
'   Public item used to accept the selected Batch number
'   from Batch Maintenance
'
Public view_batch_number    As String

Private mod_pic_moved(4)    As String
Public mod_images_displayed As Boolean
Private current_tab         As Integer
Private current_image_no(2) As Integer
Private image_selected(2)   As Boolean
Private moDigitalImage      As DigitalImage
Private mblnDontResize      As Boolean
Private moBatch             As Batch2

Private Sub chk_archived_images_Click()
'***************************************
' Module/Form Name   : frm_photograph_view
'
' Procedure Name     : chk_archived_images_Click
'
' Purpose            :
'
' Date Created       : 13/12/2002
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
        com_delivery_note.Visible = False
        com_job.Visible = False
    Else
        com_delivery_note.Visible = True
        com_job.Visible = True
    End If
    '
    '   Update SQL and refresh.
    '
    photo_image.Image_Display snap_batch!batch_no, _
                              SQL, _
                              chk_archived_images.Value

    datPhotograph.RecordSource = SQL

    RedisplayTransparencies

'
'********* Code Ends Here **************
'
   Exit Sub
'
chk_archived_images_Click_Error:
    DisplayError , "frm_photograph_view.chk_archived_images_Click", vbExclamation
End Sub

Private Sub cmdDigitalReturn_Click()
    frm_delivery_note_return_photos.load_delivery_note smgDigitalImages.Column(1).Value, _
                                                       0, _
                                                       smgDigitalImages.Column(5).Value, _
                                                       LTrim(CStr(snap_batch!batch_no)) + ".0", _
                                                       False
End Sub

Private Sub cmdExit_Click()
    SaveSetting App.Title, "PictureView", "Width", Me.Width
    SaveSetting App.Title, "PictureView", "Height", Me.Height
    SaveSetting App.Title, "PictureView", "WindowState", Me.WindowState
    Unload Me
    Set frm_photograph_view = Nothing

End Sub

Private Sub com_delivery_note_Click()
'***************************************
' Module/Form Name   : frm_photograph_view
'
' Procedure Name     : com_delivery_note_Click
'
' Purpose            :
'
' Date Created       : 13/12/2002
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
    frm_delivery_note_return_photos.load_delivery_note smgPhotograph.Column(2).Value, _
                                                       0, _
                                                       smgPhotograph.Column(7).Value, _
                                                       LTrim(CStr(snap_batch!batch_no)) + "." + LTrim(CStr(smgPhotograph.Column(1).Value)), _
                                                       False
'
'********* Code Ends Here **************
'
   Exit Sub
'
com_delivery_note_Click_Error:
    DisplayError , "frm_photograph_view.com_delivery_note_Click", vbExclamation
End Sub

Private Sub com_job_Click()
'***************************************
' Module/Form Name   : frm_photograph_view
'
' Procedure Name     : com_job_Click
'
' Purpose            :
'
' Date Created       : 13/12/2002
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
        .Read CLng(smgPhotograph.Column(9).Value)
        oJob.Read .JobNo
    End With
     
    frm_search.display_search_form Edit, _
                                   oJob, _
                                   oSearch.SearchNo
'
'********* Code Ends Here **************
'
   Exit Sub
'
com_job_Click_Error:
    DisplayError , "frm_photograph_view.com_job_Click", vbExclamation
End Sub

Private Sub display_fields()
'***************************************
' Module/Form Name   : frm_photograph_view
'
' Procedure Name     : display_fields
'
' Purpose            :
'
' Date Created       : 13/12/2002
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
    Dim oKeyword    As BatchKeyword
    Dim strKeywords As String
    
    Me.Show
    If IsNull(snap_batch!batch_no) Then
        txt_batch_no(0) = ""
        txt_batch_no(1) = ""
        txt_batch_no(2) = ""
        txt_batch_no(3) = ""
    Else
        txt_batch_no(0) = snap_batch!batch_no
        txt_batch_no(1) = snap_batch!batch_no
        txt_batch_no(2) = snap_batch!batch_no
        txt_batch_no(3) = snap_batch!batch_no
    End If

    If IsNull(snap_batch!Description) Then
        txt_description.Text = ""
    Else
        txt_description.Text = snap_batch!Description & ""
    End If

    If IsNull(snap_batch!date_taken) Then
        txt_date_taken.Text = ""
    Else
        txt_date_taken.Text = Format(snap_batch!date_taken, "mm/yyyy")
    End If

    If IsNull(snap_batch!parent_batch) Or snap_batch!parent_batch = 0 Or _
              snap_batch!parent_batch = snap_batch!batch_no Then
        txt_batch = ""
    Else
        txt_batch = snap_batch!parent_batch
    End If
    txt_no_of_photographs.Text = photograph.count_images(snap_batch!batch_no)
    '
    '   Set up formats in the combo box.
    '
    txt_format.Text = snap_batch!Format & ""
    '
    txt_copyright = snap_batch!CopyRight & ""
    txt_credit = snap_batch!Credit & ""
    txtTiffLocation.Text = snap_batch!TiffLocation & ""
    chkExclusive.Value = IIf(snap_batch!Exclusive, vbChecked, vbUnchecked)
    '
    '   Get the keywords and display.
    '
    strKeywords = ""
    For Each oKeyword In moBatch.Keywords
        strKeywords = strKeywords & oKeyword.Keyword & " "
    Next oKeyword
    If Trim(strKeywords) <> "" Then
        txtKeywords.Text = Left(strKeywords, Len(strKeywords) - 1)
    Else
        txtKeywords.Text = ""
    End If
    '
    '   Set up ImageFile and store location in memory.
    '
    DoEvents
    imgPhotograph.Picture = LoadPicture(moDigitalImage.FileLocation)
    SizePhotograph
'
'********* Code Ends Here **************
'
   Exit Sub
'
display_fields_Error:
    ErrorRaise "frm_photograph_view.display_fields"
End Sub

Public Sub DisplayTransparencies()
'***************************************
' Module/Form Name   : frm_photograph_view
'
' Procedure Name     : DisplayTransparencies
'
' Purpose            :
'
' Date Created       : 13/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo DisplayTransparencies_Error
'
'******** Code Starts Here *************
'
    Dim strImageNo      As String
    Dim strStatus       As String
    '
    '   Set up data control access.
    '
    photo_image.Image_Display snap_batch!batch_no, _
                              SQL, _
                              chk_archived_images.Value

    datPhotograph.RecordSource = SQL
    datPhotograph.DatabaseName = glo_dbname
    datPhotograph.Refresh
    '
    '   Display the images.
    '
    With smgPhotograph
        .Redraw = False
        .Clear
        If Not datPhotograph.Recordset.EOF Then datPhotograph.Recordset.MoveFirst
        While Not datPhotograph.Recordset.EOF
            '
            '   If the Photograph number is 0 then this is a digital image.
            '
            If datPhotograph.Recordset!ImageNo = 0 Then
                strImageNo = "DIGITAL"
                strStatus = ""
            Else
                strImageNo = datPhotograph.Recordset!ImageNo
                strStatus = datPhotograph.Recordset!Status
            End If
            
            .AddRow False, _
                    strImageNo, _
                    datPhotograph.Recordset!Delivery, _
                    strStatus, _
                    datPhotograph.Recordset![Customer Name], _
                    datPhotograph.Recordset![Address Line], _
                    datPhotograph.Recordset![Barcode Read], _
                    datPhotograph.Recordset!customer_no, _
                    datPhotograph.Recordset!job_no, _
                    datPhotograph.Recordset!search_no
            datPhotograph.Recordset.MoveNext
        Wend
        .Redraw = True
    End With
    EnableButtons
'
'********* Code Ends Here **************
'
   Exit Sub
'
DisplayTransparencies_Error:
    ErrorRaise "frm_photograph_view.DisplayTransparencies"
End Sub

Public Sub DisplayDigitalImages()
    Dim strImageNo      As String
    Dim strStatus       As String
    Dim oDigitalImage   As DigitalImage
    '
    '   Display the images.
    '
    With smgDigitalImages
        .Redraw = False
        .Clear
        For Each oDigitalImage In moBatch.DigitalImages
            .AddRow False, _
                    oDigitalImage.DeliveryNoteNo, _
                    (oDigitalImage.DigitalStatus = "R"), _
                    oDigitalImage.CustomerName, _
                    oDigitalImage.AddressLine1, _
                    oDigitalImage.CustomerNo
        Next oDigitalImage
        .Redraw = True
    End With
    EnableButtons
End Sub

Private Sub EnableButtons()
'***************************************
' Module/Form Name   : frm_photograph_view
'
' Procedure Name     : EnableButtons
'
' Purpose            :
'
' Date Created       : 13/12/2002
'
' Author             : GARETH SAUNDERS
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
    If smgPhotograph.CurrentRow <> 0 Then
        '
        '   Set the image manipulation buttons according to the selection.
        '
        If smgPhotograph.Column(3).Value = "I" Then
            com_delivery_note.Enabled = False
            com_job.Enabled = False
        Else
            If smgPhotograph.Column(3).Value = "P" Then
                com_delivery_note.Enabled = False
                com_job.Enabled = True
            Else
                If smgPhotograph.Column(3).Value = "O" Then
                    com_delivery_note.Enabled = True
                    com_job.Enabled = False
                Else
                    com_delivery_note.Enabled = False
                    com_job.Enabled = False
                End If
            End If
        End If
    Else
        com_delivery_note.Enabled = False
        com_job.Enabled = False
    End If
    '
    '   Digital Images.
    '
    If smgDigitalImages.CurrentRow <> 0 Then
        cmdDigitalReturn.Enabled = True
    Else
        cmdDigitalReturn.Enabled = False
    End If
'
'********* Code Ends Here **************
'
   Exit Sub
'
EnableButtons_Error:
    ErrorRaise "frm_photograph_view.EnableButtons"
End Sub

Private Sub Form_Initialize()
    mblnDontResize = False
End Sub

Private Sub Form_Load()
'***************************************
' Module/Form Name   : frm_photograph_view
'
' Procedure Name     : Form_Load
'
' Purpose            :
'
' Date Created       : 13/12/2002
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

    current_image_no(1) = 0
    current_image_no(2) = 0
    image_selected(1) = False
    image_selected(2) = False
    mod_pic_moved(1) = "N"
    mod_pic_moved(2) = "N"
    mod_pic_moved(3) = "N"
    mod_pic_moved(4) = "N"
    datPhotograph.DatabaseName = glo_dbname
'
'   Set the width & height of the form.
'
    current_tab = 1
    Me.WindowState = GetSetting(App.Title, "PictureView", "WindowState", 0)
    If GetSetting(App.Title, "PictureView", "WindowState", 0) <> 2 Then
        mblnDontResize = True
        Me.Width = GetSetting(App.Title, "PictureView", "Width", 8025)
        Me.Height = GetSetting(App.Title, "PictureView", "Height", 6000)
        mblnDontResize = False
    End If
    '
    '   Set the Grids.
    '
    Set smgPhotograph.Container = pic_batch(2)
    With smgPhotograph
        .Columns = 9
        .KeyCol = 1
        .Column(1).Align = flexAlignRightTop
        .Column(1).Header = "Image"
        .Column(2).Header = "Delivery"
        .Column(2).Align = flexAlignRightTop
        .Column(3).Header = "Out/Pending"
        .Column(4).Header = "Customer Name"
        .Column(5).Header = "Address Line 1"
        .Column(6).Header = "Scanned"
        .ColumnProperty(6).PropertyType = smgPropTypeBooleanGraphic
        .Column(7).Header = "Customer No."
        .Column(8).Header = "Job No."
        .Column(9).Header = "Search No."
    End With
    '
    Set smgDigitalImages.Container = pic_batch(3)
    With smgDigitalImages
        .Columns = 4
        .KeyCol = 1
        .Column(1).Header = "Delivery"
        .Column(1).Align = flexAlignRightTop
        .Column(2).Header = "Rejected"
        .ColumnProperty(2).PropertyType = smgPropTypeBooleanGraphic
        .Column(3).Header = "Customer Name"
        .Column(4).Header = "Address Line 1"
        .Column(5).Header = "Customer No."
    End With
    '
    com_delivery_note.Enabled = False
    com_job.Enabled = False
'
    com_position_form Me
'
'********* Code Ends Here **************
'
   Exit Sub
'
Form_Load_Error:
    DisplayError , "frm_photograph_view.Form_Load", vbExclamation
End Sub

Private Sub Form_Resize()
'***************************************
' Module/Form Name   : frm_photograph_view
'
' Procedure Name     : Form_Resize
'
' Purpose            :
'
' Date Created       : 13/12/2002
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

    'If mblnDontResize Then Exit Sub

    MIN_WIDTH = 8025
    MIN_HEIGHT = 6300
    header_height = 440

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
'
'   Reposition Done & Cancel buttons.
'
    cmdExit.Top = Height - cmdExit.Height - header_height - 150
    cmdExit.Left = (Width - cmdExit.Width - 260)
'
'   Resize tab.
'
    tab_batch.Width = Width - tab_batch.Left * 2 - 200
    tab_batch.Height = cmdExit.Top - tab_batch.Top - 100
'
'   Reset picture variables to not sized and then call sub to resize current selection.
'
    mod_pic_moved(1) = "N"
    mod_pic_moved(2) = "N"
    mod_pic_moved(3) = "N"
    mod_pic_moved(4) = "N"
    tab_select current_tab
'
'********* Code Ends Here **************
'
   Exit Sub
'
Form_Resize_Error:
    DisplayError , "frm_photograph_view.Form_Resize", vbExclamation
End Sub

Public Sub load_batch(Batch_Number As Long)
'***************************************
' Module/Form Name   : frm_photograph_view
'
' Procedure Name     : load_batch
'
' Purpose            :
'
' Date Created       : 13/12/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : Batch_Number - Long
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

    mod_images_displayed = False

    Set moDigitalImage = New DigitalImage
    moDigitalImage.BatchNo = Batch_Number
    '
    Set snap_batch = photograph.Read(Batch_Number, dbOpenSnapshot)

    If snap_batch.RecordCount = 0 Then
        MsgBox "Photograph not found", vbExclamation, "View Photograph"
        Unload Me
        Set frm_photograph_view = Nothing
        Exit Sub
    End If
    '
    '   Get the batch object.
    '
    Set moBatch = New Batch2
    moBatch.Read Batch_Number
    '
    Me.Caption = "View Photograph: " & CStr(Batch_Number)
    Me.Show
    DoEvents
    display_fields
'
'   If already loaded and images are currently being shown, redisplay.
'
    If tab_batch.SelectedItem.Index = 3 Then
        Screen.MousePointer = vbHourglass
        DisplayTransparencies
        Screen.MousePointer = vbDefault
        mod_images_displayed = True
    End If
    Me.SetFocus
'
'********* Code Ends Here **************
'
   Exit Sub
'
load_batch_Error:
    ErrorRaise "frm_photograph_view.load_batch"
End Sub

Public Sub RedisplayTransparencies()
'***************************************
' Module/Form Name   : frm_photograph_view
'
' Procedure Name     : RedisplayTransparencies
'
' Purpose            :
'
' Date Created       : 19/05/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo RedisplayTransparencies_Error
'
'******** Code Starts Here *************
'
'
'   First refresh the data.
'
    datPhotograph.Refresh
    DisplayTransparencies
'
'   Then find the current image.
'
    If chk_archived_images Then
        If current_image_no(2) <> 0 Then
                smgPhotograph.GetKeyRow current_image_no(2)
        End If
    Else
        If current_image_no(1) <> 0 Then
            smgPhotograph.GetKeyRow current_image_no(1)
        End If
    End If
'
'********* Code Ends Here **************
'
   Exit Sub
'
RedisplayTransparencies_Error:
    ErrorRaise "frm_photograph_view.RedisplayTransparencies"
End Sub

Public Sub RedisplayDigitalImages()
    Dim intCurrentRow   As Integer
    Dim lngDNoteNo      As Long
    
'***************************************
' Module/Form Name   : frm_photograph_view
'
' Procedure Name     : RedisplayDigitalImages
'
' Purpose            :
'
' Date Created       : 14/06/2006 00:18
'
' Author             : Gareth
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
    On Error GoTo ErrorProc

    intCurrentRow = smgDigitalImages.CurrentRow
    lngDNoteNo = smgDigitalImages.Column(1).Value
    '
    '   First refresh the data.
    '
    moBatch.DigitalImages.Refresh
    DisplayDigitalImages
    '
    '   Then find the current image.
    '
    smgDigitalImages.GetKeyRow CStr(lngDNoteNo)

    On Error GoTo 0
    Exit Sub
'
'********* Code Ends Here **************
'
ErrorProc:
    ErrorRaise "frm_photograph_view.RedisplayDigitalImages"
End Sub

Public Sub RedisplayFields()
'***************************************
' Module/Form Name   : frm_photograph_view
'
' Procedure Name     : RedisplayFields
'
' Purpose            :
'
' Date Created       : 13/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo RedisplayFields_Error
'
'******** Code Starts Here *************
'
    snap_batch.Requery
    display_fields
'
'********* Code Ends Here **************
'
   Exit Sub
'
RedisplayFields_Error:
    ErrorRaise "frm_photograph_view.RedisplayFields"
End Sub

Private Sub SizePhotograph()
'***************************************
' Module/Form Name   : frm_photograph_view
'
' Procedure Name     : SizePhotograph
'
' Purpose            :
'
' Date Created       : 13/12/2002
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
    Dim oDigitalImage As DigitalImage

    Set oDigitalImage = New DigitalImage
    With oDigitalImage
        .MaxHeight = IIf(fraImageFile.Height - 300 > 32767, 32767, fraImageFile.Height - 300)
        .MaxWidth = IIf(fraImageFile.Width - 300 > 32767, 32767, fraImageFile.Width - 300)
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
    ErrorRaise "frm_photograph_view.SizePhotograph"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set moBatch = Nothing
    On Error Resume Next
    gcolMaxedWindows.Remove CStr(Me.hWnd)
End Sub

Private Sub smgDigitalImages_RowChanged(CurrentRow As String)
    EnableButtons
End Sub

Private Sub smgPhotograph_Click()
'***************************************
' Module/Form Name   : frm_photograph_view
'
' Procedure Name     : smgPhotograph_Click
'
' Purpose            :
'
' Date Created       : 13/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo smgPhotograph_Click_Error
'
'******** Code Starts Here *************
'

    EnableButtons
'
'********* Code Ends Here **************
'
   Exit Sub
'
smgPhotograph_Click_Error:
    DisplayError , "frm_photograph_view.smgPhotograph_Click", vbExclamation
End Sub

Private Sub tab_batch_Click()
'***************************************
' Module/Form Name   : frm_photograph_view
'
' Procedure Name     : tab_batch_Click
'
' Purpose            :
'
' Date Created       : 13/12/2002
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
            DisplayTransparencies
            Screen.MousePointer = vbDefault
            mod_images_displayed = True
        End If
    ElseIf tab_batch.SelectedItem.Index = 4 Then
        Screen.MousePointer = vbHourglass
        DisplayDigitalImages
        Screen.MousePointer = vbDefault
    End If

    current_tab = tab_batch.SelectedItem.Index
    tab_select tab_batch.SelectedItem.Index
'
'********* Code Ends Here **************
'
   Exit Sub
'
tab_batch_Click_Error:
    DisplayError , "frm_photograph_view.tab_batch_Click", vbExclamation
End Sub

Private Sub tab_select(selected_item As Integer)
'***************************************
' Module/Form Name   : frm_photograph_view
'
' Procedure Name     : tab_select
'
' Purpose            :
'
' Date Created       : 13/12/2002
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
            '   Then resize the photograph.
            '
            If imgPhotograph.Picture <> 0 Then
                SizePhotograph
            End If
            '
            fraImageFile.Visible = True
        Case Is = 2
            txt_copyright.Width = pic_batch(1).Width - txt_copyright.Left - 200
            txt_credit.Width = pic_batch(1).Width - txt_credit.Left - 200
            txtKeywords.Width = pic_batch(1).Width - txtKeywords.Left - 200
        Case Is = 3
            '
            '   Reposition Image manipulation buttons.
            '
            com_delivery_note.Top = tab_batch.ClientHeight - com_delivery_note.Height - 100
            com_delivery_note.Left = tab_batch.ClientWidth - (com_delivery_note.Width + _
                                                              com_job.Width + 200)

            com_job.Top = com_delivery_note.Top
            com_job.Left = com_delivery_note.Left + com_delivery_note.Width + 100
            '
            '   Resize the grid.
            '
            With smgPhotograph
                .Move 100, txt_batch_no(0).Top + txt_batch_no(0).Height + 200, _
                      pic_batch(2).Width - .Left - 150, _
                      com_delivery_note.Top - .Top
                .Column(1).Width = 1000
                .Column(2).Width = 1000
                .Column(3).Width = 1000
                .Column(6).Width = 800
                .Column(7).Width = 0
                .Column(8).Width = 0
                .Column(9).Width = 0
                .Column(4).Width = (.Width - goSystemConfig.VScrollBarWidth - 100 - _
                                             .Column(1).Width - _
                                             .Column(2).Width - _
                                             .Column(3).Width - _
                                             .Column(6).Width) / 2

                .Column(5).Width = .Column(4).Width
            End With
        Case Is = 4
            '
            '   Reposition Image manipulation buttons.
            '
            cmdDigitalReturn.Top = tab_batch.ClientHeight - cmdDigitalReturn.Height - 100
            cmdDigitalReturn.Left = tab_batch.ClientWidth - (cmdDigitalReturn.Width + 100)
            '
            '   Resize the grid.
            '
            With smgDigitalImages
                .Move 100, txt_batch_no(0).Top + txt_batch_no(0).Height + 200, _
                      pic_batch(3).Width - .Left - 150, _
                      cmdDigitalReturn.Top - .Top
                .Column(1).Width = 1000
                .Column(2).Width = 1000
                .Column(3).Width = (.Width - goSystemConfig.VScrollBarWidth - 100 - _
                                             .Column(1).Width - _
                                             .Column(2).Width) / 2

                .Column(4).Width = .Column(3).Width
                .Column(5).Width = 0
            End With
    End Select
    For Each tab_frame In pic_batch
        tab_frame.Enabled = False
        tab_frame.Visible = False
    Next
    pic_batch(selected_item - 1).ZOrder 0
    pic_batch(selected_item - 1).Enabled = True
    pic_batch(selected_item - 1).Visible = True
'
'********* Code Ends Here **************
'
   Exit Sub
'
tab_select_Error:
    ErrorRaise "frm_photograph_view.tab_select"
End Sub

Public Sub ForceResize()
    gblnResizeMaxedWindows = False
    Form_Resize
    gblnResizeMaxedWindows = True
End Sub


