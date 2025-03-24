VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_photograph_new 
   Caption         =   "New Photograph"
   ClientHeight    =   6645
   ClientLeft      =   2415
   ClientTop       =   2295
   ClientWidth     =   7905
   Icon            =   "Photograph New.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6645
   ScaleWidth      =   7905
   Begin VB.PictureBox pic_batch 
      Height          =   4695
      Index           =   0
      Left            =   180
      ScaleHeight     =   4635
      ScaleWidth      =   7455
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   7515
      Begin VB.TextBox txtTiffLocation 
         DataField       =   "customername"
         DataSource      =   "dat_customer"
         Height          =   315
         Left            =   1600
         MaxLength       =   50
         TabIndex        =   24
         Top             =   3780
         Width           =   1875
      End
      Begin VB.CheckBox chkExclusive 
         Caption         =   "Check1"
         Height          =   195
         Left            =   1600
         TabIndex        =   23
         Top             =   4260
         Width           =   255
      End
      Begin VB.Frame fraImageFile 
         Caption         =   "Digital Image"
         Height          =   2055
         Left            =   3600
         TabIndex        =   12
         Top             =   1800
         Width           =   3015
         Begin VB.CommandButton cmdSetPicture 
            Caption         =   "&Set..."
            Height          =   495
            Left            =   2040
            TabIndex        =   13
            Top             =   240
            Width           =   975
         End
         Begin VB.Image imgPhotograph 
            Height          =   1740
            Left            =   120
            Stretch         =   -1  'True
            Top             =   240
            Width           =   1860
         End
      End
      Begin VB.ComboBox cmb_year 
         Height          =   315
         Left            =   2580
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1860
         Width           =   735
      End
      Begin VB.ComboBox cmb_month 
         Height          =   315
         Left            =   1600
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1860
         Width           =   735
      End
      Begin VB.TextBox txt_description 
         DataField       =   "customername"
         DataSource      =   "dat_customer"
         Height          =   1275
         Left            =   1600
         MaxLength       =   180
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   4095
      End
      Begin VB.TextBox txt_no_of_photographs 
         DataField       =   "customername"
         DataSource      =   "dat_customer"
         Height          =   315
         Left            =   1600
         MaxLength       =   6
         TabIndex        =   9
         Top             =   2820
         Width           =   675
      End
      Begin VB.ComboBox cmb_format 
         Height          =   315
         Left            =   1600
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   3300
         Width           =   855
      End
      Begin VB.TextBox txt_batch 
         DataField       =   "customername"
         DataSource      =   "dat_customer"
         Height          =   315
         Left            =   1600
         MaxLength       =   6
         TabIndex        =   7
         Top             =   2340
         Width           =   795
      End
      Begin VB.Label lblTiffLocation 
         Caption         =   "Tiff Location:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   60
         TabIndex        =   26
         Top             =   3780
         Width           =   1455
      End
      Begin VB.Label lblExclusive 
         Caption         =   "Exclusive:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   60
         TabIndex        =   25
         Top             =   4260
         Width           =   975
      End
      Begin VB.Label lab_description_lab 
         Caption         =   "Description:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   60
         TabIndex        =   1
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label lab_batch_no 
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1920
         TabIndex        =   21
         Top             =   60
         Width           =   735
      End
      Begin VB.Label lab_date_taken_lab 
         Caption         =   "Date Taken:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   60
         TabIndex        =   3
         Top             =   1920
         Width           =   1035
      End
      Begin VB.Label lab_parent_batch_lab 
         Caption         =   "Batch:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   60
         TabIndex        =   6
         Top             =   2340
         Width           =   1035
      End
      Begin VB.Label lab_no_of_photographs_lab 
         Caption         =   "No. of photographs:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   60
         TabIndex        =   8
         Top             =   2820
         Width           =   1455
      End
      Begin VB.Label lab_format_lab 
         Caption         =   "Format:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   60
         TabIndex        =   10
         Top             =   3300
         Width           =   1455
      End
   End
   Begin VB.PictureBox pic_batch 
      Height          =   3975
      Index           =   1
      Left            =   240
      ScaleHeight     =   3915
      ScaleWidth      =   7455
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   540
      Width           =   7515
      Begin VB.TextBox txt_credit 
         DataField       =   "customername"
         DataSource      =   "dat_customer"
         Height          =   675
         Left            =   1980
         MaxLength       =   45
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   1560
         Width           =   4095
      End
      Begin VB.TextBox txt_copyright 
         DataField       =   "customername"
         DataSource      =   "dat_customer"
         Height          =   615
         Left            =   1980
         MaxLength       =   30
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   360
         Width           =   4095
      End
      Begin VB.Label lab_copyright_lab 
         Caption         =   "Copyright:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label lab_credit_lab 
         Caption         =   "Credit:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   1560
         Width           =   1155
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   555
      Left            =   4440
      Picture         =   "Photograph New.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5400
      UseMaskColor    =   -1  'True
      Width           =   795
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   555
      Left            =   2040
      Picture         =   "Photograph New.frx":089C
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5400
      UseMaskColor    =   -1  'True
      Width           =   795
   End
   Begin MSComctlLib.TabStrip tab_batch 
      Height          =   5235
      Left            =   60
      TabIndex        =   20
      Top             =   60
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   9234
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Credit/Copyright"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgPhotograph 
      Left            =   6120
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frm_photograph_new"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mod_pic_moved(2) As String

Private snap_photograph As DAO.Recordset
Private table_batch As DAO.Recordset
Private table_batchno As DAO.Recordset
'
'   Used for storing the photograph number once it has been created.
'
Private mod_photograph_no           As Long
'
Private current_tab                 As Integer
Private mbolDontResize              As Boolean
Private moDigitalImage              As DigitalImage
Private oFSO                        As New Scripting.FileSystemObject

Private Sub tab_select(selected_item As Integer)
'***************************************
' Module/Form Name   : frm_photograph_new
'
' Procedure Name     : tab_select
'
' Purpose            :
'
' Date Created       : 12/12/2002
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
    Dim oCurrentControl As Control

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
            '   Then move the buttons.
            '
            cmdSetPicture.Top = 200
            cmdSetPicture.Left = fraImageFile.Width - 200 - cmdSetPicture.Width
            'cmdRemovePicture.Top = cmdSetPicture.Top + cmdSetPicture.Height + 100
            'cmdRemovePicture.Left = cmdSetPicture.Left
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
    End Select
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
        Else
            txt_copyright.SetFocus
        End If
    End If
'
'********* Code Ends Here **************
'
   Exit Sub
'
tab_select_Error:
    ErrorRaise "frm_photograph_new.tab_select"
End Sub

Private Sub determine_carry_over()
    
10        msg = "Photograph " + Str(mod_photograph_no) + " created" + vbCr + _
                "Do you wish to carry the details over?"
20        style = vbYesNo + vbQuestion + vbDefaultButton1
30        response = MsgBox(msg, style, "New Photograph")

40        If response = vbNo Then
50            txt_description = ""
60            photograph.set_up_date_controls Now, cmb_month, cmb_year
70            txt_batch = ""
80            txt_no_of_photographs = ""
90            cmb_format.ListIndex = 0
100           txt_copyright = "CLIVE NICHOLS"
110           txt_credit = "CLIVE NICHOLS"
120       End If
130       Set moDigitalImage = Nothing
140       Set moDigitalImage = New DigitalImage
150       imgPhotograph.Picture = LoadPicture("")
160       tab_batch.Tabs(1).Selected = True
170       txt_description.SetFocus
    
End Sub

Private Sub close_batch_tables()
10        On Error Resume Next
    
20        table_batch.Close
30        table_batchno.Close
    
End Sub


Private Sub display_fields()
'***************************************
' Module/Form Name   : frm_photograph_new
'
' Procedure Name     : display_fields
'
' Purpose            :
'
' Date Created       : 12/12/2002
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
    '
    '   Set up the date combo boxes.
    '
    photograph.set_up_date_controls Now, cmb_month, cmb_year
    '
    '   Set up formats in the combo box.
    '
    cmb_format.Clear
    '
    cmb_format.AddItem "Digital"
    cmb_format.AddItem "35"
    cmb_format.AddItem "54"
    cmb_format.AddItem "66"
    cmb_format.AddItem "67"
    cmb_format.AddItem "617"
    cmb_format.AddItem "645"

    cmb_format.ListIndex = 0

    txt_copyright = "CLIVE NICHOLS"
    txt_credit = "CLIVE NICHOLS"

'
'********* Code Ends Here **************
'
   Exit Sub
'
display_fields_Error:
    ErrorRaise "frm_photograph_new.display_fields"
End Sub

Private Sub create_images()
'***************************************
' Module/Form Name   : frm_photograph_new
'
' Procedure Name     : create_images
'
' Purpose            :
'
' Date Created       : 12/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo create_images_Error
'
'******** Code Starts Here *************
'
    Dim i As Integer
    Dim table_photograph As DAO.Recordset

    Set table_photograph = db.OpenRecordset("Photograph", dbOpenTable)

    For i = 1 To txt_no_of_photographs.Text
        With table_photograph
            .AddNew
            !batch_no = table_batchno!next_no
            !photograph_no = i
            !Status = "I"
            .update
        End With
    Next i

'
'********* Code Ends Here **************
'
   Exit Sub
'
create_images_Error:
    ErrorRaise "frm_photograph_new.create_images"
End Sub

Private Sub new_batch()
'***************************************
' Module/Form Name   : frm_photograph_new
'
' Procedure Name     : new_batch
'
' Purpose            :
'
' Date Created       : 12/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo new_batch_Error
'
'******** Code Starts Here *************
'
    Dim oBatch      As Batch2
    
    msg_title = "New Photograph"

    begin_trans
        Set table_batch = db.OpenRecordset("Batch", dbOpenTable)
        Set table_batchno = db.OpenRecordset("Batchno", dbOpenTable)

        On Error Resume Next
        table_batchno.MoveFirst
        If Err.Number = 3021 Then
            MsgBox "Photograph number table has not been set up correctly", vbCritical, msg_title
            close_batch_tables
            roll_back
            Exit Sub
        End If

        On Error GoTo new_batch_Error
        table_batchno.Edit

        mod_photograph_no = table_batchno!next_no

        With table_batch
            .AddNew
            !batch_no = mod_photograph_no
            !Description = txt_description.Text
            !date_taken = CDate(cmb_month.Text + " 01 " + cmb_year.Text)

            If Val(txt_batch) = 0 Then
                !parent_batch = mod_photograph_no
            Else
                !parent_batch = Val(txt_batch)
            End If
            !Format = cmb_format.Text
            !TiffLocation = Trim(txtTiffLocation.Text)
            !Exclusive = IIf(chkExclusive.Value = vbChecked, True, False)
            !CopyRight = txt_copyright.Text
            !Credit = txt_credit.Text
            '
            '   Process the digital image.
            '
            With moDigitalImage
                .BatchNo = mod_photograph_no
                On Error Resume Next
                .Associate
                If Err.Number - vbObjectError = 1 Then
                    DisplayError
                    roll_back
                    close_batch_tables
                    Exit Sub
                ElseIf Err.Number - vbObjectError = 2 Then
                    If MsgBox(Err.Description & vbCrLf _
                            & "Do you wish to overwrite the existing image with the one that you have chosen?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                        roll_back
                        close_batch_tables
                        Exit Sub
                    Else
                        .Associate True
                    End If
                ElseIf Err.Number <> 0 Then
                    ErrorSave
                    On Error GoTo new_batch_Error
                    ErrorRestore
                End If
            End With
            .update
            '
            '   Create the Keywords.
            '
            Set oBatch = New Batch2
            oBatch.Read mod_photograph_no
            oBatch.CreateAutomaticKeywords
            Set oBatch = Nothing
        End With
        create_images
        '
        '   Update the next photograph number to be used.
        '
        With table_batchno
            !next_no = !next_no + 1
            .update
        End With

        close_batch_tables
        '
        '    Move Image from its current location
        '    to the correct location in the hierarchy.
        '

    commit_trans
    '
    '   Display the last photograph number used.
    '
    mdi_npls.display_last_photo_number_used
    '
    determine_carry_over
    '
    If is_form_loaded("frm_photograph_maint") Then
        With frm_photograph_maint
            .txt_photograph_no = mod_photograph_no
            HighLightText .txt_photograph_no
        End With
    End If
'
'********* Code Ends Here **************
'
   Exit Sub
'
new_batch_Error:
    Select Case Err.Number
        Case Is = 3262
            MsgBox "'Next Customer Number' table is currently in use." + vbCr + "Please try again", vbInformation
        Case Is = 3260
            MsgBox "Photograph table is currently busy, please try again", vbInformation
        Case 3420
            Resume Next
        Case Else
            ErrorRaise "frm_photograph_new.new_batch"
    End Select
    roll_back
    close_batch_tables
End Sub

Private Function valid_input() As Boolean
'***************************************
' Module/Form Name   : frm_photograph_new
'
' Procedure Name     : valid_input
'
' Purpose            :
'
' Date Created       : 12/12/2002
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

'    If Not IsDate(msk_date_taken) Then
'        tab_batch.Tabs(1).Selected = True
'        MsgBox "Please enter a valid date", vbExclamation
'        msk_date_taken.SetFocus
'        Exit Function
'    End If

    If Len(RTrim(txt_no_of_photographs.Text)) = 0 Then
        tab_batch.Tabs(1).Selected = True
        MsgBox "Please enter the number of images for this photograph", vbExclamation
        txt_no_of_photographs.SetFocus
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
                Set snap_photograph = photograph.Read(txt_batch, dbOpenSnapshot)
                If snap_photograph.RecordCount = 0 Then
                    MsgBox "The Photograph to be used for grouping the Batch does not exist", vbExclamation
                    txt_batch.SetFocus
                    Exit Function
                End If
        End Select
    End If

    valid_input = True

'
'********* Code Ends Here **************
'
   Exit Function
'
valid_input_Error:
    ErrorRaise "frm_photograph_new.valid_input"
End Function

Private Sub cmdSetPicture_Click()
'***************************************
' Module/Form Name   : frm_photograph_new
'
' Procedure Name     : cmdSetPicture_Click
'
' Purpose            :
'
' Date Created       : 12/12/2002
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
    '
    '   Open a dialog box at the Holding Area Location.
    '
    With dlgPhotograph
        .CancelError = True
        .DialogTitle = "Select Digital Image"
        .Filter = "Picture Files (*.jpg)|*.jpg"
        If moDigitalImage.NewFileLocation = "" Then
            '.InitDir = oFSO.BuildPath(App.Path, goSystemConfig.ImageHoldingArea)
            .InitDir = oFSO.BuildPath(goSystemConfig.ServerLocation, goSystemConfig.ImageHoldingArea)
        Else
            .InitDir = moDigitalImage.NewFileLocation
        End If
        If moDigitalImage.NewFileLocation = "" Then
            .FileName = ""
        Else
            .FileName = StripPath(moDigitalImage.NewFileLocation)
        End If
        .ShowOpen
    End With
    '
    '   Show Image.
    '
    Screen.MousePointer = vbHourglass
    DoEvents
    imgPhotograph.Picture = LoadPicture(dlgPhotograph.FileName)
    moDigitalImage.NewFileLocation = dlgPhotograph.FileName
    SizePhotograph
    txt_description.SetFocus
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
        DisplayError , "frm_photograph_new.cmdSetPicture_Click", vbExclamation
    End If
End Sub

Private Sub Form_Load()
'***************************************
' Module/Form Name   : frm_photograph_new
'
' Procedure Name     : Form_Load
'
' Purpose            :
'
' Date Created       : 12/12/2002
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
    mod_pic_moved(1) = "N"
    mod_pic_moved(2) = "N"

    mbolDontResize = True
    tab_select 1
'
'   Set the width & height of the form.
'
    current_tab = 1
    Me.WindowState = GetSetting(App.Title, "PhotographNew", "WindowState", 0)
    If GetSetting(App.Title, "PhotographNew", "WindowState", 0) <> 2 Then
        Me.Width = GetSetting(App.Title, "PhotographNew", "Width", 8025)
        Me.Height = GetSetting(App.Title, "PhotographNew", "Height", 6000)
    End If

    com_position_form Me
    mbolDontResize = False
    Set moDigitalImage = New DigitalImage
    Me.Show
    display_fields
    com_position_form Me
'
'********* Code Ends Here **************
'
   Exit Sub
'
Form_Load_Error:
    DisplayError , "frm_photograph_new.Form_Load", vbExclamation
End Sub

Private Sub Form_Resize()
'***************************************
' Module/Form Name   : frm_photograph_new
'
' Procedure Name     : Form_Resize
'
' Purpose            :
'
' Date Created       : 12/12/2002
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
    MIN_HEIGHT = 6300
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

    LockWindow Me.hWnd
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
    cmdOK.Top = Height - cmdOK.Height - header_height - 100
    cmdCancel.Top = cmdOK.Top
    cmdOK.Left = (Width - (cmdOK.Width + cmdCancel.Width + 500)) / 2
    cmdCancel.Left = cmdOK.Left + cmdOK.Width + 500
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
    tab_select current_tab

    UnlockWindow
'
'********* Code Ends Here **************
'
   Exit Sub
'
Form_Resize_Error:
    DisplayError , "frm_photograph_new.Form_Resize", vbExclamation
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting App.Title, "PhotographNew", "Width", Me.Width
    SaveSetting App.Title, "PhotographNew", "Height", Me.Height
    SaveSetting App.Title, "PhotographNew", "WindowState", Me.WindowState
    On Error Resume Next
    gcolMaxedWindows.Remove CStr(Me.hWnd)
End Sub

Private Sub cmdCancel_Click()
'***************************************
' Module/Form Name   : frm_photograph_new
'
' Procedure Name     : cmdCancel_Click
'
' Purpose            :
'
' Date Created       : 12/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo cmdCancel_Click_Error
'
'******** Code Starts Here *************
'
    If is_form_loaded("frm_photograph_find_maint") Then
        frm_photograph_find_maint.refresh_list mod_photograph_no
    End If
    Unload Me
    Set frm_photograph_new = Nothing
'
'********* Code Ends Here **************
'
   Exit Sub
'
cmdCancel_Click_Error:
    DisplayError , "frm_photograph_new.cmdCancel_Click", vbExclamation
End Sub

Private Sub cmdOK_Click()
'***************************************
' Module/Form Name   : frm_photograph_new
'
' Procedure Name     : cmdOK_Click
'
' Purpose            :
'
' Date Created       : 12/12/2002
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
    new_batch
'
'********* Code Ends Here **************
'
   Exit Sub
'
cmdOK_Click_Error:
    DisplayError , "frm_photograph_new.cmdOK_Click", vbExclamation
End Sub

Private Sub tab_batch_Click()
'***************************************
' Module/Form Name   : frm_photograph_new
'
' Procedure Name     : tab_batch_Click
'
' Purpose            :
'
' Date Created       : 12/12/2002
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
    current_tab = tab_batch.SelectedItem.Index
    tab_select tab_batch.SelectedItem.Index
'
'********* Code Ends Here **************
'
   Exit Sub
'
tab_batch_Click_Error:
    DisplayError , "frm_photograph_new.tab_batch_Click", vbExclamation
End Sub

Private Sub txt_batch_KeyPress(KeyAscii As Integer)
    KeyAscii = allow_numeric_only(KeyAscii)
End Sub

Private Sub txt_description_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_no_of_photographs_KeyPress(KeyAscii As Integer)
    KeyAscii = allow_numeric_only(KeyAscii)
End Sub

Private Sub SizePhotograph()
'***************************************
' Module/Form Name   : frm_photograph_new
'
' Procedure Name     : SizePhotograph
'
' Purpose            :
'
' Date Created       : 12/12/2002
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
        .MaxHeight = fraImageFile.Height - 300
        .MaxWidth = cmdSetPicture.Left - 200
        Set .Picture = imgPhotograph.Picture
        imgPhotograph.Top = 200
        imgPhotograph.Left = 100
        imgPhotograph.Width = .Width
        imgPhotograph.Height = .Height
    End With
'
'********* Code Ends Here **************
'
   Exit Sub
'
SizePhotograph_Error:
    ErrorRaise "frm_photograph_new.SizePhotograph"
End Sub

Public Sub ForceResize()
    gblnResizeMaxedWindows = False
    Form_Resize
    gblnResizeMaxedWindows = True
End Sub


