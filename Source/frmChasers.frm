VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmChasers 
   Caption         =   "Chase Outstanding Pictures"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7380
   Icon            =   "frmChasers.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6870
   ScaleWidth      =   7380
   Begin ImageWhere.SimpleGrid smgChases 
      Height          =   3495
      Left            =   60
      TabIndex        =   11
      Top             =   1620
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   6165
      Columns         =   1
      KeyCol          =   0
   End
   Begin VB.TextBox txtReturnByDate 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   5580
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtCustomer 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1740
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   60
      Width           =   4935
   End
   Begin Crystal.CrystalReport crpChaser 
      Left            =   600
      Top             =   5580
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   375
      Left            =   4380
      TabIndex        =   9
      Top             =   5760
      Width           =   975
   End
   Begin VB.TextBox txtJobDescription 
      BackColor       =   &H8000000F&
      Height          =   555
      Left            =   1740
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   900
      Width           =   4935
   End
   Begin VB.TextBox txtDeliveryNote 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1740
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   480
      Width           =   795
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   5580
      TabIndex        =   10
      Top             =   5760
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label lblReturnByDate 
      Caption         =   "Original Return By Date:"
      Height          =   255
      Left            =   3420
      TabIndex        =   5
      Top             =   540
      Width           =   1935
   End
   Begin VB.Label lblCustomer 
      Caption         =   "Customer:"
      Height          =   255
      Left            =   180
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblJobDescription 
      Caption         =   "Job Description:"
      Height          =   255
      Left            =   180
      TabIndex        =   6
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label lblDeliveryNote 
      Caption         =   "Delivery Note:"
      Height          =   255
      Left            =   180
      TabIndex        =   4
      Top             =   540
      Width           =   1215
   End
End
Attribute VB_Name = "frmChasers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private moChasers               As Chasers
Private moJob                   As Job2
Private moCustomer              As Customer2
Private WithEvents fChaser      As frmChaser
Attribute fChaser.VB_VarHelpID = -1
Private mbolDontResize          As Boolean
Private mblnOutstandingPictures As Boolean
Private mblnMultiplSelection    As Boolean

Private Sub cmdAdd_Click()
'***************************************
' Module/Form Name   : frmChasers
'
' Procedure Name     : cmdAdd_Click
'
' Purpose            :
'
' Date Created       : 29/12/2002
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
    AddChaser
'
'********* Code Ends Here **************
'
   Exit Sub
'
cmdAdd_Click_Error:
    DisplayError , "frmChasers.cmdAdd_Click", vbExclamation
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub update_loaded_screens()
    If is_form_loaded("frm_delivery_note_maint") Then
        frm_delivery_note_maint.RefreshDeliveryNotes pblnPicturesReturned:=True
    End If
End Sub

Private Sub cmdPrint_Click()
'***************************************
' Module/Form Name   : frmChasers
'
' Procedure Name     : cmdPrint_Click
'
' Purpose            :
'
' Date Created       : 29/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo cmdPrint_Click_Error
'
'******** Code Starts Here *************
'
    
    PrintChaser
    
'
'********* Code Ends Here **************
'
   Exit Sub
'
cmdPrint_Click_Error:
    DisplayError , "frmChasers.cmdPrint_Click", vbExclamation
End Sub

Private Sub fChaser_ChaserAdded(ID As Long)
'***************************************
' Module/Form Name   : frmChasers
'
' Procedure Name     : fChaser_ChaserAdded
'
' Purpose            :
'
' Date Created       : 29/12/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : ID - Integer
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo fChaser_ChaserAdded_Error
'
'******** Code Starts Here *************
'
    moChasers.Refresh
    moChasers.CurrentId = ID
    RefreshChasers
''    smgChases.GetKeyRow CStr(ID)
    CheckPrintAvailability
    update_loaded_screens
'
'********* Code Ends Here **************
'
   Exit Sub
'
fChaser_ChaserAdded_Error:
    DisplayError , "frmChasers.fChaser_ChaserAdded", vbExclamation
End Sub

Private Sub Form_Load()
'***************************************
' Module/Form Name   : frmChasers
'
' Procedure Name     : Form_Load
'
' Purpose            :
'
' Date Created       : 29/12/2002
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
    mbolDontResize = True
    Me.WindowState = GetSetting(App.Title, "Chasers", "WindowState", 0)
    If GetSetting(App.Title, "Chasers", "WindowState", 0) <> 2 Then
        Me.Width = GetSetting(App.Title, "Chasers", "Width", 7140)
        Me.Height = GetSetting(App.Title, "Chasers", "Height", 5415)
    End If
    mbolDontResize = False
    '
    With smgChases
        .Columns = 8
        .Column(1).Header = "No"
        .Column(2).Header = "Date Actioned"
        .Column(2).Align = flexAlignCenterTop
        .Column(3).Header = "Contact"
        .Column(3).Align = flexAlignLeftTop
        .Column(4).Header = "Action"
        .Column(4).Align = flexAlignLeftTop
        .Column(5).Header = "User"
        .Column(5).Align = flexAlignLeftTop
        .Column(6).Header = "Chaser Date"
        .Column(6).Align = flexAlignCenterTop
        .Column(7).Header = "Next Action"
        .Column(7).Align = flexAlignLeftTop
        .Column(8).Header = "Comments"
        .Column(8).Align = flexAlignLeftTop
        .KeyCol = 1
    End With
'
'********* Code Ends Here **************
'
   Exit Sub
'
Form_Load_Error:
    DisplayError , "frmChasers.Form_Load", vbExclamation
End Sub

Private Sub Form_Resize()
'***************************************
' Module/Form Name   : frmChasers
'
' Procedure Name     : Form_Resize
'
' Purpose            :
'
' Date Created       : 29/12/2002
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

    MIN_WIDTH = 10950
    MIN_HEIGHT = 5415
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
    txtCustomer.Width = Me.Width - txtCustomer.Left - 300
    txtJobDescription.Width = Me.Width - txtJobDescription.Left - 300
    txtReturnByDate.Left = Me.Width - txtReturnByDate.Width - 300
    lblReturnByDate.Left = txtReturnByDate.Left - lblReturnByDate.Width - 300
    cmdClose.Top = Me.Height - goSystemConfig.TitleBarHeight - cmdClose.Height - 200
    cmdClose.Left = Me.Width - cmdClose.Width - 200
    cmdPrint.Top = cmdClose.Top
    cmdPrint.Left = cmdClose.Left - cmdClose.Width - 200
    cmdAdd.Top = cmdPrint.Top
    cmdAdd.Left = cmdPrint.Left - cmdPrint.Width - 200
    With smgChases
        .Height = cmdClose.Top - .Top - 200
        .Width = Me.Width - .Left * 2 - 200
        .Column(1).Width = 0
        .Column(2).Width = 1200
        .Column(3).Width = 2000
        .Column(4).Width = 1200
        .Column(5).Width = .Column(4).Width
        .Column(6).Width = 1200
        .Column(7).Width = .Column(4).Width
        .Column(8).Width = (.Width - _
                            .Column(1).Width - _
                            .Column(2).Width - _
                            .Column(3).Width - _
                            .Column(4).Width - _
                            .Column(5).Width - _
                            .Column(6).Width - _
                            .Column(7).Width - _
                            goSystemConfig.VScrollBarWidth - 100)

        .ResizeRows
    End With
'
'********* Code Ends Here **************
'
   Exit Sub
'
Form_Resize_Error:
    DisplayError , "frmChasers.Form_Resize", vbExclamation
End Sub

Private Sub RefreshChasers()
'***************************************
' Module/Form Name   : frmChasers
'
' Procedure Name     : RefreshChases
'
' Purpose            :
'
' Date Created       : 15/02/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 15/02/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo RefreshChasers_Error
'
'******** Code Starts Here *************
'
    Dim oChaser As Chaser

    With smgChases
        LockWindow Me.hWnd
        .Redraw = False
        .Clear
        For Each oChaser In moChasers
            If oChaser.Action <> "None" Or oChaser.NextAction <> "None" Then
                .AddRow False, _
                        oChaser.ID, _
                        Format(oChaser.DateCreated, "dd/mm/yyyy"), _
                        oChaser.Contact, _
                        oChaser.Action, _
                        oChaser.User, _
                        IIf(oChaser.Action = "Invoiced", "", Format(oChaser.ChaserDate, "dd/mm/yyyy")), _
                        oChaser.NextAction, _
                        oChaser.Comment
            End If
        Next oChaser
        .ResizeRows
        .Column(1).Sorted = smgAscending
        .GetKeyRow (moChasers.CurrentId)
        .Redraw = True
        UnlockWindow
    End With
    CheckPrintAvailability
'
'********* Code Ends Here **************
'
    Exit Sub
    '
RefreshChasers_Error:
    DisplayError , "frmChasers.RefreshChasers", vbExclamation
    smgChases.Redraw = True
End Sub

Public Sub Display(ByVal DeliveryNoteNo As Long, _
                   ByVal pblnOutstandingPictures As Boolean)
'***************************************
' Module/Form Name   : frmChasers
'
' Procedure Name     : Display
'
' Purpose            :
'
' Date Created       : 29/12/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : DeliveryNoteNo - Long
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
    mblnOutstandingPictures = pblnOutstandingPictures
    Set moChasers = New Chasers
    moChasers.DeliveryNoteNo = DeliveryNoteNo
    RefreshChasers
    Set moJob = New Job2
    moJob.Read , DeliveryNoteNo
    Set moCustomer = New Customer2
    moCustomer.Read moJob.CustomerNo
    '
    txtCustomer.Text = moCustomer.CustomerName
    txtDeliveryNote.Text = DeliveryNoteNo
    txtJobDescription.Text = moJob.reference
    txtReturnByDate.Text = Format(moChasers.OriginalReturnByDate, "dd/mm/yyyy")
    LockWindow Me.hWnd
    Me.Show
    UnlockWindow
    smgChases.SetFocus
'
'********* Code Ends Here **************
'
   Exit Sub
'
Display_Error:
    ErrorRaise "frmChasers.Display"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting App.Title, "Chasers", "Width", Me.Width
    SaveSetting App.Title, "Chasers", "Height", Me.Height
    SaveSetting App.Title, "Chasers", "WindowState", Me.WindowState
    On Error Resume Next
    gcolMaxedWindows.Remove CStr(Me.hWnd)
End Sub

Private Sub smgChases_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'***************************************
' Module/Form Name   : frmChasers
'
' Procedure Name     : smgChases_MouseUp
'
' Purpose            :
'
' Date Created       : 29/12/2002
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
On Error GoTo smgChases_MouseUp_Error
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
    Set mdi_npls.fPopUp = Me
    PopupMenu mdi_npls.mnuChasers, vbPopupMenuRightButton, , , mdi_npls.mnuChasersAdd
'
'********* Code Ends Here **************
'
   Exit Sub
'
smgChases_MouseUp_Error:
    DisplayError , "frmChasers.smgChases_MouseUp", vbExclamation
End Sub

Public Sub AddChaser()
'***************************************
' Module/Form Name   : frmChasers
'
' Procedure Name     : AddChaser
'
' Purpose            :
'
' Date Created       : 29/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo AddChaser_Error
'
'******** Code Starts Here *************
'
    If Not mblnOutstandingPictures Then
        If MsgBox("There are no oustanding pictures on this delivery note." & vbCrLf & _
                  "Do you still wish to add a chaser record?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            Exit Sub
        End If
    End If
    '
    Set fChaser = New frmChaser
    fChaser.Display Add, _
                    moJob, _
                    moChasers
    Set fChaser = Nothing
'
'********* Code Ends Here **************
'
   Exit Sub
'
AddChaser_Error:
    ErrorRaise "frmChasers.AddChaser"
End Sub

Private Sub smgChases_RowChanged(CurrentRow As String)
    CheckPrintAvailability
End Sub

Private Sub CheckPrintAvailability()
    If smgChases.Column(4).Value <> "SL1" And _
       smgChases.Column(4).Value <> "SL2" And _
       smgChases.Column(4).Value <> "Loss/Fee" Then
        mdi_npls.mnuChasersPrint.Enabled = False
        cmdPrint.Enabled = False
    Else
        mdi_npls.mnuChasersPrint.Enabled = True
        cmdPrint.Enabled = True
    End If
End Sub

Public Sub PrintChaser()
'***************************************
' Module/Form Name   : frmChasers
'
' Procedure Name     : PrintChaser
'
' Purpose            :
'
' Date Created       : 29/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo PrintChaser_Error
'
'******** Code Starts Here *************
'

''    Dim oAccessDB As Access.Application
    Dim oAccessDB As Object
    Dim oReportRequest As ReportRequest
    Dim strId As String

    Set oReportRequest = New ReportRequest
    With oReportRequest
        .Read "OVERDUELETTER"
        .Parameter2 = CStr(moChasers.Item(smgChases.Column(1).Value).DeliveryNoteNo)
        .Parameter3 = smgChases.Column(1).Value
    End With
    Select Case smgChases.Column(4).Value
        Case Is = "SL1"
            With oReportRequest
                .Parameter1 = "SL1"
                .update
            End With
        Case Is = "SL2"
            With oReportRequest
                .Parameter1 = "SL2"
                .update
            End With
        Case Is = "Loss/Fee"
            With oReportRequest
                .Parameter1 = "Loss/Fee"
                .update
            End With
    End Select
    Set oAccessDB = GetObject(glo_dbname, "Access.Application")
    oAccessDB.Visible = True
    oAccessDB.DoCmd.Maximize
    oAccessDB.DoCmd.RunCommand 10       'acCmdAppMaximize
    oAccessDB.DoCmd.OpenReport "REP_OVERDUE_LETTER", 2      'acPreview
    Set oAccessDB = Nothing
    '
    '   Update the Chaser record print date.
    '
    With moChasers.Item(CStr(smgChases.Column(1).Value))
        .DatePrinted = CDate(Now)
        .update
    End With
    strId = smgChases.Column(1).Value
    RefreshChasers
    smgChases.GetKeyRow strId
    CheckPrintAvailability
'
'********* Code Ends Here **************
'
   Exit Sub
'
PrintChaser_Error:
    ErrorRaise "frmChasers.PrintChaser"
End Sub


