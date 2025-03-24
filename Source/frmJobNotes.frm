VERSION 5.00
Begin VB.Form frmJobNotes 
   Caption         =   "Job Notes"
   ClientHeight    =   6150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9585
   Icon            =   "frmJobNotes.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6150
   ScaleWidth      =   9585
   Begin ImageWhere.SimpleGrid smgJobNotes 
      Height          =   3375
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   5953
      Columns         =   1
      KeyCol          =   0
   End
   Begin VB.TextBox txtJobReference 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000001&
      Height          =   525
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1140
      Width           =   7755
   End
   Begin VB.TextBox txtAddressLine1 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000001&
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   840
      Width           =   7755
   End
   Begin VB.TextBox txtCustomerName 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000001&
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   540
      Width           =   7755
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   375
      Left            =   7320
      TabIndex        =   9
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   8460
      TabIndex        =   10
      Top             =   5640
      Width           =   975
   End
   Begin VB.TextBox txtJobNo 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000001&
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   7755
   End
   Begin VB.Label lab_address_line_1_lab 
      Caption         =   "Address Line 1:"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   1155
   End
   Begin VB.Label lab_job_reference_lab 
      Caption         =   "Job Reference:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1140
      Width           =   1215
   End
   Begin VB.Label lab_customer_name_lab 
      Caption         =   "Customer Name:"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   540
      Width           =   1215
   End
   Begin VB.Label lblJobNo 
      Caption         =   "Job No:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmJobNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private moJob       As Job2
Private moCustomer  As Customer2
Private WithEvents mfJobNoteAdd As frmJobNoteAdd
Attribute mfJobNoteAdd.VB_VarHelpID = -1

Private Sub cmdAdd_Click()
'***************************************
' Module/Form Name   : frmJobNotes
'
' Procedure Name     : cmdAdd_Click
'
' Purpose            :
'
' Date Created       : 20/02/2005
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
    AddJobNote
'
'********* Code Ends Here **************
'
   Exit Sub
'
cmdAdd_Click_Error:
    DisplayError , "frmJobNotes.cmdAdd_Click", vbExclamation
End Sub

Private Sub AddJobNote()
'***************************************
' Module/Form Name   : frmJobNotes
'
' Procedure Name     : AddJobNote
'
' Purpose            :
'
' Date Created       : 20/02/2005
'
' Author             :
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
   On Error GoTo ErrorHandler:
'
'******** Code Starts Here *************
'
    
    Set mfJobNoteAdd = New frmJobNoteAdd
    mfJobNoteAdd.Display moJob, moCustomer
    Set mfJobNoteAdd = Nothing
'
'********* Code Ends Here **************
'
   On Error GoTo 0
   Exit Sub
'
ErrorHandler:
    ErrorRaise "frmJobNotes.cmdAdd_Click"
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Public Sub Display(poJob As Job2)
'***************************************
' Module/Form Name   : frmJobNotes
'
' Procedure Name     : Display
'
' Purpose            :
'
' Date Created       : 18/02/2005
'
' Author             : GARETH SAUNDERS
'
' Parameters         : poJob - Job2
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
    
    Dim oActivity   As Activity
    
    Set moJob = poJob
    '
    '   Set the Headings
    '
    txtJobNo.Text = moJob.JobNo
    On Error Resume Next
    moCustomer.Read moJob.CustomerNo
    On Error GoTo Display_Error
    txtCustomerName.Text = moCustomer.CustomerName
    txtAddressLine1.Text = moCustomer.Address1
    txtJobReference.Text = moJob.reference
    '
    RefreshJobNotes 0
    '
    Me.Show vbModal
'
'********* Code Ends Here **************
'
   Exit Sub
'
Display_Error:
    ErrorRaise "frmJobNotes.Display"
End Sub

Private Sub RefreshJobNotes(pDateAdded As Date)
'***************************************
' Module/Form Name   : frmJobNotes
'
' Procedure Name     : RefreshJobNotes
'
' Purpose            :
'
' Date Created       : 20/02/2005
'
' Author             :
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
   On Error GoTo ErrorHandler:
'
'******** Code Starts Here *************
'
    Dim oJobNote As Activity
    '
    '   Display the Job Notes.
    '
    moJob.JobNotes.Refresh
    LockWindow Me.hWnd
    smgJobNotes.Clear
    For Each oJobNote In moJob.JobNotes
        With smgJobNotes
            .AddRow False, _
                    Format(oJobNote.StartDate, "dd/mm/yyyy"), _
                    oJobNote.UserField1, _
                    oJobNote.UserField2, _
                    oJobNote.Description
        End With
    Next oJobNote
    smgJobNotes.ResizeRows
    smgJobNotes.GetKeyRow CStr(pDateAdded)
    UnlockWindow
'
'********* Code Ends Here **************
'
   On Error GoTo 0
   Exit Sub
'
ErrorHandler:
    ErrorRaise "frmJobNotes.cmdAdd_Click"
End Sub

Private Sub Form_Load()
    Dim lngTop  As Long
    Dim lngLeft As Long
    
    Set moCustomer = New Customer2
    With smgJobNotes
        .Columns = 4
        .Column(1).Header = "Date"
        .Column(2).Header = "Contact"
        .Column(3).Header = "Recorded By"
        .Column(4).Header = "Note"
        .KeyCol = 1
    End With
    
    Me.WindowState = GetSetting(App.Title, "JobNotes", "WindowState", 0)
    If GetSetting(App.Title, "JobNotes", "WindowState", 0) <> 2 Then
        Me.Width = GetSetting(App.Title, "JobNotes", "Width", 8000)
        Me.Height = GetSetting(App.Title, "JobNotes", "Height", 5500)
        lngTop = GetSetting(App.Title, "JobNotes", "Top", 0)
        lngLeft = GetSetting(App.Title, "JobNotes", "Left", 0)
        If lngTop <> 0 And lngLeft <> 0 Then
            Me.Top = GetSetting(App.Title, "JobNotes", "Top", 0)
            Me.Left = GetSetting(App.Title, "JobNotes", "Left", 0)
        Else
            Me.Top = (mdi_npls.Height - Me.Height) / 2
            Me.Left = (mdi_npls.Width - Me.Width) / 2
        End If
    End If
    
End Sub

Private Sub Form_Resize()
    Dim MIN_WIDTH As Integer
    Dim MIN_HEIGHT As Integer
    Dim header_height As Integer
    
'***************************************
' Module/Form Name   : frmJobNotes
'
' Procedure Name     : Form_Resize
'
' Purpose            :
'
' Date Created       : 03/10/2006 23:34
'
' Author             : Gareth
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
    On Error GoTo ErrorProc

    MIN_WIDTH = 8000
    MIN_HEIGHT = 5500

    If WindowState = vbMinimized Then
        Exit Sub
    End If

    If WindowState = vbNormal Then
        If Width < MIN_WIDTH Then
            Width = MIN_WIDTH
        End If
        If Height < MIN_HEIGHT Then
            Height = MIN_HEIGHT
        End If
    End If
        
    LockWindow Me.hWnd
    cmdClose.Move Me.Width - 250 - cmdClose.Width, _
                  Me.Height - goSystemConfig.TitleBarHeight - 200 - cmdClose.Height
    cmdAdd.Move cmdClose.Left - 200 - cmdAdd.Width, _
                cmdClose.Top
    With smgJobNotes
        .Width = Me.Width - (.Left * 2) - 100
        .Height = cmdClose.Top - .Top - 100
        .Column(1).Width = 1500
        .Column(2).Width = 2000
        .Column(3).Width = 1500
        .Column(4).Width = .Width - (.Column(1).Width + _
                                     .Column(2).Width + _
                                     .Column(3).Width + _
                                     goSystemConfig.VScrollBarWidth + 103)
        .ResizeRows
    End With

    UnlockWindow
    DoEvents
    
    On Error GoTo 0
    Exit Sub
'
'********* Code Ends Here **************
'
ErrorProc:
    DisplayError , "frmJobNotes.Form_Resize", vbExclamation
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting App.Title, "JobNotes", "Width", Me.Width
    SaveSetting App.Title, "JobNotes", "Height", Me.Height
    SaveSetting App.Title, "JobNotes", "Top", Me.Top
    SaveSetting App.Title, "JobNotes", "Left", Me.Left
    SaveSetting App.Title, "JobNotes", "WindowState", Me.WindowState
    
    Set moCustomer = Nothing
End Sub

Private Sub mfJobNoteAdd_NoteAdded(DateAdded As Date)
    RefreshJobNotes DateAdded
End Sub
