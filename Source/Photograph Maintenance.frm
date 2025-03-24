VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_photograph_maint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Photograph Maintenance"
   ClientHeight    =   2145
   ClientLeft      =   2265
   ClientTop       =   2055
   ClientWidth     =   6000
   Icon            =   "Photograph Maintenance.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2145
   ScaleWidth      =   6000
   Begin VB.Frame fraPhotograph 
      Height          =   795
      Left            =   120
      TabIndex        =   7
      Top             =   1020
      Width           =   5775
      Begin VB.TextBox txt_photograph_no 
         Height          =   315
         Left            =   4260
         TabIndex        =   9
         Top             =   300
         Width           =   915
      End
      Begin VB.Label lab_photoraph_no_lab 
         Caption         =   "&Photograph Number:"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   300
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "F&ind"
      Height          =   675
      Left            =   2640
      Picture         =   "Photograph Maintenance.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   60
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   675
      Left            =   120
      Picture         =   "Photograph Maintenance.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   60
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   675
      Left            =   960
      Picture         =   "Photograph Maintenance.frx":108E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   60
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   675
      Left            =   1800
      Picture         =   "Photograph Maintenance.frx":1750
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   60
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "&View"
      Height          =   675
      Left            =   2640
      Picture         =   "Photograph Maintenance.frx":1E12
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   60
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Label..."
      Height          =   675
      Left            =   3480
      Picture         =   "Photograph Maintenance.frx":24D4
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   60
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   675
      Left            =   5160
      Picture         =   "Photograph Maintenance.frx":2696
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   60
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin Crystal.CrystalReport crs_label 
      Left            =   3360
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin MSComDlg.CommonDialog cdg_print 
      Left            =   2640
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frm_photograph_maint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnInitial As Boolean

Private Sub cmdDelete_Click()
'***************************************
' Module/Form Name   : frm_photograph_maint
'
' Procedure Name     : cmdDelete_Click
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
On Error GoTo cmdDelete_Click_Error
'
'******** Code Starts Here *************
'
    If Len(RTrim(txt_photograph_no.Text)) = 0 Then
        MsgBox "Please enter a Photograph Number", vbExclamation, "Photograph Maintenance"
        txt_photograph_no.SetFocus
        Exit Sub
    End If

    photograph.Delete txt_photograph_no.Text, abort

'
'********* Code Ends Here **************
'
   Exit Sub
'
cmdDelete_Click_Error:
    DisplayError , "frm_photograph_maint.cmdDelete_Click", vbExclamation
End Sub

Private Sub cmdEdit_Click()
'***************************************
' Module/Form Name   : frm_photograph_maint
'
' Procedure Name     : cmdEdit_Click
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
On Error GoTo cmdEdit_Click_Error
'
'******** Code Starts Here *************
'
    If Len(RTrim(txt_photograph_no.Text)) = 0 Then
        MsgBox "Please enter a Photograph Number", vbExclamation, "Customer Maintenance"
        txt_photograph_no.SetFocus
        Exit Sub
    End If

    frm_photograph_edit.load_batch txt_photograph_no.Text

'
'********* Code Ends Here **************
'
   Exit Sub
'
cmdEdit_Click_Error:
    DisplayError , "frm_photograph_maint.cmdEdit_Click", vbExclamation
End Sub

Private Sub cmdExit_Click()
'***************************************
' Module/Form Name   : frm_photograph_maint
'
' Procedure Name     : cmdExit_Click
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
On Error GoTo cmdExit_Click_Error
'
'******** Code Starts Here *************
'
    Unload Me
    Set frm_photograph_maint = Nothing

'
'********* Code Ends Here **************
'
   Exit Sub
'
cmdExit_Click_Error:
    DisplayError , "frm_photograph_maint.cmdExit_Click", vbExclamation
End Sub

Private Sub cmdNew_Click()
    frm_photograph_new.Show
End Sub

Private Sub cmdView_Click()
'***************************************
' Module/Form Name   : frm_photograph_maint
'
' Procedure Name     : cmdView_Click
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
On Error GoTo cmdView_Click_Error
'
'******** Code Starts Here *************
'
    If Len(RTrim(txt_photograph_no.Text)) = 0 Then
        MsgBox "Please enter a Photograph Number", vbExclamation, "Customer Maintenance"
        txt_photograph_no.SetFocus
        Exit Sub
    End If

    Screen.MousePointer = ccHourglass
    frm_photograph_view.load_batch txt_photograph_no.Text
    Screen.MousePointer = ccDefault

'
'********* Code Ends Here **************
'
   Exit Sub
'
cmdView_Click_Error:
    DisplayError , "frm_photograph_maint.cmdView_Click", vbExclamation
End Sub

Private Sub Form_Initialize()
    mblnInitial = True
End Sub

Private Sub Form_Load()
    Screen.MousePointer = ccHourglass

    If goSystemConfig.BasicImageWhere Then
        cmdPrint.Visible = False
    End If

    txt_photograph_no.Text = GetSetting(App.Title, "PictureMaintenance", "PhotographNon", "")
    HighLightText txt_photograph_no
    Screen.MousePointer = ccDefault
End Sub

Private Sub cmdFind_Click()
'***************************************
' Module/Form Name   : frm_photograph_maint
'
' Procedure Name     : cmdFind_Click
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
On Error GoTo cmdFind_Click_Error
'
'******** Code Starts Here *************
'
    frm_photograph_find_maint.input_search
    frm_photograph_find_maint.ZOrder 0
'
'********* Code Ends Here **************
'
   Exit Sub
'
cmdFind_Click_Error:
    DisplayError , "frm_photograph_maint.cmdFind_Click", vbExclamation
End Sub

Private Sub cmdPrint_Click()
'***************************************
' Module/Form Name   : frm_photograph_maint
'
' Procedure Name     : cmdPrint_Click
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
On Error GoTo cmdPrint_Click_Error
'
'******** Code Starts Here *************
'
    frm_print_photo_labels.load_form txt_photograph_no
'
'********* Code Ends Here **************
'
   Exit Sub
'
cmdPrint_Click_Error:
    DisplayError , "frm_photograph_maint.cmdPrint_Click", vbExclamation
End Sub

Private Sub Form_Paint()
    If mblnInitial Then
        txt_photograph_no.SetFocus
    End If
    mblnInitial = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting App.Title, "PictureMaintenance", "PhotographNon", txt_photograph_no.Text
End Sub

Private Sub txt_photograph_no_GotFocus()
    With txt_photograph_no
        .SelStart = 0
        .SelLength = Len(txt_photograph_no)
    End With

End Sub

Private Sub txt_photograph_no_KeyPress(KeyAscii As Integer)
    KeyAscii = allow_numeric_only(KeyAscii)
End Sub


