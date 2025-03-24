VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_photograph_find_maint 
   Caption         =   "Find & maintain photographs"
   ClientHeight    =   4080
   ClientLeft      =   2190
   ClientTop       =   3345
   ClientWidth     =   6945
   Icon            =   "Photograph Find and Maintain.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4080
   ScaleWidth      =   6945
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   675
      Left            =   5160
      Picture         =   "Photograph Find and Maintain.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   60
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Label..."
      Height          =   675
      Left            =   4320
      Picture         =   "Photograph Find and Maintain.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   60
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "&View"
      Height          =   675
      Left            =   2640
      Picture         =   "Photograph Find and Maintain.frx":0B8E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   60
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   675
      Left            =   1800
      Picture         =   "Photograph Find and Maintain.frx":1250
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   60
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   675
      Left            =   960
      Picture         =   "Photograph Find and Maintain.frx":1912
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   60
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   675
      Left            =   120
      Picture         =   "Photograph Find and Maintain.frx":1FD4
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   60
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "F&ind"
      Height          =   675
      Left            =   3480
      Picture         =   "Photograph Find and Maintain.frx":2696
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   60
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin MSComDlg.CommonDialog cdg_print 
      Left            =   6000
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Crystal.CrystalReport crs_label 
      Left            =   6120
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin ImageWhere.SimpleGrid smgPhotographs 
      Height          =   1935
      Left            =   240
      TabIndex        =   8
      Top             =   1200
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   3413
      Columns         =   1
      KeyCol          =   0
   End
   Begin VB.Frame fraPhotographs 
      Height          =   2535
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   6255
   End
End
Attribute VB_Name = "frm_photograph_find_maint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public search_string As String
Private mrstPhotos As DAO.Recordset

Public Sub display_find_results(rstPhotos As Recordset, search_text As String)
      '***************************************
      ' Module/Form Name   : frm_photograph_find_maint
      '
      ' Procedure Name     : display_find_results
      '
      ' Purpose            :
      '
      ' Date Created       : 07/12/2002
      '
      ' Author             : GARETH SAUNDERS
      '
      ' Parameters         : rstPhotos - Recordset
      '                    : search_text - String
      '
      ' Amendment History  : Date       Author    Description
      '                    : --------------------------------
      '
      '***************************************
      '
10    On Error GoTo display_find_results_Error
      '
      '******** Code Starts Here *************
      '

20        Me.Caption = "Find & maintain photographs " & "containing " & search_text
30        Set mrstPhotos = rstPhotos.Clone
40        DisplayPhotos
50        smgPhotographs.CurrentRow = 1
60        smgPhotographs.TopRow = 1
      '
      '********* Code Ends Here **************
      '
70       Exit Sub
      '
display_find_results_Error:
80        ErrorRaise "frm_photograph_find_maint.display_find_results"
End Sub

Public Sub input_search()
      '***************************************
      ' Module/Form Name   : frm_photograph_find_maint
      '
      ' Procedure Name     : input_search
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
10    On Error GoTo input_search_Error
      '
      '******** Code Starts Here *************
      '
20        frm_photograph_find_maint.Show
30        DoEvents
40        frm_photograph_find.Show 1
      '
      '********* Code Ends Here **************
      '
50       Exit Sub
      '
input_search_Error:
60        ErrorRaise "frm_photograph_find_maint.input_search"
End Sub

Public Sub refresh_list(Optional photograph_no As Variant)
      '***************************************
      ' Module/Form Name   : frm_photograph_find_maint
      '
      ' Procedure Name     : refresh_list
      '
      ' Purpose            :
      '
      ' Date Created       : 07/12/2002
      '
      ' Author             : GARETH SAUNDERS
      '
      ' Parameters         : photograph_no - Variant
      '
      ' Amendment History  : Date       Author    Description
      '                    : --------------------------------
      '
      '***************************************
      '
10    On Error GoTo refresh_list_Error
      '
      '******** Code Starts Here *************
      '
          Dim find_Criteria As String
    
20        Screen.MousePointer = vbHourglass
30        photograph.search search_string, mrstPhotos, dbOpenSnapshot, abort
40        If abort Then Exit Sub
          '
50        DisplayPhotos
          '
          '   Find the photograph no in the list and re-highlight.
          '
60        If Not IsMissing(photograph_no) Then
70            smgPhotographs.GetKeyRow CStr(photograph_no)
80        End If
    
90        Screen.MousePointer = vbDefault
      '
      '********* Code Ends Here **************
      '
100      Exit Sub
      '
refresh_list_Error:
110       Select Case Err.Number
              Case Is = 91
120               Screen.MousePointer = vbDefault
130               Exit Sub
140           Case Else
150               DisplayError , "frm_photograph_find_maint.refresh_list", vbExclamation
160       End Select
End Sub

Private Sub cmdDelete_Click()
      '***************************************
      ' Module/Form Name   : frm_photograph_find_maint
      '
      ' Procedure Name     : cmdDelete_Click
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
10    On Error GoTo cmdDelete_Click_Error
      '
      '******** Code Starts Here *************
      '
20        If smgPhotographs.CurrentRow = 0 Then
30            MsgBox "Select a photograph", vbExclamation, App.Title
40            Exit Sub
50        End If
    
60        photograph.Delete smgPhotographs.Column(1).Value, abort
      '
      '********* Code Ends Here **************
      '
70       Exit Sub
      '
cmdDelete_Click_Error:
80        DisplayError , "frm_photograph_find_maint.cmdDelete_Click", vbExclamation
End Sub

Private Sub cmdEdit_Click()
      '***************************************
      ' Module/Form Name   : frm_photograph_find_maint
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
10    On Error GoTo cmdEdit_Click_Error
      '
      '******** Code Starts Here *************
      '
20        If smgPhotographs.CurrentRow = 0 Then
30            MsgBox "Select a photograph", vbExclamation, App.Title
40            Exit Sub
50        End If
    
60        Screen.MousePointer = vbHourglass
70        frm_photograph_edit.load_batch smgPhotographs.Column(1).Value
80        Screen.MousePointer = vbDefault
      '
      '********* Code Ends Here **************
      '
90       Exit Sub
      '
cmdEdit_Click_Error:
100       DisplayError , "frm_photograph_find_maint.cmdEdit_Click", vbExclamation
End Sub

Private Sub cmdExit_Click()
10        Unload Me
20        Set frm_photograph_find_maint = Nothing
End Sub

Private Sub cmdNew_Click()
      '***************************************
      ' Module/Form Name   : frm_photograph_find_maint
      '
      ' Procedure Name     : cmdNew_Click
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
10    On Error GoTo cmdNew_Click_Error
      '
      '******** Code Starts Here *************
      '
20        frm_photograph_new.Show
      '
      '********* Code Ends Here **************
      '
30       Exit Sub
      '
cmdNew_Click_Error:
40        DisplayError , "frm_photograph_find_maint.cmdNew_Click", vbExclamation
End Sub

Private Sub cmdView_Click()
      '***************************************
      ' Module/Form Name   : frm_photograph_find_maint
      '
      ' Procedure Name     : cmdView_Click
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
10    On Error GoTo cmdView_Click_Error
      '
      '******** Code Starts Here *************
      '
20        If smgPhotographs.CurrentRow = 0 Then
30            MsgBox "Select a photograph", vbExclamation, App.Title
40            Exit Sub
50        End If
    
60        frm_photograph_view.load_batch smgPhotographs.Column(1).Value
      '
      '********* Code Ends Here **************
      '
70       Exit Sub
      '
cmdView_Click_Error:
80        DisplayError , "frm_photograph_find_maint.cmdView_Click", vbExclamation
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrstPhotos = Nothing
    On Error Resume Next
    gcolMaxedWindows.Remove CStr(Me.hWnd)
End Sub

Private Sub smgPhotographs_DblClick()
      '***************************************
      ' Module/Form Name   : frm_photograph_find_maint
      '
      ' Procedure Name     : smgPhotographs_DblClick
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
10    On Error GoTo smgPhotographs_DblClick_Error
      '
      '******** Code Starts Here *************
      '
20        Screen.MousePointer = vbHourglass
30        frm_photograph_edit.load_batch smgPhotographs.Column(1).Value
40        Screen.MousePointer = vbDefault
      '
      '********* Code Ends Here **************
      '
50       Exit Sub
      '
smgPhotographs_DblClick_Error:
60        DisplayError , "frm_photograph_find_maint.smgPhotographs_DblClick", vbExclamation
End Sub

Private Sub Form_Activate()
10        Me.WindowState = 2
End Sub

Private Sub Form_Load()
10        Me.Caption = "Find & maintain photographs "
20        With smgPhotographs
30            .Columns = 2
40            .KeyCol = 1
50            .Column(1).Header = "Number"
60            .Column(2).Header = "Description"
70        End With

     Set smgPhotographs.Container = fraPhotographs
End Sub

Private Sub Form_Resize()
'***************************************
' Module/Form Name   : frm_photograph_find_maint
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
    Dim lngColumn2Width As Long
    
    If Not ResizeForm(Me) Then
        Exit Sub
    End If
    
    cmdExit.Left = frm_photograph_find_maint.Width - 960
    
    If frm_photograph_find_maint.Width > 1000 Then
        fraPhotographs.Width = frm_photograph_find_maint.Width - 400
    End If
    If frm_photograph_find_maint.Height > 1500 Then
        fraPhotographs.Height = frm_photograph_find_maint.Height - 1400
    End If
    fraPhotographs.Left = 200
    
    If goSystemConfig.BasicImageWhere Then
        cmdPrint.Visible = False
    End If
    
    With smgPhotographs
        .Visible = False
        .Left = 100
        .Width = fraPhotographs.Width - 2 * .Left
        .Top = 200
        .Height = fraPhotographs.Height - 2 * .Top
        .Column(1).Width = 1000
        lngColumn2Width = CLng(.Width) - _
                          CLng(.Column(1).Width) - _
                          CLng(goSystemConfig.VScrollBarWidth) - 100
'        .Column(2).Width = CLng(CLng(.Width) - _
'                                CLng(.Column(1).Width) - _
'                                CLng(goSystemConfig.VScrollBarWidth) - 100)
        .Column(2).Width = IIf(lngColumn2Width > 32767, 32767, lngColumn2Width)

        .Visible = True
    End With

'
'********* Code Ends Here **************
'
   Exit Sub
'
Form_Resize_Error:
    DisplayError , "frm_photograph_find_maint.Form_Resize", vbExclamation
End Sub

Private Sub cmdFind_Click()
      '***************************************
      ' Module/Form Name   : frm_photograph_find_maint
      '
      ' Procedure Name     : cmdFind_Click
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
10    On Error GoTo cmdFind_Click_Error
      '
      '******** Code Starts Here *************
      '
20        frm_photograph_find.display_photograph_find search_string
      '
      '********* Code Ends Here **************
      '
30       Exit Sub
      '
cmdFind_Click_Error:
40        DisplayError , "frm_photograph_find_maint.cmdFind_Click", vbExclamation
End Sub

Private Sub cmdPrint_Click()
      '***************************************
      ' Module/Form Name   : frm_photograph_find_maint
      '
      ' Procedure Name     : cmdPrint_Click
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
10    On Error GoTo cmdPrint_Click_Error
      '
      '******** Code Starts Here *************
      '
20        If smgPhotographs.CurrentRow = 0 Then
30            MsgBox "Select a photograph", vbExclamation, App.Title
40            Exit Sub
50        End If
    
60        frm_print_photo_labels.load_form LTrim(CStr(smgPhotographs.Column(1).Value))
      '
      '********* Code Ends Here **************
      '
70       Exit Sub
      '
cmdPrint_Click_Error:
80        DisplayError , "frm_photograph_find_maint.cmdPrint_Click", vbExclamation
End Sub

Private Sub DisplayPhotos()
      '***************************************
      ' Module/Form Name   : frm_photograph_find_maint
      '
      ' Procedure Name     : DisplayPhotos
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
10    On Error GoTo DisplayPhotos_Error
      '
      '******** Code Starts Here *************
      '
20        With smgPhotographs
30            Screen.MousePointer = vbHourglass
40            .Redraw = False
50            .Clear
60            If Not mrstPhotos.EOF Then
70                mrstPhotos.MoveFirst
80            End If
90            While Not mrstPhotos.EOF
100               .AddRow False, _
                          mrstPhotos!batch_no, _
                          mrstPhotos!Description
110               mrstPhotos.MoveNext
120           Wend
130           .ResizeRows
140           .Redraw = True
150           Screen.MousePointer = vbDefault
160       End With
      '
      '********* Code Ends Here **************
      '
170      Exit Sub
      '
DisplayPhotos_Error:
180       ErrorRaise "frm_photograph_find_maint.DisplayPhotos"
End Sub

