VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmApplyKeywordsToPhotographs 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Apply Keywords To Photographs"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   Icon            =   "frmApplyKeywordsToPhotographs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   6630
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkUseExisting 
      Caption         =   "Use Existing Extract"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Frame fraRange 
      Caption         =   "Range"
      Height          =   855
      Left            =   240
      TabIndex        =   9
      Top             =   1200
      Width           =   6135
      Begin VB.TextBox txtTo 
         Height          =   285
         Left            =   4440
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtFrom 
         Height          =   285
         Left            =   2280
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
      Begin VB.CheckBox chkAll 
         Caption         =   "All"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblTo 
         Caption         =   "To:"
         Height          =   255
         Left            =   3960
         TabIndex        =   2
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblFrom 
         Caption         =   "From:"
         Height          =   255
         Left            =   1680
         TabIndex        =   0
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   555
      Left            =   5640
      Picture         =   "frmApplyKeywordsToPhotographs.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2160
      UseMaskColor    =   -1  'True
      Width           =   795
   End
   Begin MSComctlLib.ProgressBar pgbUpdate 
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   555
      Left            =   4320
      Picture         =   "frmApplyKeywordsToPhotographs.frx":089C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2160
      UseMaskColor    =   -1  'True
      Width           =   795
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      Height          =   555
      Left            =   5640
      Picture         =   "frmApplyKeywordsToPhotographs.frx":0E16
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2160
      UseMaskColor    =   -1  'True
      Width           =   795
   End
   Begin VB.Label lblProgressUpdate 
      Caption         =   "Label1"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   6135
   End
End
Attribute VB_Name = "frmApplyKeywordsToPhotographs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mvarCancel  As Boolean

Private Sub chkAll_Click()
    EnableDisableButtons
End Sub

Private Sub cmdOK_Click()
'***************************************
' Module/Form Name   : frmApplyKeywordsToPhotographs
'
' Procedure Name     : cmdOK_Click
'
' Purpose            :
'
' Date Created       : 14/07/2005
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo cmbOK_Click_Error
'
'******** Code Starts Here *************
'
    Dim oBatches                As Batches
    Dim oBatch                  As Batch2
    Dim oBatchKeywords          As BatchKeywords
    Dim lngCount                As Long
    Dim oKeywordExclusions      As KeywordExclusions
    Dim oKeywordSynonyms        As KeywordSynonyms
    Dim oFSO                    As Scripting.FileSystemObject
    Dim strFile                 As String
    Dim intFileNo               As Integer
    
    If Not ValidEntry Then
        Exit Sub
    End If
    
    If MsgBox("This process may take a long time. Do you wish to continue?", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
        Exit Sub
    End If
    '
    MsgBox "Please ensure that:" & vbCrLf & vbCrLf & _
           "1. A full database backup has been performed before applying all keywords." & vbCrLf & _
           "2. All other users are logged out of Image Where completely before continuing.", vbInformation
    '
    '   For each Batch Record
    '
    Screen.MousePointer = vbHourglass
    cmdClose.Visible = False
    cmdClose.Enabled = False
    cmdCancel.Visible = True
    cmdCancel.Enabled = True
    cmdOK.Enabled = False
    
    Set oFSO = New Scripting.FileSystemObject
    strFile = oFSO.BuildPath(gstrUpgradePath, "\BatchKeywords.txt")
    Set oFSO = Nothing
    
    Set oBatches = New Batches
    
    If chkUseExisting.Value = vbChecked Then GoTo DeleteExistingPhotos
        
    lblProgressUpdate.Caption = "Retrieving all Photographs"
    pgbUpdate.Value = 0
    DoEvents
    oBatches.Refresh CLng(txtFrom.Text), CLng(txtTo.Text)
    pgbUpdate.Min = 0
    pgbUpdate.Max = Val(txtTo.Text) - Val(txtFrom.Text) + 1
    lngCount = 0
    '
    '   Get the Keyword Exclusions and Synonyms
    '
    lblProgressUpdate.Caption = "Retrieving Keyword Exclusions"
    DoEvents
    Set oKeywordExclusions = New KeywordExclusions
    oKeywordExclusions.Refresh
    '
    lblProgressUpdate.Caption = "Retrieving Keyword Synonyms"
    DoEvents
    Set oKeywordSynonyms = New KeywordSynonyms
    oKeywordSynonyms.Refresh
    '
    '   Open the File to export the Keywords to.
    '
    intFileNo = FreeFile
    Close
    Open strFile For Output As #intFileNo
    Print #intFileNo, "BATCHNO,KEYWORD,KEYWORDTYPE"
    '
    pgbUpdate.Visible = True
    For Each oBatch In oBatches
                
        If oBatch.BatchNo >= CLng(txtFrom.Text) And _
           oBatch.BatchNo <= CLng(txtTo.Text) Then
            
            lngCount = lngCount + 1
            
            lblProgressUpdate.Caption = "Extracting Data for Photograph: " & Format(oBatch.BatchNo, "00000")
            DoEvents
            '
            '   Add keywords from Caption, ignoring excluded words and adding any clustered ones.
            '
            oBatch.CreateAutomaticKeywords oKeywordExclusions, oKeywordSynonyms, 1
            pgbUpdate.Value = IIf(lngCount < pgbUpdate.Max, lngCount, pgbUpdate.Max)
            DoEvents
            If mvarCancel Then
                If MsgBox("Are you sure you wish to cancel the job?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                    Exit For
                Else
                    mvarCancel = False
                End If
            End If
        End If
    Next oBatch
    Close #intFileNo
    If mvarCancel Then
        cmdClose.Visible = True
        cmdClose.Enabled = True
        cmdCancel.Visible = False
        cmdCancel.Enabled = False
        Set oBatches = Nothing
        MsgBox "Cancelled by User", vbInformation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
DeleteExistingPhotos:
    '
    '   Remove all Batch Keywords
    '
    pgbUpdate.Visible = False
    lblProgressUpdate.Caption = "Removing Existing Keywords"
    DoEvents
    Set oBatchKeywords = New BatchKeywords
    oBatchKeywords.DeleteAll CLng(txtFrom.Text), CLng(txtTo.Text)
    Set oBatchKeywords = Nothing
    DoEvents
    '
    '   Clear down memory before importing data.
    '
    lblProgressUpdate.Caption = "Tidying up after extract..."
    DoEvents
    Set oBatches = Nothing
    Set oBatches = New Batches
    Set oKeywordExclusions = Nothing
    Set oKeywordSynonyms = Nothing
    '
    '   Import the text file into the database.
    '
    lblProgressUpdate.Caption = "Importing Keywords"
    DoEvents
''    oBatches.ImportTextFile strFile
    oBatches.ImportTextFileWithAccess strFile
    DoEvents
    Set oBatches = Nothing
    '
    pgbUpdate.Value = pgbUpdate.Max
    lblProgressUpdate.Caption = "Completed Successfully"
    SaveSetting App.Title, "ApplyKeywords", "All", IIf(chkAll.Value = vbChecked, "Y", "N")
    SaveSetting App.Title, "ApplyKeywords", "UseExisting", IIf(chkUseExisting.Value = vbChecked, "Y", "N")
    SaveSetting App.Title, "ApplyKeywords", "From", txtFrom.Text
    SaveSetting App.Title, "ApplyKeywords", "To", txtTo.Text
    
    Screen.MousePointer = vbDefault
    '
    '   Recommend that the database is compacted afterwards.
    '
    MsgBox "Keywords have now been applied to all photographs." & vbCrLf & _
           "It is recommended that you now compact the database.", vbInformation
    cmdClose.Visible = True
    cmdClose.Enabled = True
    cmdCancel.Visible = False
    cmdCancel.Enabled = False
    chkAll.Enabled = False
    chkUseExisting.Enabled = False
    txtFrom.Enabled = False
    txtTo.Enabled = False
    txtFrom.BackColor = vbButtonFace
    txtTo.BackColor = vbButtonFace
'
'********* Code Ends Here **************
'
   Exit Sub
'
cmbOK_Click_Error:
    cmdCancel.Visible = False
    cmdClose.Visible = True
    cmdClose.Enabled = True
    chkAll.Enabled = False
    chkUseExisting.Enabled = False
    Close #1
    DisplayError , "frmApplyKeywordsToPhotographs.cmbOK_Click", vbExclamation
    DoEvents
End Sub

Private Sub cmdCancel_Click()
    mvarCancel = True
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    mvarCancel = False
    lblProgressUpdate.Caption = ""
    pgbUpdate.Value = 0
    chkAll.Value = IIf(GetSetting(App.Title, "ApplyKeywords", "All", "Y") = "Y", vbChecked, vbUnchecked)
    chkUseExisting.Value = IIf(GetSetting(App.Title, "ApplyKeywords", "UseExisting", "N") = "Y", vbChecked, vbUnchecked)
    txtFrom.Text = GetSetting(App.Title, "ApplyKeywords", "From", "1")
    txtTo.Text = GetSetting(App.Title, "ApplyKeywords", "To", CStr(goSystemConfig.LastPhotoNumberUsed))
End Sub

Private Sub EnableDisableButtons()
    If chkAll.Value = vbChecked Then
        txtFrom.Enabled = False
        txtFrom.BackColor = vbButtonFace
        txtFrom.Text = "1"
        txtTo.Enabled = False
        txtTo.BackColor = vbButtonFace
        txtTo.Text = CStr(goSystemConfig.LastPhotoNumberUsed)
    Else
        txtFrom.Enabled = True
        txtFrom.BackColor = vbWindowBackground
        txtTo.Enabled = True
        txtTo.BackColor = vbWindowBackground
    End If
End Sub

Private Function ValidEntry() As Boolean
    ValidEntry = False
    
    If Val(txtFrom.Text) > Val(txtTo.Text) Then
        MsgBox "'From' must be less than or equal to 'To'", vbExclamation
        txtFrom.SetFocus
        Exit Function
    End If
    
    If Val(txtFrom.Text) = 0 Or Val(txtTo.Text) = 0 Then
        MsgBox "Please enter non zero From and To values", vbExclamation
        txtFrom.SetFocus
        Exit Function
    End If

    If Val(txtTo.Text) > 100000 Then
        MsgBox "Please enter a value less than 100,000", vbExclamation
        txtTo.SetFocus
        Exit Function
    End If
    
    ValidEntry = True

End Function

Private Sub txtFrom_GotFocus()
    HighLightText txtFrom
End Sub

Private Sub txtFrom_KeyPress(KeyAscii As Integer)
    KeyAscii = allow_numeric_only(KeyAscii)
End Sub

Private Sub txtTo_GotFocus()
    HighLightText txtTo
End Sub

Private Sub txtTo_KeyPress(KeyAscii As Integer)
    KeyAscii = allow_numeric_only(KeyAscii)
End Sub
