VERSION 5.00
Begin VB.Form frmKeywordClustering 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Keyword Clustering"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8280
   Icon            =   "frmKeywordClustering.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   8280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin ImageWhere.SimpleGrid smgKeywordClustering 
      Height          =   2055
      Left            =   2040
      TabIndex        =   5
      Top             =   1320
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   3625
      Columns         =   1
      KeyCol          =   0
   End
   Begin VB.Frame fraKeywordSynonyms 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      Begin VB.CommandButton cmdReplaceKeyword 
         Caption         =   "&Replace"
         Height          =   315
         Left            =   5640
         TabIndex        =   7
         Top             =   3720
         Width           =   1095
      End
      Begin VB.CommandButton cmdNewKeyword 
         Caption         =   "&New"
         Height          =   315
         Left            =   4440
         TabIndex        =   6
         Top             =   3720
         Width           =   1095
      End
      Begin VB.CommandButton cmdDeleteKeyword 
         Caption         =   "&Delete"
         Height          =   315
         Left            =   6840
         TabIndex        =   8
         Top             =   3720
         Width           =   1095
      End
      Begin VB.TextBox txtKeyword 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   3850
      End
      Begin VB.TextBox txtSynonym 
         Height          =   285
         Left            =   4080
         TabIndex        =   4
         Top             =   480
         Width           =   3850
      End
      Begin VB.Label lblKeyword 
         Caption         =   "&Keyword:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label lblSynonym 
         Caption         =   "&Synonym:"
         Height          =   255
         Left            =   4080
         TabIndex        =   3
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   7240
      TabIndex        =   9
      Top             =   4560
      Width           =   975
   End
End
Attribute VB_Name = "frmKeywordClustering"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private moKeywordSynonyms   As KeywordSynonyms
Private mblnUnhandleEvents  As Boolean
Private mblnKeywordFound    As Boolean

Public Sub Display()
    Set moKeywordSynonyms = New KeywordSynonyms
    
    moKeywordSynonyms.Refresh
    
    With smgKeywordClustering
        .Columns = 2
        .KeyCol = 1
        .Column(1).Header = "Keyword"
        .Column(2).Header = "Synonym"
        .Column(1).Width = (.Width - 364) / 2
        .Column(2).Width = .Column(1).Width
        .Height = 2700
    End With
    '
    txtKeyword.Move smgKeywordClustering.Left, txtKeyword.Top, smgKeywordClustering.Column(1).Width + 50
    txtSynonym.Move smgKeywordClustering.Column(1).Width + 150, txtKeyword.Top, smgKeywordClustering.Width - (smgKeywordClustering.Column(1).Width + 50)
    lblKeyword.Move txtKeyword.Left, lblKeyword.Top, 2000
    lblSynonym.Move txtSynonym.Left, lblSynonym.Top, 2000
''    cmdReplaceKeyword.Move cmdNewKeyword.Left, cmdNewKeyword.Top
    '
    cmdNewKeyword.Enabled = False
    cmdDeleteKeyword.Enabled = False
    cmdReplaceKeyword.Enabled = False
    '
    Screen.MousePointer = vbHourglass
    DisplayClustering
    Screen.MousePointer = vbDefault
    '
    Me.Show vbModal
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub DisplayClustering(Optional pstrKeyword As String = "", Optional pstrSynonym As String = "")
    
    Dim oKeywordSynonym     As KeywordSynonym
    
'***************************************
' Module/Form Name   : frmKeywordClustering
'
' Procedure Name     : DisplayClustering
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
    Dim intRow      As Integer
    
    LockWindow Me.hWnd
    With smgKeywordClustering
        .Redraw = False
        .Clear
        For Each oKeywordSynonym In moKeywordSynonyms
            .AddRow False, oKeywordSynonym.Keyword, oKeywordSynonym.Synonym
        Next oKeywordSynonym

        If moKeywordSynonyms.Count > 0 Then
            If pstrKeyword <> "" Then
                .GetKeyRow pstrKeyword
            Else
                .Deselect
                .TopRow = 1
            End If
            '
            If .CurrentRow <> 0 Then
                txtKeyword.Text = .Column(1).Value
                txtSynonym.Text = .Column(2).Value
            End If
        End If
        .Redraw = True
    End With
    UnlockWindow
'
'********* Code Ends Here **************
'
    On Error GoTo 0
    Exit Sub
'
ErrorHandler:
    ErrorRaise "frmKeywordClustering.DisplayClustering"
End Sub

Private Sub cmdDeleteKeyword_Click()
'***************************************
' Module/Form Name   : frmKeywordClustering
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
    DeleteKeywordSynonym
'
'********* Code Ends Here **************
'
    On Error GoTo 0
    Exit Sub
'
ErrorHandler:
    DisplayError , "frmKeywordClustering.cmdDeleteKeyword_Click", vbExclamation
End Sub

Private Sub DeleteKeywordSynonym()
    Dim strKeyword      As String
    
'***************************************
' Module/Form Name   : frmKeywordClustering
'
' Procedure Name     : DeleteKeywordSynonym
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
    strKeyword = smgKeywordClustering.Column(1).Value
    If MsgBox("Are you sure you wish to delete keyword synonym '" & strKeyword & "/" & smgKeywordClustering.Column(2).Value & "'", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
        Exit Sub
    End If
    
    moKeywordSynonyms.RemovePending strKeyword & "/" & smgKeywordClustering.Column(2).Value
    moKeywordSynonyms.update
    DisplayClustering
    '
    SetControlsState
''    SetKeywordUpdateButtons
    '
    On Error Resume Next
    txtKeyword.SetFocus
'
'********* Code Ends Here **************
'
    On Error GoTo 0
    Exit Sub
'
ErrorHandler:
    ErrorRaise "frmKeywordClustering.DeleteKeywordSynonym"
End Sub

Private Sub cmdNewKeyword_Click()
'***************************************
' Module/Form Name   : frmKeywordClustering
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
    NewKeywordSynonym
'
'********* Code Ends Here **************
'
    On Error GoTo 0
    Exit Sub
'
ErrorHandler:
    DisplayError , "frmKeywordClustering.cmdNewKeyword_Click", vbExclamation
End Sub

Private Sub NewKeywordSynonym()
    
'***************************************
' Module/Form Name   : frmKeywordClustering
'
' Procedure Name     : NewKeywordSynonym
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
    moKeywordSynonyms.Add txtKeyword.Text, txtSynonym.Text
    moKeywordSynonyms.update
    DisplayClustering txtKeyword.Text, txtSynonym.Text
    '
    SetKeywordUpdateButtons
    '
    On Error Resume Next
    txtKeyword.SetFocus
'
'********* Code Ends Here **************
'
    On Error GoTo 0
    Exit Sub
'
ErrorHandler:
    ErrorRaise "frmKeywordClustering.NewKeywordSynonym"
End Sub

Private Sub cmdReplaceKeyword_Click()
'***************************************
' Module/Form Name   : frmKeywordClustering
'
' Procedure Name     : cmdReplaceKeyword_Click
'
' Purpose            :
'
' Date Created       : 25/07/2005
'
' Author             : Gareth Saunders
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                      25/07/2005 Gareth Saunders
'
'***************************************
'
    On Error GoTo ErrorHandler:
'
'******** Code Starts Here *************
'
    ReplaceKeywordSynonym
'
'********* Code Ends Here **************
'
    On Error GoTo 0
    Exit Sub
'
ErrorHandler:
    DisplayError , "frmKeywordClustering.cmdReplaceKeyword_Click", vbExclamation
End Sub

Private Sub ReplaceKeywordSynonym()
    Dim strKeyword      As String
    
'***************************************
' Module/Form Name   : frmKeywordClustering
'
' Procedure Name     : ReplaceKeywordSynonym
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
    Dim oKeywordSynonym     As KeywordSynonym
    
    Set oKeywordSynonym = moKeywordSynonyms.Item(Trim(txtKeyword.Text) & "/" & Trim(smgKeywordClustering.Column(2).Value))
    oKeywordSynonym.PendingSynonym = txtSynonym.Text
    moKeywordSynonyms.update
    DisplayClustering txtKeyword.Text, txtSynonym.Text
    '
    SetKeywordUpdateButtons
    '
    On Error Resume Next
    txtKeyword.SetFocus
'
'********* Code Ends Here **************
'
    On Error GoTo 0
    Exit Sub
'
ErrorHandler:
    ErrorRaise "frmKeywordClustering.ReplaceKeywordSynonym"
End Sub

Private Sub Form_Initialize()
    mblnUnhandleEvents = False
End Sub

Private Sub Form_Load()
    
    Set smgKeywordClustering.Container = fraKeywordSynonyms
    With smgKeywordClustering
        .Height = 2775
        .Left = 120
        .Top = 840
        .Width = 7850
    End With

End Sub

Private Sub smgKeywordClustering_GotFocus()
    ''smgKeywordClustering.CurrentRow = smgKeywordClustering.TopRow
End Sub

Private Sub smgKeywordClustering_KeyPress(KeyAscii As Integer)
    txtKeyword.Text = Chr(KeyAscii)
    txtKeyword.SetFocus
End Sub

Private Sub smgKeywordClustering_RowChanged(CurrentRow As String)
    mblnUnhandleEvents = True
    
    txtKeyword.Text = smgKeywordClustering.Column(1).Value
    txtSynonym.Text = smgKeywordClustering.Column(2).Value
    
    SetKeywordUpdateButtons
    mblnUnhandleEvents = False
End Sub

Private Sub txtKeyword_Change()
'***************************************
' Module/Form Name   : frmKeywordClustering
'
' Procedure Name     : txtKeyword_Change
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
    If mblnUnhandleEvents Then Exit Sub
    
    SetControlsState
'
'********* Code Ends Here **************
'
    On Error GoTo 0
    Exit Sub
'
ErrorHandler:
    DisplayError , "frmKeywordClustering.txtKeyword_Change", vbExclamation
End Sub

Private Sub SetControlsState()
    
'***************************************
' Module/Form Name   : frmKeywordClustering
'
' Procedure Name     : SetControlsState
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
    Dim intIndex            As Integer
    Dim intCurrentKeyword   As Integer
    Dim blnExactMatch       As Boolean
    Dim strKeyword          As String
    '
    '   Read through the Keywords collection until we find an exact match or
    '   we've gone past it.
    '
    LockWindow Me.hWnd
    blnExactMatch = False
    If Trim(txtKeyword.Text) = "" Then
        txtSynonym.Text = ""
        intCurrentKeyword = 1
    Else
        intCurrentKeyword = 1
        For intIndex = 1 To moKeywordSynonyms.Count
            strKeyword = moKeywordSynonyms.Item(intIndex).Keyword
            If LCase(strKeyword) = LCase(Trim(txtKeyword.Text)) Then
                blnExactMatch = True
                Exit For
            ElseIf LCase(strKeyword) > LCase(Trim(txtKeyword.Text)) Then
                Exit For
            End If
        Next intIndex
    End If
    '
    '   Ensure that the row is not problematic.
    '
    intCurrentKeyword = intIndex - 1
    If intCurrentKeyword <= 0 Then
        intCurrentKeyword = 1
    End If
    If intCurrentKeyword > smgKeywordClustering.Rows Then
        intCurrentKeyword = smgKeywordClustering.Rows
    End If
    '
    '   Set the location of the grid.
    '
    If blnExactMatch Then
        smgKeywordClustering.GetKeyRow txtKeyword.Text
        mblnUnhandleEvents = True
        txtSynonym.Text = smgKeywordClustering.Column(2).Value
        mblnUnhandleEvents = False
        mblnKeywordFound = True
    Else
        smgKeywordClustering.CurrentRow = intCurrentKeyword
        If intCurrentKeyword <> 0 Then
            smgKeywordClustering.TopRow = intCurrentKeyword
        End If
        smgKeywordClustering.Deselect
        mblnKeywordFound = False
    End If
    '
    SetKeywordUpdateButtons
    UnlockWindow
'
'********* Code Ends Here **************
'
    On Error GoTo 0
    Exit Sub
'
ErrorHandler:
    ErrorRaise "frmKeywordClustering.SetControlsState"
End Sub

Private Sub SetKeywordUpdateButtons()
'***************************************
' Module/Form Name   : frmKeywordClustering
'
' Procedure Name     : SetKeywordUpdateButtons
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
    If Trim(txtKeyword.Text) <> "" And Trim(txtSynonym.Text) <> "" Then
        If LCase(Trim(txtKeyword.Text)) = LCase(smgKeywordClustering.Column(1).Value) Then
            cmdDeleteKeyword.Enabled = True
            If LCase(txtSynonym.Text) = LCase(smgKeywordClustering.Column(2).Value) Then
                cmdNewKeyword.Enabled = False
                cmdReplaceKeyword.Enabled = False
                cmdNewKeyword.Default = True
            Else
                cmdNewKeyword.Enabled = True
                cmdReplaceKeyword.Enabled = True
''                cmdReplaceKeyword.Default = True
                cmdNewKeyword.Default = True
            End If
        Else
            cmdNewKeyword.Enabled = True
            cmdNewKeyword.Default = True
            cmdReplaceKeyword.Enabled = False
            cmdDeleteKeyword.Enabled = False
        End If
    Else
        cmdNewKeyword.Enabled = False
    End If
'
'********* Code Ends Here **************
'
    On Error GoTo 0
    Exit Sub
'
ErrorHandler:
    ErrorRaise "frmKeywordClustering.SetKeywordUpdateButtons"
End Sub

Private Sub txtKeyword_GotFocus()
    HighLightText txtKeyword
End Sub

Private Sub txtSynonym_Change()
    If mblnUnhandleEvents Then Exit Sub
    
    SetKeywordUpdateButtons
End Sub

Private Sub txtSynonym_GotFocus()
    HighLightText txtSynonym
End Sub
