VERSION 5.00
Begin VB.Form frmKeywordExclusions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maintain Keyword Exclusions"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4245
   Icon            =   "KeywordExclusions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   4245
   StartUpPosition =   1  'CenterOwner
   Begin ImageWhere.SimpleGrid smgKeywordsExcluded 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   7435
      Columns         =   1
      KeyCol          =   0
   End
   Begin VB.CommandButton cmdNewKeyword 
      Caption         =   "&New..."
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdDeleteKeyword 
      Caption         =   "&Delete"
      Height          =   315
      Left            =   3000
      TabIndex        =   2
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   555
      Left            =   3300
      Picture         =   "KeywordExclusions.frx":0742
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5040
      UseMaskColor    =   -1  'True
      Width           =   795
   End
   Begin VB.Menu mnuKeywordExclusions 
      Caption         =   "Keyword Exclusions Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
      End
   End
End
Attribute VB_Name = "frmKeywordExclusions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private moKeywordExclusions As KeywordExclusions

Private Sub cmdClose_Click()
    Unload Me
End Sub

Public Sub Display()
    Set moKeywordExclusions = New KeywordExclusions
    moKeywordExclusions.Refresh
    '
    '   Initialise the grid
    '
    With smgKeywordsExcluded
        .Columns = 1
        .Column(1).Header = "Excluded Keyword"
        .Column(1).Width = .Width - 364
    End With
    '
    '   Display the keywords.
    '
    DisplayExclusions
    '
    Me.Show vbModal
    
End Sub

Private Sub DisplayExclusions(Optional ByVal pstrKeywordExcluded As String = "")
    Dim oKeywordExcluded    As KeywordExcluded
    
'***************************************
' Module/Form Name   : frmKeywordExclusions
'
' Procedure Name     : DisplayExclusions
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
    smgKeywordsExcluded.Clear
    For Each oKeywordExcluded In moKeywordExclusions
        If oKeywordExcluded.Dirty <> pdgDelete Then
            smgKeywordsExcluded.AddRow False, oKeywordExcluded.Keyword
        End If
    Next oKeywordExcluded
    smgKeywordsExcluded.ResizeRows
    If pstrKeywordExcluded <> "" Then
        smgKeywordsExcluded.GetKeyRow pstrKeywordExcluded
    End If
'
'********* Code Ends Here **************
'
    On Error GoTo 0
    Exit Sub
'
ErrorHandler:
    ErrorRaise "frmKeywordExclusions.DisplayExclusions"
End Sub

Private Sub cmdDeleteKeyword_Click()
    DeleteKeywordExcluded
End Sub

Private Sub cmdNewKeyword_Click()
    NewKeywordExcluded
End Sub

Private Sub NewKeywordExcluded()
    Dim fNewKeywordExcluded As frmNewKeywordExclusion
'***************************************
' Module/Form Name   : frmKeywordExclusions
'
' Procedure Name     : NewKeywordExcluded
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
    Set fNewKeywordExcluded = New frmNewKeywordExclusion
    fNewKeywordExcluded.Display moKeywordExclusions
    If Not fNewKeywordExcluded.Cancel Then
        moKeywordExclusions.Add fNewKeywordExcluded.KeywordExcluded
        moKeywordExclusions.update
        DisplayExclusions fNewKeywordExcluded.KeywordExcluded
    End If
    
    Set fNewKeywordExcluded = Nothing
'
'********* Code Ends Here **************
'
    On Error GoTo 0
    Exit Sub
'
ErrorHandler:
    ErrorRaise "frmKeywordExclusions.NewKeywordExcluded"
End Sub

Private Sub DeleteKeywordExcluded()
    Dim strKeywordExclusion As String
    
'***************************************
' Module/Form Name   : frmKeywordExclusions
'
' Procedure Name     : DeleteKeywordExcluded
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
    strKeywordExclusion = smgKeywordsExcluded.Column(1).Value
    If MsgBox("Are you sure you wish to delete keyword exclusion '" & strKeywordExclusion & "'", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
        Exit Sub
    End If
    
    moKeywordExclusions.RemovePending strKeywordExclusion
    moKeywordExclusions.update
    DisplayExclusions
'
'********* Code Ends Here **************
'
    On Error GoTo 0
    Exit Sub
'
ErrorHandler:
    ErrorRaise "frmKeywordExclusions.DeleteKeywordExcluded"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set moKeywordExclusions = Nothing
End Sub

Private Sub mnuDelete_Click()
'***************************************
' Module/Form Name   : frmKeywordExclusions
'
' Procedure Name     : mnuDelete_Click
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
    DeleteKeywordExcluded
'
'********* Code Ends Here **************
'
    On Error GoTo 0
    Exit Sub
'
ErrorHandler:
    ErrorRaise "frmKeywordExclusions.mnuDelete_Click"
End Sub

Private Sub mnuNew_Click()
'***************************************
' Module/Form Name   : frmKeywordExclusions
'
' Procedure Name     : mnuNew_Click
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
    NewKeywordExcluded
'
'********* Code Ends Here **************
'
    On Error GoTo 0
    Exit Sub
'
ErrorHandler:
    ErrorRaise "frmKeywordExclusions.mnuNew_Click"
End Sub

Private Sub smgKeywordsExcluded_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuKeywordExclusions, vbPopupMenuRightButton
    End If
End Sub
