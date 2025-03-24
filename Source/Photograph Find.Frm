VERSION 5.00
Begin VB.Form frm_photograph_find 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find Photographs"
   ClientHeight    =   2235
   ClientLeft      =   1920
   ClientTop       =   1440
   ClientWidth     =   7245
   Icon            =   "Photograph Find.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2235
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraFindPhotographs 
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   7030
      Begin VB.OptionButton optCaption 
         Caption         =   "Caption"
         Height          =   255
         Left            =   5640
         TabIndex        =   3
         Top             =   255
         Value           =   -1  'True
         Width           =   1200
      End
      Begin VB.OptionButton optKeywords 
         Caption         =   "Keywords"
         Height          =   255
         Left            =   5640
         TabIndex        =   4
         Top             =   555
         Width           =   1200
      End
      Begin VB.TextBox txt_description 
         Height          =   555
         Left            =   1140
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label lab_description 
         Caption         =   "Description:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   300
         Width           =   855
      End
   End
   Begin VB.CommandButton ssc_cancel 
      Caption         =   "&Cancel"
      Height          =   555
      Left            =   6360
      Picture         =   "Photograph Find.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1560
      UseMaskColor    =   -1  'True
      Width           =   795
   End
   Begin VB.CommandButton ssc_done 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   555
      Left            =   5520
      Picture         =   "Photograph Find.frx":089C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1560
      UseMaskColor    =   -1  'True
      Width           =   795
   End
   Begin VB.Image img_find_photograph 
      Height          =   480
      Left            =   120
      Picture         =   "Photograph Find.frx":0E16
      Top             =   180
      Width           =   480
   End
   Begin VB.Label lab_title_lab 
      Caption         =   "Please enter the string you wish to find ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      Top             =   180
      Width           =   4215
   End
End
Attribute VB_Name = "frm_photograph_find"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private mod_dat_batch As Data
Private snap_batch As DAO.Recordset

Public Sub display_photograph_find(search_string)
    txt_description = search_string
    Me.Show 1
End Sub

Private Sub Form_Load()
    txt_description.Text = GetSetting(App.Title, "Find Photograph", "Description", "")
    optCaption.Value = IIf(GetSetting(App.Title, "Find Photograph", "Caption", "Y") = "Y", True, False)
    optKeywords.Value = Not optCaption.Value
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting App.Title, "Find Photograph", "Description", txt_description.Text
    SaveSetting App.Title, "Find Photograph", "Caption", IIf(optCaption.Value, "Y", "N")
End Sub

Private Sub ssc_cancel_Click()
    Unload Me
    Set frm_photograph_find = Nothing
End Sub

Private Sub ssc_done_Click()
'***************************************
' Module/Form Name   : frm_photograph_find
'
' Procedure Name     : ssc_done_Click
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
On Error GoTo ssc_done_Click_Error
'
'******** Code Starts Here *************
'
    Dim start As Long
    Dim finish As Long
    Dim SQL As String
    
    If txt_description = "" Then
        MsgBox "Enter a string to find", vbExclamation
        Exit Sub
    End If
    
    'start = OSTimeGetTime
    
    frm_photograph_find_maint.search_string = txt_description
    photograph.search txt_description.Text, snap_batch, dbOpenSnapshot, abort, optKeywords.Value
    If abort Then
'        MsgBox "No photographs found", vbInformation, App.Title
        Exit Sub
    End If
    
    Screen.MousePointer = ccDefault
    DoEvents
    
    'MsgBox "It took " & (finish - start) / 1000 & " seconds!", 48
    
    If snap_batch.RecordCount = 0 Then
        MsgBox "No photographs found", vbInformation
        txt_description.SetFocus
        Exit Sub
    End If
    
    frm_photograph_find_maint.display_find_results snap_batch, txt_description.Text
    
    Unload Me
    Set frm_photograph_find = Nothing

    frm_photograph_find_maint.Visible = True
    frm_photograph_find_maint.SetFocus
'
'********* Code Ends Here **************
'
   Exit Sub
'
ssc_done_Click_Error:
    DisplayError , "frm_photograph_find.ssc_done_Click", vbExclamation
End Sub


