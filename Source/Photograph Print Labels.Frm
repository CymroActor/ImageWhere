VERSION 5.00
Begin VB.Form frm_print_photo_labels 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Photograph Labels"
   ClientHeight    =   3135
   ClientLeft      =   3570
   ClientTop       =   2235
   ClientWidth     =   4410
   ControlBox      =   0   'False
   Icon            =   "Photograph Print Labels.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3135
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   555
      Left            =   2100
      Picture         =   "Photograph Print Labels.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2400
      UseMaskColor    =   -1  'True
      Width           =   795
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   555
      Left            =   3420
      Picture         =   "Photograph Print Labels.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2400
      UseMaskColor    =   -1  'True
      Width           =   795
   End
   Begin VB.Frame fra_scanned 
      Caption         =   "&Scanned"
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   2220
      Width           =   1815
      Begin VB.OptionButton exclude_scanned 
         Caption         =   "Exclude "
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.OptionButton include_scanned 
         Caption         =   "Include "
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   540
         Width           =   1215
      End
   End
   Begin VB.Frame fra_photo_labels 
      Caption         =   "&Photograph Label Ranges"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   540
      Width           =   4155
      Begin VB.TextBox txt_photo_label_range 
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   3735
      End
      Begin VB.Label lab_guidelines 
         Caption         =   $"Photograph Print Labels.frx":0E16
         Height          =   675
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   3735
      End
   End
   Begin VB.Image img_labels 
      Height          =   480
      Left            =   2100
      Picture         =   "Photograph Print Labels.frx":0EA3
      Top             =   60
      Width           =   480
   End
   Begin VB.Image img_photograph 
      Height          =   480
      Left            =   1560
      Picture         =   "Photograph Print Labels.frx":11AD
      Top             =   60
      Width           =   480
   End
End
Attribute VB_Name = "frm_print_photo_labels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private range(30, 2)

Private Sub cmdCancel_Click()
    Me.Enabled = False
    Unload Me
    Set frm_print_photo_labels = Nothing
End Sub

Private Sub cmdOK_Click()
'***************************************
' Module/Form Name   : frm_print_photo_labels
'
' Procedure Name     : cmdOK_Click
'
' Purpose            :
'
' Date Created       : 18/12/2002
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
    Screen.MousePointer = vbHourglass

    txt_photo_label_range.SetFocus

    Photograph_label.print_ranges txt_photo_label_range, _
                                  include_scanned, _
                                  abort
    Screen.MousePointer = vbDefault
'
'********* Code Ends Here **************
'
   Exit Sub
'
cmdOK_Click_Error:
    DisplayError , "frm_print_photo_labels.cmdOK_Click", vbExclamation
End Sub

Public Sub load_form(photograph_no As String)
    txt_photo_label_range = photograph_no
    Show
End Sub
