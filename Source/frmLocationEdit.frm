VERSION 5.00
Begin VB.Form frmLocationEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Location"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7140
   Icon            =   "frmLocationEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdBrowsePrefix 
      Caption         =   "&Browse..."
      Height          =   375
      Left            =   6000
      TabIndex        =   10
      Top             =   840
      Width           =   915
   End
   Begin VB.TextBox txtSuffix 
      Height          =   315
      Left            =   1560
      MaxLength       =   10
      TabIndex        =   7
      Top             =   1560
      Width           =   1995
   End
   Begin VB.TextBox txtPrefix 
      Height          =   675
      Left            =   1560
      MaxLength       =   100
      TabIndex        =   5
      Top             =   840
      Width           =   4275
   End
   Begin VB.TextBox txtID 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1560
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   1
      Top             =   120
      Width           =   555
   End
   Begin VB.TextBox txtDescription 
      Height          =   315
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   3
      Top             =   480
      Width           =   4275
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   555
      Left            =   5280
      Picture         =   "frmLocationEdit.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2040
      UseMaskColor    =   -1  'True
      Width           =   795
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   555
      Left            =   6240
      Picture         =   "frmLocationEdit.frx":0586
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2040
      UseMaskColor    =   -1  'True
      Width           =   795
   End
   Begin VB.Label lblSuffix 
      Caption         =   "&Suffix:"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Width           =   1155
   End
   Begin VB.Label lblPrefix 
      Caption         =   "&Prefix:"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   1155
   End
   Begin VB.Label lblID 
      Caption         =   "&Name:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1155
   End
   Begin VB.Label lblDescription 
      Caption         =   "&Description:"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   1155
   End
End
Attribute VB_Name = "frmLocationEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event LocationUpdated(Key As String)
Private moLocation As Location

Public Sub Display(ByRef poLocation As Location)
'***************************************
' Module/Form Name   : frmLocationEdit
'
' Procedure Name     : Display
'
' Purpose            :
'
' Date Created       : 23/06/2004
'
' Author             : GARETH SAUNDERS
'
' Parameters         : poLocation - Location
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
    Set moLocation = poLocation
    
    txtID.Text = moLocation.ID
    txtDescription.Text = Trim(moLocation.Description)
    txtPrefix.Text = Trim(moLocation.Prefix)
    txtSuffix.Text = Trim(moLocation.Suffix)
    Me.Show vbModal
'
'********* Code Ends Here **************
'
   Exit Sub
'
Display_Error:
    DisplayError , "frmLocationEdit.Display", vbExclamation
End Sub

Private Sub cmdBrowsePrefix_Click()
    Dim spath               As String
    Dim strPathReturned     As String
    'the path used in the Browse function
    'must be correctly formatted depending
    'on whether the path is a drive, a
    'folder, or "".
     spath = FixPath(txtPrefix.Text)
     
    'call the function, returning the path
    'selected (or "" if cancelled)
    strPathReturned = BrowseForFolderByPath(Me, spath)
    If strPathReturned <> "" Then
         txtPrefix.Text = strPathReturned
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
'***************************************
' Module/Form Name   : frmLocationEdit
'
' Procedure Name     : cmdOK_Click
'
' Purpose            :
'
' Date Created       : 23/06/2004
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
    With moLocation
        .Description = txtDescription.Text
        .Prefix = txtPrefix.Text
        .Suffix = txtSuffix.Text
        .update
        RaiseEvent LocationUpdated(.Key)
    End With
    '
    Unload Me
    DoEvents
'
'********* Code Ends Here **************
'
   Exit Sub
'
cmdOK_Click_Error:
    DisplayError , "frmLocationEdit.cmdOK_Click", vbExclamation
End Sub

Private Sub Command1_Click()

End Sub
