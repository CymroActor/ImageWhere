VERSION 5.00
Begin VB.Form frmPreparedFiles 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Prepared Digital Images"
   ClientHeight    =   3075
   ClientLeft      =   2355
   ClientTop       =   2385
   ClientWidth     =   7710
   Icon            =   "frmPreparedFiles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtSearches 
      BackColor       =   &H8000000F&
      Height          =   1695
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "frmPreparedFiles.frx":0442
      Top             =   840
      Width           =   7455
   End
   Begin VB.CommandButton cmdOpenFolder 
      Caption         =   "Open &Folder"
      Height          =   375
      Left            =   6360
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4920
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label lblMessage 
      Caption         =   "Label1"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   7455
   End
End
Attribute VB_Name = "frmPreparedFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mstrDigitalSearchesFolder   As String

Public Property Get DigitalSearchesFolder() As String
    DigitalSearchesFolder = mstrDigitalSearchesFolder
End Property

Public Property Let DigitalSearchesFolder(ByVal vNewValue As String)
    mstrDigitalSearchesFolder = vNewValue
End Property

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub cmdOpenFolder_Click()
    Shell "Explorer " & mstrDigitalSearchesFolder, vbNormalFocus
End Sub
