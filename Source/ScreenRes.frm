VERSION 5.00
Begin VB.Form frmScreenRes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Screen Res"
   ClientHeight    =   9255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7500
   Icon            =   "ScreenRes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9255
   ScaleWidth      =   7500
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCaption 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   7680
      Width           =   7215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   555
      Left            =   6600
      Picture         =   "ScreenRes.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8640
      UseMaskColor    =   -1  'True
      Width           =   795
   End
   Begin VB.Image imgPhotograph 
      Height          =   7260
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   7260
   End
End
Attribute VB_Name = "frmScreenRes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Public Sub Display(ByRef poDigitalImage As DigitalImage)
'***************************************
' Module/Form Name   : frmScreenRes
'
' Procedure Name     : Display
'
' Purpose            :
'
' Date Created       : 28/06/2004
'
' Author             : GARETH SAUNDERS
'
' Parameters         : pstrImage - String
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
    
    imgPhotograph.Picture = LoadPicture(poDigitalImage.FileLocation("SC"))
    With poDigitalImage
        .MaxHeight = txtCaption.Top - imgPhotograph.Top - 200
        .MaxWidth = Me.Width - (imgPhotograph.Left * 2)
        Set .Picture = imgPhotograph.Picture
        imgPhotograph.Visible = False
        imgPhotograph.Top = 200
        imgPhotograph.Width = .Width
        imgPhotograph.Height = .Height
        imgPhotograph.Left = (Me.Width - imgPhotograph.Width) / 2
        imgPhotograph.Visible = True
        '
        '   Set the Caption.
        '
        txtCaption.Alignment = vbCenter
        txtCaption.Text = poDigitalImage.Description
    End With
    '
    '   Now we can possibly resize the form.
    '
    txtCaption.Top = imgPhotograph.Top + imgPhotograph.Height + 200
    cmdClose.Top = txtCaption.Top + txtCaption.Height + 200
    Me.Height = cmdClose.Top + cmdClose.Height + 200 + goSystemConfig.TitleBarHeight
    '
    Me.Show vbModal
'
'********* Code Ends Here **************
'
   Exit Sub
'
Display_Error:
    DisplayError , "frmScreenRes.Display", vbExclamation
End Sub
