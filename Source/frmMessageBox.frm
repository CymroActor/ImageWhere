VERSION 5.00
Begin VB.Form frmMessageBox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Image Where"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6360
   Icon            =   "frmMessageBox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   975
      Left            =   4080
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmMessageBox.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      UseMaskColor    =   -1  'True
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      FillColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   2400
      Picture         =   "frmMessageBox.frx":0F9E
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmMessageBox.frx":12A8
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmMessageBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

