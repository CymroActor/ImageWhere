VERSION 5.00
Begin VB.Form frm_image_new 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create new images"
   ClientHeight    =   2085
   ClientLeft      =   4095
   ClientTop       =   2085
   ClientWidth     =   3420
   ControlBox      =   0   'False
   Icon            =   "Image New.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2085
   ScaleWidth      =   3420
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Cancel"
      Height          =   675
      Left            =   2520
      Picture         =   "Image New.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   675
      Left            =   240
      Picture         =   "Image New.frx":089C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1320
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.TextBox txt_no_of_new_images 
      Height          =   315
      Left            =   2340
      MaxLength       =   2
      TabIndex        =   0
      Top             =   780
      Width           =   975
   End
   Begin VB.TextBox txt_photograph_no 
      BackColor       =   &H8000000F&
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   2340
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Lab_no_of_new_images 
      Caption         =   "Number of new images:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   780
      Width           =   1755
   End
   Begin VB.Label lab_photograph_no 
      Caption         =   "Photograph no:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   1275
   End
End
Attribute VB_Name = "frm_image_new"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    Dim mod_image_no As Long
Private Sub update_loaded_screens()
    
10        If is_form_loaded("frm_photograph_view") Then
20            frm_photograph_view.datPhotograph.Refresh
30        End If
    
40        If is_form_loaded("frm_photograph_edit") Then
50            With frm_photograph_edit
60                .display_fields
70                .redisplay_images
80            End With
90        End If

End Sub

Public Sub load_photograph(photograph_no)
10        txt_photograph_no.Text = photograph_no
20        mod_image_no = photograph_no
30        Me.Show 1
End Sub


Private Sub cmdNew_Click()
10        If Len(txt_no_of_new_images) = 0 Then
20            MsgBox "Enter the number of images to be created", vbExclamation
30            txt_no_of_new_images.SetFocus
40            Exit Sub
50        End If
  
60        photo_image.create_new_images mod_image_no, _
                                             txt_no_of_new_images, _
                                             abort

    
70        update_loaded_screens
80        Unload Me
90        Set frm_image_new = Nothing

End Sub


Private Sub cmdExit_Click()
10        Unload Me
20        Set frm_image_new = Nothing

End Sub
Private Sub txt_no_of_new_images_KeyPress(KeyAscii As Integer)
10        If (KeyAscii < 48 Or KeyAscii > 57) And _
              Not KeyAscii = 8 Then
20            KeyAscii = 0
30            Beep
40        End If
End Sub


