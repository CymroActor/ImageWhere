VERSION 5.00
Begin VB.Form frmJobReceiveFeedBack 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Receive Web Job Feedback"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4770
   Icon            =   "frmJobReceiveFeedBack.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   4770
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraJob 
      Caption         =   "Job"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.OptionButton optRefined 
         Caption         =   "Reviewed and Confirmed"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Value           =   -1  'True
         Width           =   3255
      End
      Begin VB.OptionButton optCancelled 
         Caption         =   "Cancelled"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox txtNoPhotosRequested 
         Height          =   405
         Left            =   3000
         MaxLength       =   6
         TabIndex        =   4
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblNoPhotosRequested 
         Caption         =   "No of Photographs Requested: (refer to automated email)"
         Height          =   495
         Left            =   480
         TabIndex        =   3
         Top             =   1080
         Width           =   2295
      End
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   555
      Left            =   2880
      Picture         =   "frmJobReceiveFeedBack.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1920
      UseMaskColor    =   -1  'True
      Width           =   795
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   555
      Left            =   3840
      Picture         =   "frmJobReceiveFeedBack.frx":0C84
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1920
      UseMaskColor    =   -1  'True
      Width           =   795
   End
End
Attribute VB_Name = "frmJobReceiveFeedBack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private moJob As Job2

Private Sub cmdCancel_Click()
    Unload Me
    DoEvents
End Sub

Private Sub cmdDone_Click()
'***************************************
' Module/Form Name   : frmJobReceiveFeedBack
'
' Procedure Name     : cmdDone_Click
'
' Purpose            :
'
' Date Created       : 09/06/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 10/06/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo cmdDone_Click_Error
'
'******** Code Starts Here *************
'
    Dim strEmail As String
    Dim intNoPhotosRequested As Integer
    Dim oOutlook As OutlookClass
    Dim oCustomer As Customer2

    If Not ValidEntry Then
        Exit Sub
    End If
    '
    Screen.MousePointer = vbHourglass
    gdbADO.BeginTrans
        '
        '   Record the job as having been received.
        '
        If optCancelled.Value = True Then
            intNoPhotosRequested = 0
        Else
            intNoPhotosRequested = CInt(txtNoPhotosRequested.Text)
        End If
        moJob.ReceiveWebFeedback intNoPhotosRequested
        '
        '   Send an Email.
        '
        If optCancelled.Value = True Then
            If goSystemConfig.WebSearchTestEmail = "" Then
                strEmail = moJob.Contact.Email
            Else
                strEmail = goSystemConfig.WebSearchTestEmail
            End If
            Set oCustomer = New Customer2
            oCustomer.Read moJob.CustomerNo
            Set oOutlook = New OutlookClass
            oOutlook.HTMLEmail strEmail, _
                               oCustomer.CustomerName & "  Job: " & moJob.reference & " (" & moJob.JobNo & ")", _
                               moJob.CancellationHTMLEmail, _
                               goSystemConfig.CancellationImages
            Set oCustomer = Nothing
            If Not oOutlook.MailSent Then
                MsgBox "Email Failed", vbExclamation, App.Title
                gdbADO.RollbackTrans
                Screen.MousePointer = vbDefault
                Exit Sub
            Else
                MsgBox "An email has automatically been sent to confirm cancellation of the customer's job.", vbInformation, App.Title
            End If
        End If
    gdbADO.CommitTrans

    Unload Me
    Screen.MousePointer = vbDefault
'
'********* Code Ends Here **************
'
    Exit Sub
    '
cmdDone_Click_Error:
    DisplayError , "frmJobReceiveFeedBack.cmdDone_Click", vbExclamation
End Sub

Private Sub Form_Paint()
    txtNoPhotosRequested.SetFocus
End Sub

Private Sub optCancelled_Click()
'***************************************
' Module/Form Name   : frmJobReceiveFeedBack
'
' Procedure Name     : optCancelled_Click
'
' Purpose            :
'
' Date Created       : 31/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo optCancelled_Click_Error
'
'******** Code Starts Here *************
'
    EnableControls
'
'********* Code Ends Here **************
'
   Exit Sub
'
optCancelled_Click_Error:
    DisplayError , "frmJobReceiveFeedBack.optCancelled_Click", vbExclamation
End Sub

Private Sub optRefined_Click()
'***************************************
' Module/Form Name   : frmJobReceiveFeedBack
'
' Procedure Name     : optRefined_Click
'
' Purpose            :
'
' Date Created       : 31/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo optRefined_Click_Error
'
'******** Code Starts Here *************
'
    EnableControls
    On Error Resume Next
    txtNoPhotosRequested.SetFocus
'
'********* Code Ends Here **************
'
   Exit Sub
'
optRefined_Click_Error:
    DisplayError , "frmJobReceiveFeedBack.optRefined_Click", vbExclamation
End Sub

Private Sub EnableControls()
    If optRefined.Value = True Then
        txtNoPhotosRequested.Locked = False
        txtNoPhotosRequested.BackColor = vbWindowBackground
    Else
        txtNoPhotosRequested.Locked = True
        txtNoPhotosRequested.BackColor = vbButtonFace
    End If
End Sub

Private Sub txtNoPhotosRequested_KeyPress(KeyAscii As Integer)
    allow_numeric_only (KeyAscii)
End Sub

Public Sub Display(oJob As Job2)
    Set moJob = oJob
    Me.Show vbModal
End Sub

Private Function ValidEntry() As Boolean
'***************************************
' Module/Form Name   : frmJobReceiveFeedBack
'
' Procedure Name     : ValidEntry
'
' Purpose            :
'
' Date Created       : 10/06/2002
'
' Author             : GARETH SAUNDERS
'
' Returns            : Boolean
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 10/06/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo ValidEntry_Error
'
'******** Code Starts Here *************
'
    Dim oContact    As Contact
    
    ValidEntry = False
    If optRefined.Value = True Then
        If Trim(txtNoPhotosRequested.Text) = "" Then
            MsgBox "Please enter the number of photographs requested by the customer.", vbExclamation, App.Title
            txtNoPhotosRequested.SetFocus
            Exit Function
        End If

        If CLng(txtNoPhotosRequested.Text) = 0 Then
            MsgBox "You have indicated that the customer has requested photographs." & vbCrLf & _
                   "You may not therefore enter zero.", vbExclamation, App.Title
            txtNoPhotosRequested.SetFocus
            Exit Function
        End If

        If CLng(txtNoPhotosRequested.Text) > moJob.NoPhotos Then
            MsgBox "There" & IIf(moJob.NoPhotos > 1, " are ", " is ") & CStr(moJob.NoPhotos) & IIf(moJob.NoPhotos > 1, " photographs ", " photograph ") & "on this job." & vbCrLf & _
                   "You may not therefore enter more than this number.", vbExclamation, App.Title
            txtNoPhotosRequested.SetFocus
            Exit Function
        End If
    End If

    If optCancelled Then
        If MsgBox("You have indicated that the Customer has requested the job be cancelled." & vbCrLf & _
                  "Is this correct?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            Exit Function
        End If
    Else
        If MsgBox("You have indicated that the Customer has requested " & txtNoPhotosRequested & IIf(CLng(txtNoPhotosRequested.Text) > 1, " photographs ", " photograph ") & "be delivered." & vbCrLf & _
                  "Is this correct?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            Exit Function
        End If
    End If
    '
    '   Does Valid Contact still exist?
    '
    On Error Resume Next
    Set oContact = moJob.Contact
    Select Case Err.Number
    Case Is = 0, vbObjectError + 5, vbObjectError + 9
        On Error GoTo ValidEntry_Error
    Case Else
        ErrorSave
        On Error GoTo ValidEntry_Error
        ErrorRestore
    End Select
    '
    If oContact Is Nothing Then
        MsgBox "Contact on Job no longer exists! No confirmation email can be sent.", vbExclamation, App.Title
        cmdCancel.SetFocus
        Exit Function
    Else
        If oContact.WebUser = False Then
            MsgBox "Contact on Job is no longer a Web User! No confirmation email can be sent.", vbExclamation, App.Title
            cmdCancel.SetFocus
            Exit Function
        End If
    End If
    '
    Set oContact = Nothing
    ValidEntry = True
'
'********* Code Ends Here **************
'
    Exit Function
    '
ValidEntry_Error:
    ErrorRaise "frmJobReceiveFeedBack.ValidEntry"
End Function
