VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OutlookClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Private moOutlook As Outlook.Application
Private moOutlook As Object
'Private WithEvents moMailItem As Outlook.MailItem
Private moMailItem As Object

Private mvarMailSent As Boolean

Public Property Get MailSent() As Boolean
10        MailSent = mvarMailSent
End Property
    
Public Function HTMLEmail(ByVal strTo As String, _
                          ByVal strSubject As String, _
                          ByVal strBody As String, _
                          Optional ByRef pcolAttachments As Collection = Nothing) As Boolean
'***************************************
' Module/Form Name   : OutlookClass
'
' Procedure Name     : HTMLEmail
'
' Purpose            :
'
' Date Created       : 18/10/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : strTo - String
'                    : strSubject - String
'                    : strBody - String
'                    : colImages - Collection
'
' Returns            : Boolean
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 10/11/2002 GARETH SAUNDERS
'                    : 24/11/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo HTMLEmail_Error
'
'******** Code Starts Here *************
'
    Dim colAttach           As Object
    Dim oMailAttach         As Object
    Dim oSession            As MAPI.Session
    Dim oMAPIMsg            As MAPI.Message
    Dim oMAPIAttachs        As MAPI.Attachments
    Dim oMAPIAttach         As MAPI.Attachment
    Dim colFields           As MAPI.Fields
    Dim oField              As MAPI.Field
    Dim strEntryID          As String
    Dim strImage            As Variant
    Dim intCount            As Integer
    Dim oRecipient          As Object
'    Dim oAttachment         As Outlook.Attachment
    Dim varAttachment       As Variant

    HTMLEmail = False
    'Create and open a new contact form for input.
    Set moMailItem = moOutlook.CreateItem(0)
    '
    '   The email is now ready to receive the HTML Body which should already reference these images.
    '
    With moMailItem
        .To = strTo
        .Subject = strSubject
        .HTMLBody = strBody
        Set oRecipient = .ReplyRecipients.Add("pictures@clivenichols.com")
        If Not pcolAttachments Is Nothing Then
            For Each varAttachment In pcolAttachments
                .Attachments.Add varAttachment, 1, 1
            Next varAttachment
        End If
        .Close (2)
        .Display vbModal
    End With
    '
    ' clean up objects
    '
    Set moMailItem = Nothing

    HTMLEmail = True
    '
    '********* Code Ends Here **************
    '
    Exit Function
    '
HTMLEmail_Error:
    ErrorRaise "common_routines.HTMLEmail"
End Function

Private Sub Class_Terminate()
'          Dim oCtl As Office.CommandBarControl
          Dim oCtl As Object
          '
          '   Send the mail immediately.
          '
10        On Error Resume Next
20        If mvarMailSent Then
30            Set oCtl = moOutlook.ActiveExplorer.CommandBars.FindControl(ID:=5488)
40            oCtl.Execute
50        End If
60        Set moOutlook = Nothing
End Sub

Private Sub moMailItem_Send(Cancel As Boolean)
10        mvarMailSent = True
End Sub

Private Sub Class_Initialize()
10        mvarMailSent = False
          'Create a Microsoft Outlook object.
20        'Set moOutlook = New Outlook.Application
30        Set moOutlook = CreateObject("Outlook.Application")
End Sub
