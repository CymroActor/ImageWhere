VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Log"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private intFileNo   As Integer
Private moFSO       As Scripting.FileSystemObject

Public Sub WriteLog(strDescription As String, strSource As String)
    '
    '   Open the log file.
    '
    intFileNo = FreeFile(1)
    'Open App.Path & "\Image Where.log" For Append As intFileNo
    Open goSystemConfig.ServerLocation & "\Image Where.log" For Append As intFileNo
    
    Write #intFileNo, _
          Format(Now, "dd/mm/yyyy hh:mm:ss"), _
          strDescription, _
          strSource, _
          goSystemConfig.ComputerName, _
          goSystemConfig.UserName, _
          App.Major & "." & App.Minor & "." & App.Revision
    '
    Close intFileNo
End Sub

Public Sub WriteDetailedLog(strDescription As String, strSource As String)
    If Not goSystemConfig.DetailedAuditing Then
        Exit Sub
    End If
    '
    '   Open the log file.
    '
    intFileNo = FreeFile(1)
    'Open App.Path & "\Image Where.log" For Append As intFileNo
    Open goSystemConfig.ServerLocation & "\Image Where.log" For Append As intFileNo

    Write #intFileNo, _
            Format(Now, "dd/mm/yyyy hh:mm:ss"), _
            strDescription, _
            strSource, _
            goSystemConfig.ComputerName, _
            goSystemConfig.UserName, _
            App.Major & "." & App.Minor & "." & App.Revision
    '
    Close intFileNo
End Sub

Public Sub PostLogFile()
'***************************************
' Module/Form Name   : Log
'
' Procedure Name     : PostLogFile
'
' Purpose            :
'
' Date Created       : 02/06/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 02/06/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo PostLogFile_Error
'
'******** Code Starts Here *************
'
    Dim colFiles As Collection
    '
    '   Ignore if not a support user.
    '
    If Not goSystemConfig.SupportUser Then
        Exit Sub
    End If
    '
    '   Need to refresh the System Config object.
    '
    goSystemConfig.Refresh
    '
    '   Send Log Files to the Support email address.
    '
    Set colFiles = Nothing
    Set colFiles = New Collection
    colFiles.Add moFSO.BuildPath(gstrAppPath, "\Image Where.log")
    colFiles.Add moFSO.BuildPath(gstrUpgradePath, "\Upgrade.log")
    '
    '   Do we need to send the HTML Email Files?
    '
    If goSystemConfig.HTMLEmailSupportRequestDate < Now Then
        If goSystemConfig.DateHTMLEmailPosted < goSystemConfig.HTMLEmailSupportRequestDate Then
            OutputEmails
            colFiles.Add moFSO.BuildPath(gstrAppPath, "Emails\PostedHTMLEmailOutput.htm")
            colFiles.Add moFSO.BuildPath(gstrAppPath, "Emails\ConfirmationHTMLEmailOutput.htm")
            colFiles.Add moFSO.BuildPath(gstrAppPath, "Emails\CancellationHTMLEmailOutput.htm")
            goSystemConfig.DateHTMLEmailPosted = Now
        End If
    End If
    '
    '   Do we need to execute some diagnostic SQL?
    '
    If goSystemConfig.DiagnosticsRequestDate < Now Then
        If goSystemConfig.DateDiagnosticsPosted < goSystemConfig.DiagnosticsRequestDate Then
            ExecuteDiagnosticSQL
            colFiles.Add moFSO.BuildPath(gstrAppPath, "Diagnostics.txt")
            goSystemConfig.DateDiagnosticsPosted = Now
        End If
    End If
    '
    If Not Email(goSystemConfig.SupportEmail, _
                 "Image Where Log: " & goCompanyInfo.CompanyName, _
                 "", _
                 colFiles) Then
        Exit Sub
    End If
    '
    '   Update the System Config table.
    '
    With goSystemConfig
        .DateLogFilePosted = Now
        .Update
    End With
    '
    '   Decide whether to checkpoint the log files.
    '
    CheckPoint moFSO.BuildPath(gstrAppPath, "Image Where.log")
    CheckPoint moFSO.BuildPath(gstrUpgradePath, "Upgrade.log")
    '
    '********* Code Ends Here **************
    '
Exit Sub
    '
PostLogFile_Error:
    '
    '   Simply write any errors to the log file.
    '
    WriteLog Err.Number & " - " & Err.Description & " at line: " & CStr(Erl), "Log.PostLogFile"
End Sub

Private Sub ExecuteDiagnosticSQL()
'***************************************
' Module/Form Name   : Log
'
' Procedure Name     : ExecuteDiagnosticSQL
'
' Purpose            :
'
' Date Created       : 08/07/2004
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo ExecuteDiagnosticSQL_Error
'
'******** Code Starts Here *************
'
    Dim intFileNo           As Integer
    Dim rstDiagnostics      As DAO.Recordset
    Dim strOutput           As String
    Dim fld                 As Field
    
    intFileNo = FreeFile
    '
    '   Generate the Diagnostics recordset.
    '
    Set rstDiagnostics = db.OpenRecordset(goSystemConfig.DiagnosticSQL, dbOpenSnapshot)
    '
    '   Output the Query.
    '
    strOutput = "Query: " & goSystemConfig.DiagnosticSQL
    Open moFSO.BuildPath(gstrAppPath, "Diagnostics.txt") For Output As intFileNo
    Print #intFileNo, strOutput
    Close #1
    '
    '   Output the headings.
    '
    If rstDiagnostics.EOF Then
        strOutput = "No Data Returned"
        Open moFSO.BuildPath(gstrAppPath, "Diagnostics.txt") For Append As intFileNo
        Print #intFileNo, strOutput
        Close #1
    Else
        strOutput = """"
        For Each fld In rstDiagnostics.Fields
            strOutput = strOutput & fld.Name & ""","""
        Next fld
        strOutput = Left(strOutput, Len(strOutput) - 2)
        '
        Open moFSO.BuildPath(gstrAppPath, "Diagnostics.txt") For Append As intFileNo
        Print #intFileNo, strOutput
        Close #1
    End If
    '
    While Not rstDiagnostics.EOF
        strOutput = ""
        For Each fld In rstDiagnostics.Fields
            Select Case fld.Type
            Case Is = dbText, dbMemo
                strOutput = strOutput & ",""" & NVL(fld.Value, NVLString) & """"
            Case Else
                strOutput = strOutput & "," & CStr(NVL(fld.Value, NVLNumeric))
            End Select
        Next fld
        strOutput = Mid(strOutput, 2)
        Open moFSO.BuildPath(gstrAppPath, "Diagnostics.txt") For Append As intFileNo
        Print #intFileNo, strOutput
        Close #1
        '
        rstDiagnostics.MoveNext
    Wend
    '
    rstDiagnostics.Close
    Set rstDiagnostics = Nothing
    '
'
'********* Code Ends Here **************
'
   Exit Sub
'
ExecuteDiagnosticSQL_Error:
    strOutput = Err.Description
    Open gstrAppPath & "\Diagnostics.txt" For Output As intFileNo
    Write #intFileNo, strOutput
    Close #1
End Sub

Private Sub CheckPoint(strFile As String)
'***************************************
' Module/Form Name   : Log
'
' Procedure Name     : Checkpoint
'
' Purpose            :
'
' Date Created       : 26/02/2003
'
' Author             : GARETH SAUNDERS
'
' Parameters         : strFile - String
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo Checkpoint_Error
'
'******** Code Starts Here *************
'
    Dim oFSO As Scripting.FileSystemObject
    Dim oFile As Scripting.File
    Dim strFileBackup As String
    
    Set oFSO = New Scripting.FileSystemObject
    If oFSO.FileExists(strFile) Then
        Set oFile = oFSO.GetFile(strFile)
        If oFile.Size > 100000 Then
            Set oFile = Nothing
            strFileBackup = Left(strFile, InStr(strFile, ".") - 1) & "(" & Format(Now, "yyyy-mm-dd hh-Nn-ss") & ").log"
            If oFSO.FileExists(strFileBackup) Then
                oFSO.DeleteFile (strFileBackup)
            End If
            oFSO.MoveFile strFile, strFileBackup
        Else
            Set oFile = Nothing
        End If
    End If
'
'********* Code Ends Here **************
'
   Exit Sub
'
Checkpoint_Error:
    ErrorRaise "Log.Checkpoint"
End Sub

Private Sub OutputEmails()
'***************************************
' Module/Form Name   : Log
'
' Procedure Name     : OutputEmails
'
' Purpose            :
'
' Date Created       : 27/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo OutputEmails_Error
'
'******** Code Starts Here *************
'
    Dim intFileNo As Integer
    '
    '   Posted Email.
    '
    intFileNo = FreeFile
    Open gstrAppPath & "\Emails\PostedHTMLEmailOutput.htm" For Output As intFileNo
    Print #intFileNo, goSystemConfig.PostedHTMLEmail
    Close #1
    '
    '   Confirmation Email.
    '
    intFileNo = FreeFile
    Open gstrAppPath & "\Emails\ConfirmationHTMLEmailOutput.htm" For Output As intFileNo
    Print #intFileNo, goSystemConfig.ConfirmationHTMLEmail
    Close #1
    '
    '   Cancellation Email.
    '
    intFileNo = FreeFile
    Open gstrAppPath & "\Emails\CancellationHTMLEmailOutput.htm" For Output As intFileNo
    Print #intFileNo, goSystemConfig.CancellationHTMLEmail
    Close #1
'
'********* Code Ends Here **************
'
   Exit Sub
'
OutputEmails_Error:
    ErrorRaise "Log.OutputEmails"
End Sub

Private Sub Class_Initialize()
    Set moFSO = New Scripting.FileSystemObject
End Sub

Private Sub Class_Terminate()
    Set moFSO = Nothing
End Sub
