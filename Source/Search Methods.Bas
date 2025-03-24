Attribute VB_Name = "search"
Option Explicit
Private SQL As String
Private dyna_search As DAO.Recordset
Public Function Read(search_no, rsettype) As DAO.Recordset
    
10        SQL = "SELECT * " _
              & "FROM search " _
              & "WHERE search_no = " + Str(search_no) + ";"

20        Set Read = db.OpenRecordset(SQL, rsettype)

End Function

Private Sub close_search_recordsets()
10        On Error GoTo error_proc
    
20        dyna_search.Close
    
30        Exit Sub
    
error_proc:
40        Select Case Err.Number
              Case Is = 91
50                Resume Next
60            Case Is = 3420
70                Resume Next
80            Case Else
90                MsgBox Str(Err.Number) + " - " + Err.Description, vbCritical
100       End Select
End Sub
''
''Public Sub Delete(search_no As Long)
''    '***************************************
''    ' Module/Form Name   : search
''    '
''    ' Procedure Name     : delete
''    '
''    ' Purpose            :
''    '
''    ' Date Created       : 22/04/2002
''    '
''    ' Author             : GARETH SAUNDERS
''    '
''    ' Parameters         : search_no - Long
''    '
''    ' Amendment History  : Date       Author    Description
''    '                    : --------------------------------
''    '                    : 22/04/2002 GARETH SAUNDERS
''    '
''    '***************************************
''    '
''    On Error GoTo Delete_Error
''    '
''    '******** Code Starts Here *************
''    '
''
''    Dim response As Integer
''    Dim msg As String
''    Dim style As String
''    Dim snap_search_result As DAO.Recordset
''
''    SQL = "SELECT count(search_result.search_result_key) as no_of_photographs " _
''        & "FROM search_result INNER JOIN photograph ON photograph.photograph_key = search_result.photograph_key " _
''        & "WHERE search_result.search_no = " & CStr(search_no)
''
''    Set snap_search_result = db.OpenRecordset(SQL, dbOpenDynaset)
''
''    If snap_search_result.RecordCount = 0 Then
''        MsgBox "Search not found", vbExclamation, "Delete Search"
''        Exit Sub
''    End If
''
''    If Not snap_search_result!no_of_photographs = 0 Then
''        MsgBox "Search has photographs!" + vbCr + "These must be removed first before deleting the Search", vbExclamation, "Delete Search"
''        Exit Sub
''    End If
''
''    Set dyna_search = Read(search_no, dbOpenDynaset)
''
''    With dyna_search
''        .Edit
''        .Delete
''    End With
''
''    close_search_recordsets
''    '
''    '********* Code Ends Here **************
''    '
''    Exit Sub
''    '
''Delete_Error:
''    ErrorRaise "search.delete"
''End Sub
''
''
