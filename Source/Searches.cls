VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Searches"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mCol As Collection
Private mlngJobNo As Long
''Private mrecSearches As ADODB.Recordset
Private mrecSearches As DAO.Recordset
Private mblnPendingRefresh          As Boolean

Public Property Get JobNo() As Long
    JobNo = mlngJobNo
End Property

Public Property Get PendingRefresh() As Boolean
    PendingRefresh = mblnPendingRefresh
End Property

Public Property Let PendingRefresh(vData As Boolean)
    mblnPendingRefresh = vData
End Property

Public Property Let JobNo(ByVal vNewValue As Long)
  
    mlngJobNo = vNewValue
    Refresh

End Property

Public Sub Refresh()
    '***************************************
    ' Module/Form Name   : Searches
    '
    ' Procedure Name     : Refresh
    '
    ' Purpose            :
    '
    ' Date Created       : 28/04/2002
    '
    ' Author             : GARETH SAUNDERS
    '
    ' Amendment History  : Date       Author    Description
    '                    : --------------------------------
    '                    : 28/04/2002 GARETH SAUNDERS
    '
    '***************************************
    '
    On Error GoTo Refresh_Error
    '
    '******** Code Starts Here *************
    '
    Dim strSQL As String
    Dim objSearch As Search2
    '
    '   Clear all existing Searches.
    '
    Set mCol = Nothing
    Set mCol = New Collection
    '
    '   Refresh Searches Collection.
    '
    strSQL = "SELECT SEARCH.*, "
    strSQL = strSQL & "(SELECT count(*) "
    strSQL = strSQL & "FROM search_result inner join photograph on photograph.photograph_key = search_result.photograph_key "
    strSQL = strSQL & "WHERE search.search_no = search_result.search_no "
    strSQL = strSQL & "and photograph.deleted = false) as NoPhotographs "
    strSQL = strSQL & "FROM SEARCH "
    strSQL = strSQL & "WHERE JOB_NO = " & mlngJobNo & " ORDER BY SEARCH_NO"

    Set mrecSearches = Nothing
''    Set mrecSearches = New ADODB.Recordset
''    mrecSearches.CursorLocation = adUseClient
''    mrecSearches.Open strSQL, gdbADO, adOpenDynamic, adLockBatchOptimistic, adCmdText
''    Set mrecSearches.ActiveConnection = Nothing
    Set mrecSearches = db.OpenRecordset(strSQL, dbOpenSnapshot)
    With mrecSearches
        While Not mrecSearches.EOF
            Set objSearch = Add(!search_no)
            objSearch.JobNo = NVL(!job_no, NVLNumeric)
            objSearch.Description = NVL(!Search_Description, NVLString)
            objSearch.SearchDate = NVL(!date_of_search, NVLNumeric)
            objSearch.NoPhotographs = NVL(!NoPhotographs, NVLNumeric)
            objSearch.PhotoType = NVL(!PhotoType, NVLString)
            .MoveNext
        Wend
    End With
    Set mrecSearches = Nothing
    '
    '********* Code Ends Here **************
    '
    Exit Sub
    '
Refresh_Error:
    ErrorRaise "Searches.Refresh"
End Sub
Public Function Add(SearchNo As Long) As Search2

    'create a new object
    Dim objNewMember As Search2
    Set objNewMember = New Search2

    'set the properties passed into the method
    objNewMember.SearchNo = SearchNo
    mCol.Add objNewMember, CStr(SearchNo)
    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As Search2
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
    mblnPendingRefresh = False
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

Public Sub Delete(lngSearchNo As Long)
'***************************************
' Module/Form Name   : Searches
'
' Procedure Name     : Delete
'
' Purpose            :
'
' Date Created       : 01/12/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : lngSearchNo - Long
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 01/12/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo Delete_Error
'
'******** Code Starts Here *************
'
    Dim oSearch         As Search2
''    Dim rstSearch       As ADODB.Recordset
    Dim rstSearch       As DAO.Recordset
    Dim strSQL          As String
    '
    '   Locate the Search in the ADO recordset.
    '
''    gdbADO.BeginTrans
    begin_trans
''        With mrecSearches
''        strSQL = "SELECT * FROM SEARCH WHERE SEARCH_NO = " & lngSearchNo
        strSQL = "DELETE * FROM SEARCH WHERE SEARCH_NO = " & lngSearchNo
''        Set rstSearch = db.OpenRecordset(strSQL, dbOpenSnapshot)
''        With rstSearch
''            .MoveFirst
''            .Find "Search_No = " & CStr(lngSearchNo), , adSearchForward, 0
            '
            '   Delete the Search from the recordset.
            '
''            !Search_Description = "DELETED"
''            Set .ActiveConnection = gdbADO
''            On Error Resume Next
''            .UpdateBatch adAffectCurrent
''            If Err.Number = -2147217864 Then
''                On Error GoTo Delete_Error
''                Err.Raise vbObjectError + 1, , "Record has changed since first retrieving it."
''            ElseIf Err.Number <> 0 Then
''                ErrorSave
''                On Error GoTo Delete_Error
''                ErrorRestore
''            Else
''                On Error GoTo Delete_Error
''            End If
''            .Delete
''            .UpdateBatch
''            Set .ActiveConnection = Nothing
        db.Execute strSQL
        '
        '   If there is a Transparency Search with this Search No. then reset its Digital Search No.
        '
        strSQL = "UPDATE SEARCH "
        strSQL = strSQL & "SET DIGITALSEARCHNO = NULL "
        strSQL = strSQL & "WHERE DIGITALSEARCHNO = " & CStr(lngSearchNo)
''            gdbADO.Execute strSQL
        db.Execute strSQL
        '
        For Each oSearch In mCol
            If oSearch.DigitalSearchNo = lngSearchNo Then
                oSearch.DigitalSearchNo = ""
            End If
        Next oSearch
''    gdbADO.CommitTrans
    commit_trans
    '
    '   Refresh the collection of Searches.
    '
    On Error Resume Next
    goLog.WriteDetailedLog CStr(mCol.Item(CStr(lngSearchNo)).JobNo) & "/" & CStr(lngSearchNo), "Search Deleted"
    mCol.Remove CStr(lngSearchNo)
    '
    '********* Code Ends Here **************
    '
    Exit Sub
    '
Delete_Error:
    ErrorRaise "Searches.Delete"
End Sub

Public Sub AddObject(ByRef poSearch As Search2)
'***************************************
' Module/Form Name   : Searches
'
' Procedure Name     : AddObject
'
' Purpose            :
'
' Date Created       : 25/06/2004
'
' Author             : ADMINISTRATOR
'
' Parameters         : poSearch - Search2
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo AddObject_Error
'
'******** Code Starts Here *************
'
    mCol.Add poSearch, CStr(poSearch.SearchNo)
'
'********* Code Ends Here **************
'
   Exit Sub
'
AddObject_Error:
    ErrorRaise "Searches.AddObject"
End Sub

