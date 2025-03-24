VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SearchResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mCol As Collection
Private mlngSearchNo As Long
Private msrType As enmRefreshType
Public Enum enmRefreshType
    srAll
    srDuplicates
End Enum

Public Property Get SearchNo(srType As enmRefreshType) As Long
    SearchNo = mlngSearchNo
    srType = msrType
End Property

Public Property Let SearchNo(srType As enmRefreshType, ByVal vNewValue As Long)
    mlngSearchNo = vNewValue
    msrType = srType
    If srType = srAll Then
        Refresh
    End If
End Property

Public Sub Refresh(Optional ByVal plngSearchResultKey As Long = 0, _
                   Optional ByVal plngBatchNo As Long = 0, _
                   Optional ByVal plngDNoteNo As Long = 0)
'***************************************
' Module/Form Name   : SearchResults
'
' Procedure Name     : Refresh
'
' Purpose            :
'
' Date Created       : 29/11/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 29/11/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo Refresh_Error
'
'******** Code Starts Here *************
'
    Dim strSQL              As String
''    Dim recSearchResults    As ADODB.Recordset
    Dim recSearchResults    As DAO.Recordset
    Dim objSearchResult     As SearchResult
    Dim strWhereClause      As String
    '
    '   Clear all existing Searches.
    '
    Set mCol = Nothing
    Set mCol = New Collection
    '
    '   Set the Where clause.
    '
    strWhereClause = ""
    If plngSearchResultKey <> 0 Then
        strWhereClause = " WHERE SEARCH_RESULT.SEARCH_RESULT_KEY = " & CStr(plngSearchResultKey) & " "
    End If
    '
    If plngBatchNo <> 0 Then
        strWhereClause = " WHERE SEARCH_RESULT.BATCH_NO = " & CStr(plngBatchNo) & " "
    End If
    '
    If plngDNoteNo <> 0 Then
        strWhereClause = " WHERE SEARCH_RESULT.DELIVERY_NOTE_NO = " & CStr(plngDNoteNo) & " "
    End If
    '
    '   Refresh Search Results Collection.
    '
    strSQL = "SELECT SEARCH_RESULT.*, BATCH.DESCRIPTION, BATCH.CREDIT, BATCH.EXCLUSIVE, SEARCH.PHOTOTYPE "
    strSQL = strSQL & "FROM (SEARCH_RESULT INNER JOIN BATCH ON SEARCH_RESULT.BATCH_NO = BATCH.BATCH_NO) "
    strSQL = strSQL & "INNER JOIN SEARCH ON SEARCH_RESULT.SEARCH_NO = SEARCH.SEARCH_NO "
    If strWhereClause = "" Then
        strSQL = strSQL & "WHERE SEARCH_RESULT.SEARCH_NO = " & mlngSearchNo & " "
    Else
        strSQL = strSQL & strWhereClause
    End If
    strSQL = strSQL & "ORDER BY BATCH.BATCH_NO, SEARCH_RESULT.PHOTOGRAPH_NO "
    
''    Set recSearchResults = New ADODB.Recordset
''    recSearchResults.Open strSQL, gdbADO, adOpenStatic, , adCmdText
    Set recSearchResults = db.OpenRecordset(strSQL, dbOpenSnapshot)
    With recSearchResults
        While Not recSearchResults.EOF
            Set objSearchResult = Add(!search_result_key)
            objSearchResult.SearchNo = NVL(!search_no, NVLNumeric)
            objSearchResult.BatchNo = NVL(!batch_no, NVLNumeric)
            objSearchResult.PhotographNo = NVL(!photograph_no, NVLNumeric)
            objSearchResult.PhotoDesc = NVL(!Description, NVLString)
            objSearchResult.DateReturned = NVL(!date_returned, NVLNumeric)
            objSearchResult.BookedInBy = NVL(!booked_in_by, NVLString)
            objSearchResult.Confirmed = !Confirmed & ""
            objSearchResult.Credit = !Credit & ""
            objSearchResult.PhotoType = !PhotoType & ""
            objSearchResult.Exclusive = !Exclusive & ""
            objSearchResult.DigitalStatus = !DigitalStatus & ""
            objSearchResult.CheckPoint
            .MoveNext
        Wend
    End With
    
    recSearchResults.Close
    Set recSearchResults = Nothing
    '
    '********* Code Ends Here **************
    '
    Exit Sub
    '
Refresh_Error:
    ErrorRaise "SearchResults.Refresh"
End Sub

Public Sub AddObject(ByRef poSearchResult As SearchResult)
'***************************************
' Module/Form Name   : SearchResults
'
' Procedure Name     : AddObject
'
' Purpose            :
'
' Date Created       : 24/06/2004
'
' Author             : GARETH SAUNDERS
'
' Parameters         : poSearchResult - SearchResult
'
' Returns            :
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
    mCol.Add poSearchResult, CStr(poSearchResult.SearchResultKey)
'
'********* Code Ends Here **************
'
   Exit Sub
'
AddObject_Error:
    ErrorRaise "SearchResults.AddObject"
End Sub

Public Function Add(SearchResultKey As Long) As SearchResult
'***************************************
' Module/Form Name   : SearchResults
'
' Procedure Name     : Add
'
' Purpose            :
'
' Date Created       : 29/11/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : SearchNo - Long
'
' Returns            : Search2
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 29/11/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo Add_Error
'
'******** Code Starts Here *************
'

    'create a new object
    Dim objNewMember As SearchResult
    Set objNewMember = New SearchResult

    'set the properties passed into the method
    objNewMember.SearchResultKey = SearchResultKey
    mCol.Add objNewMember, CStr(SearchResultKey)
    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing
'
'********* Code Ends Here **************
'
    Exit Function
    '
Add_Error:
    ErrorRaise "SearchResults.Add"
End Function

Public Property Get Item(vntIndexKey As Variant) As SearchResult
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
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

Public Sub DuplicateEntries()
'***************************************
' Module/Form Name   : SearchResults
'
' Procedure Name     : DuplicateEntries
'
' Purpose            :
'
' Date Created       : 13/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo DuplicateEntries_Error
'
'******** Code Starts Here *************
'
    Dim strSQL As String
''    Dim recSearchResults As ADODB.Recordset
    Dim recSearchResults As DAO.Recordset
    Dim objSearchResult As SearchResult
    '
    '   Clear all existing Searches.
    '
    Set mCol = Nothing
    Set mCol = New Collection
    '
    '   Refresh Duplicate Search Results Collection.
    '
    strSQL = "SELECT BATCH.batch_no, BATCH.description, Count(BATCH.batch_no) AS CountOfbatch_no "
    strSQL = strSQL & "FROM SEARCH_RESULT INNER JOIN BATCH ON SEARCH_RESULT.batch_no = BATCH.batch_no "
    strSQL = strSQL & "WHERE search_result.search_no = " & mlngSearchNo & " "
    strSQL = strSQL & "GROUP BY BATCH.batch_no, BATCH.description "
    strSQL = strSQL & "HAVING Count(batch.batch_no) > 1 "
    strSQL = strSQL & "ORDER BY BATCH.batch_no"

'    Set recSearchResults = New ADODB.Recordset
'    recSearchResults.Open strSQL, gdbADO, adOpenStatic, , adCmdText
    Set recSearchResults = db.OpenRecordset(strSQL, dbOpenSnapshot)
    With recSearchResults
        While Not recSearchResults.EOF
            Set objSearchResult = Add(!batch_no)
            objSearchResult.BatchNo = NVL(!batch_no, NVLNumeric)
            objSearchResult.PhotoDesc = NVL(!Description, NVLString)
            .MoveNext
        Wend
    End With

    recSearchResults.Close
    Set recSearchResults = Nothing
'
'********* Code Ends Here **************
'
   Exit Sub
'
DuplicateEntries_Error:
    ErrorRaise "SearchResults.DuplicateEntries"
End Sub

Public Sub ConfirmAll()
'***************************************
' Module/Form Name   : SearchResults
'
' Procedure Name     : ConfirmAll
'
' Purpose            :
'
' Date Created       : 25/06/2004
'
' Author             : ADMINISTRATOR
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo ConfirmAll_Error
'
'******** Code Starts Here *************
'
    Dim strSQL          As String
    
    strSQL = "UPDATE SEARCH_RESULT "
    strSQL = strSQL & "SET CONFIRMED = True "
    strSQL = strSQL & "WHERE SEARCH_NO = " & CStr(mlngSearchNo)
    
''    gdbADO.Execute strSQL
    db.Execute strSQL
    Refresh
'
'********* Code Ends Here **************
'
   Exit Sub
'
ConfirmAll_Error:
    ErrorRaise "SearchResults.ConfirmAll"
End Sub

Public Sub UnConfirmAll()
'***************************************
' Module/Form Name   : SearchResults
'
' Procedure Name     : UnConfirmAll
'
' Purpose            :
'
' Date Created       : 25/06/2004
'
' Author             : ADMINISTRATOR
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo UnConfirmAll_Error
'
'******** Code Starts Here *************
'
    Dim strSQL          As String
    
    strSQL = "UPDATE SEARCH_RESULT "
    strSQL = strSQL & "SET CONFIRMED = False "
    strSQL = strSQL & "WHERE SEARCH_NO = " & CStr(mlngSearchNo)
    
''    gdbADO.Execute strSQL
    db.Execute strSQL
    Refresh
'
'********* Code Ends Here **************
'
   Exit Sub
'
UnConfirmAll_Error:
    ErrorRaise "SearchResults.UnConfirmAll"
End Sub

Public Sub Sort()
    Dim lngBatchArray()     As Long
    Dim intIdx              As Integer
    Dim oSearchResult       As SearchResult
    Dim oSearchResultsTemp  As SearchResults
    '
    '   Move the Batch numbers to the array to be sorted.
    '
    ReDim lngBatchArray(mCol.Count - 1, 2)
    Set oSearchResultsTemp = New SearchResults
    For intIdx = 0 To mCol.Count - 1
        Set oSearchResult = mCol.Item(intIdx + 1)
        lngBatchArray(intIdx, 0) = oSearchResult.BatchNo
        lngBatchArray(intIdx, 1) = oSearchResult.SearchResultKey
        '
        '   Move the Search Results to a temporary collection.
        '
        oSearchResultsTemp.AddObject oSearchResult
    Next intIdx
    '
    '   Sort the array.
    '
    QuickSortVariants lngBatchArray, LBound(lngBatchArray()), UBound(lngBatchArray())
    '
    '   Destroy the original collection.
    '
    While mCol.Count > 0
        mCol.Remove 1
    Wend
    Set mCol = Nothing
    Set mCol = New Collection
    '
    '   Use the array to control the copy back from the temporary collection
    '   to the new original collection.
    '
    For intIdx = 0 To UBound(lngBatchArray, 1)
        Set oSearchResult = oSearchResultsTemp.Item(CStr(lngBatchArray(intIdx, 1)))
        mCol.Add oSearchResult, CStr(oSearchResult.SearchResultKey)
    Next intIdx
    '
    '   Destroy the temporary collection.
    '
    While oSearchResultsTemp.Count > 0
        oSearchResultsTemp.Remove 1
    Wend
    Set oSearchResultsTemp = Nothing
    
End Sub

Private Sub QuickSortVariants(vArray As Variant, inLow As Long, inHi As Long)
      
   Dim pivot        As Variant
   Dim tmpSwapSort  As Variant
   Dim tmpSwapKey   As Variant
   Dim tmpLow       As Long
   Dim tmpHi        As Long
    
   tmpLow = inLow
   tmpHi = inHi
    
   pivot = vArray((inLow + inHi) \ 2, 0)
  
   While (tmpLow <= tmpHi)
  
      While (vArray(tmpLow, 0) < pivot And tmpLow < inHi)
         tmpLow = tmpLow + 1
      Wend
      
      While (pivot < vArray(tmpHi, 0) And tmpHi > inLow)
         tmpHi = tmpHi - 1
      Wend

      If (tmpLow <= tmpHi) Then
         tmpSwapSort = vArray(tmpLow, 0)
         tmpSwapKey = vArray(tmpLow, 1)
         vArray(tmpLow, 0) = vArray(tmpHi, 0)
         vArray(tmpLow, 1) = vArray(tmpHi, 1)
         vArray(tmpHi, 0) = tmpSwapSort
         vArray(tmpHi, 1) = tmpSwapKey
         tmpLow = tmpLow + 1
         tmpHi = tmpHi - 1
      End If
   
   Wend
  
   If (inLow < tmpHi) Then QuickSortVariants vArray, inLow, tmpHi
   If (tmpLow < inHi) Then QuickSortVariants vArray, tmpLow, inHi
  
End Sub

