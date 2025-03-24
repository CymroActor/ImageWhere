VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Jobs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private mCol                As Collection
Private mvarCurrentJobNo    As Long
Private mblnPendingRefresh  As Boolean

Public Function Add(JobNo As Long, Optional sKey As String) As Job2
'***************************************
' Module/Form Name   : Jobs
'
' Procedure Name     : Add
'
' Purpose            :
'
' Date Created       : 08/04/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : JobNo - Long
'                    : sKey - String
'
' Returns            : Job2
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 08/04/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo Add_Error
'
'******** Code Starts Here *************
'
    'create a new object
    Dim objNewMember As Job2
    Set objNewMember = New Job2

    objNewMember.JobNo = JobNo
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If
    
        'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing
'
'********* Code Ends Here **************
'
    Exit Function
    '
Add_Error:
    ErrorRaise "Jobs.Add"
End Function

Public Property Get Item(vntIndexKey As Variant) As Job2
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

Public Sub RefreshCourtCaseJobs(ByVal pblnShowOpenOnly As Boolean)
'***************************************
' Module/Form Name   : Jobs
'
' Procedure Name     : Refresh
'
' Purpose            :
'
' Date Created       : 08/04/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 08/04/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo Refresh_Error
'
'******** Code Starts Here *************
'
    Dim strSQL                  As String
''    Dim rstJobs                 As ADODB.Recordset
    Dim rstJobs                 As DAO.Recordset
    Dim oJob                    As Job2
    Dim blnCurrentJobNoFound    As Boolean
    Dim strPreviousJobNo        As Long
    Dim strWhereClause          As String
    '
    '   Clear all existing Jobs.
    '
    Set mCol = Nothing
    Set mCol = New Collection
    '
    If pblnShowOpenOnly Then
        strWhereClause = strWhereClause & IIf(strWhereClause = "", " WHERE", " AND") & " job.delivery_note_no < 1"
    End If
    
    strWhereClause = strWhereClause & IIf(strWhereClause = "", " WHERE", " AND") & " Delivery_note.date_created >= #3/1/2004# AND Delivery_note.date_created<= #2/28/2005#"
    '
    '   Refresh Jobs.
    '
    strSQL = "SELECT customer_name, Job.*, Delivery_note.DispatchTransparency, Delivery_note.DispatchDigital, " _
      & "(SELECT count(search_result.search_result_key) FROM search, search_result, photograph WHERE job.job_no = search.job_no " _
      & "and search.search_no = search_result.search_no and photograph.photograph_key = search_result.photograph_key and photograph.deleted = false) as no_of_photographs, " _
      & "(SELECT count(search_result.search_result_key) FROM search, search_result, photograph WHERE job.job_no = search.job_no " _
      & "and search.search_no = search_result.search_no and photograph.photograph_key = search_result.photograph_key and photograph.deleted = false " _
      & "and search_result.confirmed = true) as no_confirmed " _
      & "FROM (job INNER JOIN customer ON customer.customer_no = job.customer_no) " _
      & "INNER JOIN Delivery_note ON job.delivery_note_no = Delivery_note.delivery_note_no "
    strSQL = strSQL & strWhereClause

''    Set rstJobs = New ADODB.Recordset
''    rstJobs.Open strSQL, gdbADO, adOpenForwardOnly, , adCmdText
    Set rstJobs = db.OpenRecordset(strSQL, dbOpenSnapshot)
    With rstJobs
    strPreviousJobNo = 0
    blnCurrentJobNoFound = False
    While Not .EOF
        Set oJob = Add(!job_no, !job_no)
        With oJob
            .reference = NVL(rstJobs!job_reference)
            .DateTaken = NVL(rstJobs!date_taken, NVLNumeric)
            .TakenBy = NVL(rstJobs!taken_by, NVLString)
            .CustomerNo = NVL(rstJobs!customer_no, NVLNumeric)
            .RequestedBy = NVL(rstJobs!requested_by, NVLString)
            .RequestedByID = NVL(rstJobs!RequestedByID, NVLNumeric)
            .RequestorsTelNo = NVL(rstJobs!requestors_telno, NVLString)
            .DeliveryAddressLine1 = NVL(rstJobs!del_address_line_1, NVLString)
            .DeliveryAddressLine2 = NVL(rstJobs!del_address_line_2, NVLString)
            .DeliveryAddressLine3 = NVL(rstJobs!del_address_line_3, NVLString)
            .DeliveryCountyState = NVL(rstJobs!del_county_or_state, NVLString)
            .DeliveryCountry = NVL(rstJobs!del_country, NVLString)
            .DeliveryPostCode = NVL(rstJobs!del_post_code, NVLString)
            .DeliveryNoteNo = NVL(rstJobs!delivery_note_no, NVLNumeric)
            .DateRequired = NVL(rstJobs!date_required, NVLNumeric)
            .RequestDetails = NVL(rstJobs!request_details, NVLString)
            .DeliveryCustomerName = NVL(rstJobs!del_customer_name, NVLString)
            .RequestorsEmail = NVL(rstJobs!RequestorsEmail, NVLString)
            .WebStatus = NVL(rstJobs!WebStatus, NVLString)
            .WebDatePosted = NVL(rstJobs!WebDatePosted, NVLNumeric)
            .NoPhotos = NVL(rstJobs!no_of_photographs, NVLNumeric)
            .NoConfirmed = NVL(rstJobs!no_confirmed, NVLNumeric)
            .DispatchTransparency = NVL(rstJobs!DispatchTransparency, NVLString)
            .DispatchDigital = NVL(rstJobs!DispatchDigital, NVLString)
            .PreferredPhotoType = NVL(rstJobs!PreferredPhotoType, NVLString)
    ''35              oJob.CheckPoint
        End With
        '
        '   Determine the current Job No.
        '
        If Not blnCurrentJobNoFound Then
            If !job_no > mvarCurrentJobNo Then
                If strPreviousJobNo = 0 Then
                    mvarCurrentJobNo = !job_no
                    blnCurrentJobNoFound = True
                Else
                    mvarCurrentJobNo = strPreviousJobNo
                    blnCurrentJobNoFound = True
                End If
            ElseIf !job_no = mvarCurrentJobNo Then
                blnCurrentJobNoFound = True
            End If
        End If
        strPreviousJobNo = !job_no
        .MoveNext
    Wend
    End With
    If Not blnCurrentJobNoFound Then
    mvarCurrentJobNo = strPreviousJobNo
    End If
    rstJobs.Close
    Set rstJobs = Nothing
'
'********* Code Ends Here **************
'
Exit Sub
    '
Refresh_Error:
    ErrorRaise "Jobs.Refresh"
End Sub

Public Sub Refresh(ByVal pintCustomerNo As Integer, _
                   ByVal pblnShowOpenOnly As Boolean)
'***************************************
' Module/Form Name   : Jobs
'
' Procedure Name     : Refresh
'
' Purpose            :
'
' Date Created       : 08/04/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 08/04/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo Refresh_Error
'
'******** Code Starts Here *************
'
    Dim strSQL                  As String
''    Dim rstJobs                 As ADODB.Recordset
    Dim rstJobs                 As DAO.Recordset
    Dim oJob                    As Job2
    Dim blnCurrentJobNoFound    As Boolean
    Dim strPreviousJobNo        As Long
    Dim strWhereClause          As String
    '
    '   Clear all existing Jobs.
    '
    Set mCol = Nothing
    Set mCol = New Collection
    '
    '   Determine the filters if needed.
    strWhereClause = ""
    If pintCustomerNo <> 0 Then
        strWhereClause = " WHERE Job.Customer_No = " & CStr(pintCustomerNo)
    End If
    '
    If pblnShowOpenOnly Then
        strWhereClause = strWhereClause & IIf(strWhereClause = "", " WHERE", " AND") & " job.delivery_note_no < 1"
    End If
    '
    '   Refresh Jobs.
    '
    strSQL = "SELECT customer_name, Job.*, Delivery_note.DispatchTransparency, Delivery_note.DispatchDigital, " _
      & "(SELECT count(search_result.search_result_key) FROM search, search_result, photograph WHERE job.job_no = search.job_no " _
      & "and search.search_no = search_result.search_no and photograph.photograph_key = search_result.photograph_key and photograph.deleted = false) as no_of_photographs, " _
      & "(SELECT count(search_result.search_result_key) FROM search, search_result, photograph WHERE job.job_no = search.job_no " _
      & "and search.search_no = search_result.search_no and photograph.photograph_key = search_result.photograph_key and photograph.deleted = false " _
      & "and search_result.confirmed = true) as no_confirmed " _
      & "FROM (job INNER JOIN customer ON customer.customer_no = job.customer_no) " _
      & "LEFT JOIN Delivery_note ON job.delivery_note_no = Delivery_note.delivery_note_no "
    strSQL = strSQL & strWhereClause

''    Set rstJobs = New ADODB.Recordset
''    rstJobs.Open strSQL, gdbADO, adOpenForwardOnly, , adCmdText
    Set rstJobs = db.OpenRecordset(strSQL, dbOpenSnapshot)
    With rstJobs
    strPreviousJobNo = 0
    blnCurrentJobNoFound = False
    While Not .EOF
        Set oJob = Add(!job_no, !job_no)
        With oJob
            .reference = NVL(rstJobs!job_reference)
            .DateTaken = NVL(rstJobs!date_taken, NVLNumeric)
            .TakenBy = NVL(rstJobs!taken_by, NVLString)
            .CustomerNo = NVL(rstJobs!customer_no, NVLNumeric)
            .RequestedBy = NVL(rstJobs!requested_by, NVLString)
            .RequestedByID = NVL(rstJobs!RequestedByID, NVLNumeric)
            .RequestorsTelNo = NVL(rstJobs!requestors_telno, NVLString)
            .DeliveryAddressLine1 = NVL(rstJobs!del_address_line_1, NVLString)
            .DeliveryAddressLine2 = NVL(rstJobs!del_address_line_2, NVLString)
            .DeliveryAddressLine3 = NVL(rstJobs!del_address_line_3, NVLString)
            .DeliveryCountyState = NVL(rstJobs!del_county_or_state, NVLString)
            .DeliveryCountry = NVL(rstJobs!del_country, NVLString)
            .DeliveryPostCode = NVL(rstJobs!del_post_code, NVLString)
            .DeliveryNoteNo = NVL(rstJobs!delivery_note_no, NVLNumeric)
            .DateRequired = NVL(rstJobs!date_required, NVLNumeric)
            .RequestDetails = NVL(rstJobs!request_details, NVLString)
            .DeliveryCustomerName = NVL(rstJobs!del_customer_name, NVLString)
            .RequestorsEmail = NVL(rstJobs!RequestorsEmail, NVLString)
            .WebStatus = NVL(rstJobs!WebStatus, NVLString)
            .WebDatePosted = NVL(rstJobs!WebDatePosted, NVLNumeric)
            .NoPhotos = NVL(rstJobs!no_of_photographs, NVLNumeric)
            .NoConfirmed = NVL(rstJobs!no_confirmed, NVLNumeric)
            .DispatchTransparency = NVL(rstJobs!DispatchTransparency, NVLString)
            .DispatchDigital = NVL(rstJobs!DispatchDigital, NVLString)
            .PreferredPhotoType = NVL(rstJobs!PreferredPhotoType, NVLString)
    ''35              oJob.CheckPoint
        End With
        '
        '   Determine the current Job No.
        '
        If Not blnCurrentJobNoFound Then
            If !job_no > mvarCurrentJobNo Then
                If strPreviousJobNo = 0 Then
                    mvarCurrentJobNo = !job_no
                    blnCurrentJobNoFound = True
                Else
                    mvarCurrentJobNo = strPreviousJobNo
                    blnCurrentJobNoFound = True
                End If
            ElseIf !job_no = mvarCurrentJobNo Then
                blnCurrentJobNoFound = True
            End If
        End If
        strPreviousJobNo = !job_no
        .MoveNext
    Wend
    End With
    If Not blnCurrentJobNoFound Then
    mvarCurrentJobNo = strPreviousJobNo
    End If
    rstJobs.Close
    Set rstJobs = Nothing
'
'********* Code Ends Here **************
'
Exit Sub
    '
Refresh_Error:
    ErrorRaise "Jobs.Refresh"
End Sub

Public Property Get CurrentJobNo() As Long
    CurrentJobNo = mvarCurrentJobNo
End Property

Public Property Let CurrentJobNo(vData As Long)
    mvarCurrentJobNo = vData
End Property

Public Sub UpdateCountsWarm()
'***************************************
' Module/Form Name   : Jobs
'
' Procedure Name     : UpdateCountsWarm
'
' Purpose            :
'
' Date Created       : 07/07/2004
'
' Author             : ADMINISTRATOR
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo UpdateCountsWarm_Error
'
'******** Code Starts Here *************
'
    Dim oJob                As Job2
    Dim oSearch             As Search2
    Dim oSearchResult       As SearchResult
    Dim intConfirmed        As Integer
    Dim intNoPhotographs    As Integer
    
    For Each oJob In mCol
        intConfirmed = 0
        intNoPhotographs = 0
        For Each oSearch In oJob.Searches
            For Each oSearchResult In oSearch.SearchResults
                If oSearchResult.Confirmed Then
                    intConfirmed = intConfirmed + 1
                End If
            Next oSearchResult
            intNoPhotographs = intNoPhotographs + oSearch.SearchResults.Count
        Next oSearch
        oJob.NoConfirmed = intConfirmed
        oJob.NoPhotos = intNoPhotographs
    Next oJob
    
    Set oJob = Nothing
    Set oSearch = Nothing
    Set oSearchResult = Nothing
'
'********* Code Ends Here **************
'
   Exit Sub
'
UpdateCountsWarm_Error:
    ErrorRaise "Jobs.UpdateCountsWarm"
End Sub

Public Property Get PendingRefresh() As Boolean
    PendingRefresh = mblnPendingRefresh
End Property

Public Property Let PendingRefresh(vData As Boolean)
    mblnPendingRefresh = vData
End Property


