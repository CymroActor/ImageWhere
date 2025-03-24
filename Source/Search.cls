VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Search2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mvarSearchNo                As Long
Private mvarJobNo                   As Long
Private mvarDescription             As String
Private mvarSearchDate              As Date
Private mvarNoPhotographs           As Integer
Private mvarDigitalSearchNo         As Long

''Private mrecSearch                  As ADODB.Recordset
Private mvarSearchResults           As SearchResults
Private mvarDuplicateSearchResults  As SearchResults
Private mvarActivities              As Activities
Private mvarPhotoType               As String
Private mvarPhotoTypeDescription    As String

Public Sub Delete()
    '***************************************
    ' Module/Form Name   : Search2
    '
    ' Procedure Name     : Delete
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
    On Error GoTo Delete_Error
    '
    '******** Code Starts Here *************
    '
    Dim SQL As String
    '
    '   Can't Delete if there are photographs attached.
    '
    If Not mvarSearchResults Is Nothing Then
        mvarSearchResults.Refresh
    End If
    '
    If SearchResults.Count <> 0 Then
        Err.Raise vbObjectError + 4, , "Search has photographs." & vbCrLf & "These must be removed first before deleting the Search"
    End If
    '
''    Set mrecSearch.ActiveConnection = gdbADO
    
    SQL = "DELETE * FROM SEARCH "
    SQL = SQL & "WHERE SEARCH_NO = " & mvarSearchNo
    
''    gdbADO.Execute SQL
    db.Execute SQL
    goLog.WriteDetailedLog CStr(mvarJobNo) & "/" & CStr(mvarSearchNo), "Search Deleted"
    '
    '********* Code Ends Here **************
    '
    Exit Sub
    '
Delete_Error:
    ErrorRaise "Search2.Delete"
End Sub

Public Property Get SearchNo() As Long
    SearchNo = mvarSearchNo
End Property

Friend Property Let SearchNo(vData As Long)
    mvarSearchNo = vData
End Property

Public Property Get JobNo() As Long
    JobNo = mvarJobNo
End Property

Friend Property Let JobNo(vData As Long)
    mvarJobNo = vData
End Property

Public Property Let Description(ByVal vData As String)
    mvarDescription = vData
End Property

Public Property Get Description() As String
    Description = mvarDescription
End Property

Public Property Get DisplayDescription() As String
    If mvarPhotoType = "D" Then
        DisplayDescription = mvarDescription & " (DIGITAL)"
    Else
        DisplayDescription = mvarDescription
    End If
End Property

Public Property Get SearchDate() As Date
    SearchDate = mvarSearchDate
End Property

Public Property Let SearchDate(ByVal vData As Date)
    mvarSearchDate = vData
End Property

Public Property Get NoPhotographs() As Integer
    NoPhotographs = mvarNoPhotographs
End Property

Public Property Let NoPhotographs(ByVal vData As Integer)
    mvarNoPhotographs = vData
End Property

Public Property Let PhotoType(ByVal vData As String)
    mvarPhotoType = vData
End Property

Public Property Get PhotoType() As String
    PhotoType = mvarPhotoType
End Property

Public Property Get PhotoTypeDescription() As String
    PhotoTypeDescription = goSystemConfig.PhotoTypeDescription(mvarPhotoType)
End Property

Public Property Let DigitalSearchNo(ByVal vData As Long)
    mvarDigitalSearchNo = vData
End Property

Public Property Get DigitalSearchNo() As Long
    DigitalSearchNo = mvarDigitalSearchNo
End Property

Public Sub Add(JobNo As Long, _
               Description As String, _
               SearchDate As Date)
    '***************************************
    ' Module/Form Name   : Search2
    '
    ' Procedure Name     : Add
    '
    ' Purpose            :
    '
    ' Date Created       : 28/04/2002
    '
    ' Author             : GARETH SAUNDERS
    '
    ' Parameters         : SearchNo - Long
    '                    : JobNo - Long
    '                    : Description - String
    '                    : SearchDate - Date
    '
    ' Amendment History  : Date       Author    Description
    '                    : --------------------------------
    '                    : 28/04/2002 GARETH SAUNDERS
    '
    '***************************************
    '
On Error GoTo Add_Error
    '
    '******** Code Starts Here *************
    '

    Dim strSQL      As String
''    Dim rstSearch   As ADODB.Recordset
    Dim rstSearch   As DAO.Recordset
    
    mvarJobNo = JobNo
    mvarDescription = Description
    mvarSearchDate = SearchDate

''    Set rstSearch = Nothing
''    Set rstSearch = New ADODB.Recordset
    Set rstSearch = db.OpenRecordset("Search", dbOpenTable)
    With rstSearch
''        .Open "Search", gdbADO, adOpenDynamic, adLockOptimistic
        .AddNew
        !job_no = mvarJobNo
        !Search_Description = mvarDescription
        !date_of_search = mvarSearchDate
        !PhotoType = mvarPhotoType
        !DigitalSearchNo = Iif(mvarDigitalSearchNo = 0, Null, mvarDigitalSearchNo)
        .update
        .Bookmark = .LastModified
        mvarSearchNo = !search_no
        .Close
        Set rstSearch = Nothing
    End With
    goLog.WriteDetailedLog CStr(mvarJobNo) & "/" & CStr(mvarSearchNo), "Search Created"
    '
    '********* Code Ends Here **************
    '
Exit Sub
    '
Add_Error:
    If Err.Number = 3022 Then
        Err.Raise 1 + vbObjectError, , "Search already exists"
    Else
        ErrorRaise "Search2.Add"
    End If
End Sub

Public Sub AddDAO(JobNo As Long, _
                  Description As String, _
                  SearchDate As Date)
'***************************************
' Module/Form Name   : Search2
'
' Procedure Name     : AddDAO
'
' Purpose            :
'
' Date Created       : 08/07/2004
'
' Author             : GARETH SAUNDERS
'
' Parameters         : JobNo - Long
'                    : Description - String
'                    : SearchDate - Date
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo AddDAO_Error
'
'******** Code Starts Here *************
'
    Dim strSQL As String
    Dim rstSearch As DAO.Recordset
    
    mvarJobNo = JobNo
    mvarDescription = Description
    mvarSearchDate = SearchDate

    Set rstSearch = Nothing
    Set rstSearch = db.OpenRecordset("Search", dbOpenDynaset)
    With rstSearch
        .AddNew
        !job_no = mvarJobNo
        !Search_Description = mvarDescription
        !date_of_search = mvarSearchDate
        !PhotoType = mvarPhotoType
        !DigitalSearchNo = Iif(mvarDigitalSearchNo = 0, Null, mvarDigitalSearchNo)
        On Error Resume Next
        .update
        If Err.Number = 3022 Then
            On Error GoTo AddDAO_Error
            Err.Raise 1 + vbObjectError, , "Search already exists"
        ElseIf Err.Number = 0 Then
            On Error GoTo AddDAO_Error
        Else
            ErrorSave
            On Error GoTo AddDAO_Error
            ErrorRestore
        End If
        .Bookmark = .LastModified
        mvarSearchNo = !search_no
        .Close
        Set rstSearch = Nothing
    End With
    goLog.WriteDetailedLog CStr(mvarJobNo) & "/" & CStr(mvarSearchNo), "Search Created (DAO)"
'
'********* Code Ends Here **************
'
   Exit Sub
'
AddDAO_Error:
    ErrorRaise "Search2.AddDAO"
End Sub

Public Sub update()
'***************************************
' Module/Form Name   : Search2
'
' Procedure Name     : update
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
On Error GoTo update_Error
'
'******** Code Starts Here *************
'
    Dim strSQL      As String
    Dim rstSearch   As DAO.Recordset
    
    strSQL = "SELECT * FROM SEARCH " _
         & "WHERE SEARCH_NO = " & CStr(mvarSearchNo)
    
''    Set mrecSearch = New ADODB.Recordset
''    mrecSearch.Open strSQL, gdbADO, adOpenDynamic, adLockPessimistic, adCmdText
    Set rstSearch = db.OpenRecordset(strSQL, dbOpenDynaset)
    
''    If mrecSearch.EOF Then
    If rstSearch.EOF Then
        Err.Raise vbObjectError + 2, , "Search does not Exist"
    End If
''    With mrecSearch
    With rstSearch
          ''        CheckRecordSame
        .Edit
        !Search_Description = mvarDescription
        !date_of_search = mvarSearchDate
        !DigitalSearchNo = mvarDigitalSearchNo
        !PhotoType = mvarPhotoType
        .update
          ''        CheckPoint
    End With
    
''    mrecSearch.Close
    rstSearch.Close
    Set rstSearch = Nothing
    goLog.WriteDetailedLog CStr(mvarJobNo) & "/" & CStr(mvarSearchNo), "Search Updated"
'
'********* Code Ends Here **************
'
Exit Sub
    '
update_Error:
    ErrorRaise "Search2.update"
End Sub

Public Sub Read(Optional ByVal plngSearchNo As Long = 0, _
                Optional ByVal plngJobNo As Long = 0, _
                Optional ByVal pstrDescription As String = "", _
                Optional ByVal pstrPhotoType As String = "")
'***************************************
' Module/Form Name   : Search2
'
' Procedure Name     : Read
'
' Purpose            :
'
' Date Created       : 28/04/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : SearchNo - Long
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 28/04/2002 GARETH SAUNDERS
'                    : 01/12/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo Read_Error
'
'******** Code Starts Here *************
'

    Dim strSQL              As String
    Dim strWhereClause      As String
    Dim rstSearch           As DAO.Recordset
    '
    '   Validate the input parameters.
    '
    If (plngSearchNo = 0 And _
        (plngJobNo = 0 Or pstrDescription = "" Or pstrPhotoType = "")) Or _
       (plngSearchNo <> 0 And _
        (plngJobNo <> 0 Or pstrDescription <> "" Or pstrPhotoType <> "")) Then
            Err.Raise vbObjectError + 5, , "Incorrect parameters supplied to Read function."
    End If
    '
    '   Construct the Where Clause.
    '
    If plngSearchNo <> 0 Then
        strWhereClause = "WHERE SEARCH_NO = " & CStr(plngSearchNo)
    Else
        strWhereClause = "WHERE JOB_NO = " & CStr(plngJobNo) & " "
        strWhereClause = strWhereClause & "AND SEARCH_DESCRIPTION = '" & Replace(pstrDescription, "'", "''") & "' "
        strWhereClause = strWhereClause & "AND PHOTOTYPE = '" & pstrPhotoType & "' "
    End If
    '
    strSQL = "SELECT SEARCH.*, "
    strSQL = strSQL & "(SELECT count(*) "
    strSQL = strSQL & "FROM search_result, photograph "
    strSQL = strSQL & "WHERE search.search_no = search_result.search_no "
    strSQL = strSQL & "and photograph.photograph_key = search_result.photograph_key "
    strSQL = strSQL & "and photograph.deleted = false) as NoPhotographs "
    strSQL = strSQL & "FROM SEARCH "
    strSQL = strSQL & strWhereClause
    '
''    Set mrecSearch = Nothing
''    Set mrecSearch = New ADODB.Recordset
''    mrecSearch.CursorLocation = adUseClient
''    mrecSearch.Open strSQL, gdbADO, adOpenDynamic, adLockOptimistic, adCmdText
''    Set mrecSearch.ActiveConnection = Nothing
    Set rstSearch = db.OpenRecordset(strSQL, , dbOpenSnapshot)
    Refresh rstSearch
''    CheckPoint
'
'********* Code Ends Here **************
'
    Exit Sub
        '
Read_Error:
    ErrorRaise "Search2.Read"
End Sub

''Private Function CheckRecordSame() As Boolean
''    '
''    '   Check if record has changed by comparing it with the data
''    '   at last refresh.
''    '
''    With SearchCheckPoint
''        If mrecSearch!search_no <> .SearchNo Or _
''           mrecSearch!job_no <> .JobNo Or _
''           mrecSearch!Search_description <> .Description Or _
''           mrecSearch!Date_of_search <> .SearchDate Then
''            Refresh
''            Err.Raise vbObjectError + 3, , "Search has been updated since first obtaining it."
''        End If
''    End With
''End Function
''
Private Function Refresh(ByRef prstSearch As DAO.Recordset)
'***************************************
' Module/Form Name   : Search2
'
' Procedure Name     : Refresh
'
' Purpose            :
'
' Date Created       : 01/12/2002
'
' Author             : GARETH SAUNDERS
'
' Returns            :
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 01/12/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo Refresh_Error
'
'******** Code Starts Here *************
'
    With prstSearch
        mvarSearchNo = NVL(!search_no)
        mvarJobNo = NVL(!job_no)
        mvarDescription = NVL(!Search_Description)
        mvarSearchDate = NVL(!date_of_search)
        mvarNoPhotographs = NVL(!NoPhotographs, NVLNumeric)
        mvarDigitalSearchNo = NVL(!DigitalSearchNo, NVLNumeric)
        mvarPhotoType = NVL(!PhotoType, NVLString)
    End With
'
'********* Code Ends Here **************
'
Exit Function
    '
Refresh_Error:
    ErrorRaise "Search2.Refresh"
End Function

''Public Sub CheckPoint()
''    With SearchCheckPoint
''        .SearchNo = mvarSearchNo
''        .JobNo = mvarJobNo
''        .Description = mvarDescription
''        .SearchDate = mvarSearchDate
''    End With
''
''    Exit Sub
''End Sub
''
Public Property Get SearchResults() As SearchResults
    If mvarSearchResults Is Nothing Then
        Set mvarSearchResults = New SearchResults
        mvarSearchResults.SearchNo(srAll) = mvarSearchNo
    End If

    Set SearchResults = mvarSearchResults
End Property

Public Property Set SearchResults(vData As SearchResults)
    Set mvarSearchResults = vData
End Property

Public Property Get Activities() As Activities
    If mvarActivities Is Nothing Then
        Set mvarActivities = New Activities
        mvarActivities.SearchNo = mvarSearchNo
    End If

    Set Activities = mvarActivities
End Property

Public Property Set Activities(vData As Activities)
    Set mvarActivities = vData
End Property

Public Property Get DuplicateSearchResults() As SearchResults
    If mvarDuplicateSearchResults Is Nothing Then
        Set mvarDuplicateSearchResults = New SearchResults
        mvarDuplicateSearchResults.SearchNo(srDuplicates) = mvarSearchNo
    End If

    Set DuplicateSearchResults = mvarDuplicateSearchResults
End Property

Private Sub Class_Initialize()
    mvarDigitalSearchNo = 0
End Sub

Public Sub AddImage(ByRef poDigitalImage As DigitalImage)
'***************************************
' Module/Form Name   : Search2
'
' Procedure Name     : AddImage
'
' Purpose            :
'
' Date Created       : 29/06/2004
'
' Author             : GARETH SAUNDERS
'
' Parameters         : poDigitalImage - DigitalImage
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo AddImage_Error
'
'******** Code Starts Here *************
'
'
'********* Code Ends Here **************
'
   Exit Sub
'
AddImage_Error:
    ErrorRaise "Search2.AddImage"
End Sub

Public Property Get WebSearchUseByDate(ByVal dteWebDatePosted As Date) As Date
    Dim intUseByPeriod  As Integer
    
    If mvarPhotoType = "D" Then
        intUseByPeriod = goSystemConfig.WebUseByPeriodDigital
    Else
        intUseByPeriod = goSystemConfig.WebUseByPeriod
    End If
    '
    WebSearchUseByDate = DateAdd("d", intUseByPeriod, dteWebDatePosted)
End Property

