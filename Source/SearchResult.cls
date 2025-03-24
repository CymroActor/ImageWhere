VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SearchResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mvarSearchResultKey As Long
Private mvarSearchNo        As Long
Private mvarCustomerNo      As Integer
Private mvarBatchNo         As Long
Private mvarPhotographNo    As Integer
Private mvarPhotoDesc       As String
Private mvarDateReturned    As Date
Private mvarBookedInBy      As String
Private mvarConfirmed       As Boolean
Private mvarCredit          As String
Private mvarPhotoType       As String
Private mvarPhotographKey   As Long
Private mvarJobDescription  As String
Private mvarExclusive       As Boolean
Private mvarDigitalStatus   As String

''Private mrecSearchResult As ADODB.Recordset
Private Type udtCheckPoint
    SearchResultKey     As Long
    SearchNo            As Long
    CustomerNo          As Integer
    BatchNo             As Long
    PhotographNo        As Integer
    PhotoDesc           As String
    DateReturned        As Date
    BookedInBy          As String
    Confirmed           As Boolean
    PhotoType           As String
    PhotographKey       As Long
    JobDescription      As String
    Exclusive           As Boolean
    DigitalStatus       As String
End Type

Private SearchResultCheckPoint As udtCheckPoint

Public Property Get SearchResultKey() As Long
    SearchResultKey = mvarSearchResultKey
End Property

Friend Property Let SearchResultKey(vData As Long)
    mvarSearchResultKey = vData
End Property

Public Property Get SearchNo() As Long
    SearchNo = mvarSearchNo
End Property

Friend Property Let SearchNo(vData As Long)
    mvarSearchNo = vData
End Property

Public Property Get CustomerNo() As Integer
    CustomerNo = mvarCustomerNo
End Property

Friend Property Let CustomerNo(vData As Integer)
    mvarCustomerNo = vData
End Property

Public Property Get BatchNo() As Long
    BatchNo = mvarBatchNo
End Property

Friend Property Let BatchNo(vData As Long)
    mvarBatchNo = vData
End Property

Public Property Get PhotographNo() As Long
    PhotographNo = mvarPhotographNo
End Property

Friend Property Let PhotographNo(vData As Long)
    mvarPhotographNo = vData
End Property

Public Property Let PhotoDesc(ByVal vData As String)
    mvarPhotoDesc = vData
End Property

Public Property Get PhotoDesc() As String
    PhotoDesc = mvarPhotoDesc
End Property

Public Property Get DateReturned() As Date
    DateReturned = mvarDateReturned
End Property

Public Property Let DateReturned(ByVal vData As Date)
    mvarDateReturned = vData
End Property

Public Property Let BookedInBy(ByVal vData As String)
    mvarBookedInBy = vData
End Property

Public Property Get BookedInBy() As String
    BookedInBy = mvarBookedInBy
End Property

Public Property Get Confirmed() As Boolean
    Confirmed = mvarConfirmed
End Property

Public Property Let Confirmed(ByVal vData As Boolean)
    mvarConfirmed = vData
End Property

Public Property Let Credit(ByVal vData As String)
    mvarCredit = vData
End Property

Public Property Get Credit() As String
    Credit = mvarCredit
End Property

Public Property Let PhotoType(ByVal vData As String)
    mvarPhotoType = vData
End Property

Public Property Get PhotoType() As String
    PhotoType = mvarPhotoType
End Property

Public Property Get PhotographKey() As Long
    PhotographKey = mvarPhotographKey
End Property

Friend Property Let PhotographKey(vData As Long)
    mvarPhotographKey = vData
End Property

Public Property Let JobDescription(ByVal vData As String)
    mvarJobDescription = vData
End Property

Public Property Get JobDescription() As String
    JobDescription = mvarJobDescription
End Property

Public Property Get Exclusive() As Boolean
    Exclusive = mvarExclusive
End Property

Public Property Let Exclusive(ByVal vData As Boolean)
    mvarExclusive = vData
End Property

Public Property Get DigitalStatus() As String
    DigitalStatus = mvarDigitalStatus
End Property

Public Property Let DigitalStatus(ByVal vData As String)
    If vData <> "" And vData <> "A" And vData <> "R" Then
        Err.Raise vbObjectError + 4, , "DigitalStatus can only be '','A' or 'R'"
    End If
    
    mvarDigitalStatus = vData
End Property

Public Sub CheckPoint()
    With SearchResultCheckPoint
        .SearchResultKey = mvarSearchResultKey
        .SearchNo = mvarSearchNo
        .BatchNo = mvarBatchNo
        .PhotographNo = mvarPhotographNo
        .PhotoDesc = mvarPhotoDesc
        .DateReturned = mvarDateReturned
        .BookedInBy = mvarBookedInBy
        .Confirmed = mvarConfirmed
        .PhotoType = mvarPhotoType
        .PhotographKey = mvarPhotographKey
        .JobDescription = mvarJobDescription
        .Exclusive = mvarExclusive
        .DigitalStatus = mvarDigitalStatus
    End With
    
    Exit Sub
End Sub

Private Sub Class_Initialize()
    mvarCredit = ""
    mvarPhotoType = "T"
    mvarExclusive = False
    mvarDigitalStatus = ""
End Sub

Public Sub Delete()
'***************************************
' Module/Form Name   : SearchResult
'
' Procedure Name     : Delete
'
' Purpose            :
'
' Date Created       : 24/06/2004
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo Delete_Error
'
'******** Code Starts Here *************
'
    Dim strSQL              As String
   
    strSQL = "DELETE * FROM SEARCH_RESULT "
    strSQL = strSQL & "WHERE SEARCH_RESULT_KEY = " & CStr(mvarSearchResultKey)
    
''    gdbADO.Execute strSQL
    db.Execute strSQL
    goLog.WriteDetailedLog CStr(mvarSearchNo) & "/" & CStr(mvarBatchNo) & "/" & CStr(mvarPhotographNo), "Search Result Deleted"
'
'********* Code Ends Here **************
'
   Exit Sub
'
Delete_Error:
    ErrorRaise "SearchResult.Delete"
End Sub
''
''Public Sub Create()
'''***************************************
''' Module/Form Name   : SearchResult
'''
''' Procedure Name     : Create
'''
''' Purpose            :
'''
''' Date Created       : 30/06/2004
'''
''' Author             : GARETH SAUNDERS
'''
''' Amendment History  : Date       Author    Description
'''                    : --------------------------------
'''
'''***************************************
'''
''On Error GoTo Create_Error
'''
'''******** Code Starts Here *************
'''
''    Dim strSQL              As String
''    Dim rstSearchResult     As ADODB.Recordset
''
''    Set rstSearchResult = New ADODB.Recordset
''    rstSearchResult.Open "Search_Result", gdbADO, adOpenDynamic, adLockOptimistic
''
''    With rstSearchResult
''        .AddNew
''        !search_no = mvarSearchNo
''        !job_description = mvarJobDescription
''        !customer_no = mvarCustomerNo
''        !batch_no = mvarBatchNo
''        !photograph_no = mvarPhotographNo
''        !date_returned = mvarDateReturned
''        !booked_in_by = IIf(mvarBookedInBy = "", Null, mvarBookedInBy)
''        !Confirmed = mvarConfirmed
''        !photograph_key = mvarPhotographKey
''        .Update
''        mvarSearchResultKey = !search_result_key
''    End With
''    goLog.WriteDetailedLog CStr(mvarSearchNo) & "/" & CStr(mvarBatchNo) & "/" & CStr(mvarPhotographNo), "Search Result Created"
'''
'''********* Code Ends Here **************
'''
''   Exit Sub
'''
''Create_Error:
''    ErrorRaise "SearchResult.Create"
''End Sub

Public Sub CreateDAO()
'***************************************
' Module/Form Name   : SearchResult
'
' Procedure Name     : CreateDAO
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
On Error GoTo CreateDAO_Error
'
'******** Code Starts Here *************
'
    Dim rstSearchResult     As DAO.Recordset
    
    Set rstSearchResult = db.OpenRecordset("Search_result", dbOpenTable)
    With rstSearchResult
        .AddNew
        !search_no = mvarSearchNo
        !job_description = mvarJobDescription
        !customer_no = mvarCustomerNo
        !batch_no = mvarBatchNo
        !photograph_no = mvarPhotographNo
        !date_returned = mvarDateReturned
        !booked_in_by = IIf(mvarBookedInBy = "", Null, mvarBookedInBy)
        !Confirmed = mvarConfirmed
        !photograph_key = mvarPhotographKey
        On Error Resume Next
        .update
        If Err.Number = 3201 Then
            Err.Raise vbObjectError + 2, , "Search has been removed."
        ElseIf Err.Number = 3022 Then
            Err.Raise vbObjectError + 3, , "Photograph already exists on Search"
        ElseIf Err.Number <> 0 Then
            ErrorSave
            On Error GoTo CreateDAO_Error
            ErrorRestore
        Else
            On Error GoTo CreateDAO_Error
        End If
        .Bookmark = .LastModified
        mvarSearchResultKey = !search_result_key
    End With
    goLog.WriteDetailedLog CStr(mvarSearchNo) & "/" & CStr(mvarBatchNo) & "/" & CStr(mvarPhotographNo), "Search Result Created (DAO)"
'
'********* Code Ends Here **************
'
   Exit Sub
'
CreateDAO_Error:
    ErrorRaise "SearchResult.CreateDAO"
End Sub
''
''Public Sub Update()
'''***************************************
''' Module/Form Name   : SearchResult
'''
''' Procedure Name     : Update
'''
''' Purpose            :
'''
''' Date Created       : 30/06/2004
'''
''' Author             : GARETH SAUNDERS
'''
''' Amendment History  : Date       Author    Description
'''                    : --------------------------------
'''
'''***************************************
'''
''On Error GoTo Update_Error
'''
'''******** Code Starts Here *************
'''
''    Dim rstSearchResult     As ADODB.Recordset
''    Dim strSQL              As String
''
''    strSQL = "SELECT * FROM SEARCH_RESULT " _
''         & "WHERE SEARCH_RESULT_KEY = " & CStr(mvarSearchResultKey)
''
''    Set rstSearchResult = New ADODB.Recordset
''    rstSearchResult.Open strSQL, gdbADO, adOpenDynamic, adLockPessimistic, adCmdText
''
''    If rstSearchResult.EOF Then
''        Err.Raise vbObjectError + 1, , "Search Result does not Exist"
''    End If
''    With rstSearchResult
''          ''        CheckRecordSame
''        !date_returned = mvarDateReturned
''        !booked_in_by = IIf(mvarBookedInBy = "", Null, mvarBookedInBy)
''        !Confirmed = mvarConfirmed
''        .Update
''          ''        CheckPoint
''    End With
''
''    rstSearchResult.Close
''    Set rstSearchResult = Nothing
''    goLog.WriteDetailedLog CStr(mvarSearchNo) & "/" & CStr(mvarBatchNo) & "/" & CStr(mvarPhotographNo), "Search Result Updated"
'''
'''********* Code Ends Here **************
'''
''   Exit Sub
'''
''Update_Error:
''    ErrorRaise "SearchResult.Update"
''End Sub

Public Sub UpdateDAO()
'***************************************
' Module/Form Name   : SearchResult
'
' Procedure Name     : UpdateDAO
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
On Error GoTo UpdateDAO_Error
'
'******** Code Starts Here *************
'
    Dim rstSearchResult     As DAO.Recordset
    Dim strSQL              As String

    strSQL = "SELECT * FROM SEARCH_RESULT " _
         & "WHERE SEARCH_RESULT_KEY = " & CStr(mvarSearchResultKey)
    
    Set rstSearchResult = db.OpenRecordset(strSQL, dbOpenDynaset)
    
    If rstSearchResult.EOF Then
        Err.Raise vbObjectError + 1, , "Search Result does not Exist"
    End If
    With rstSearchResult
          ''        CheckRecordSame
        .Edit
        !date_returned = mvarDateReturned
        !booked_in_by = IIf(mvarBookedInBy = "", Null, mvarBookedInBy)
        !Confirmed = mvarConfirmed
        !DigitalStatus = IIf(mvarDigitalStatus = "", Null, mvarDigitalStatus)
        .update
        .Bookmark = .LastModified
          ''        CheckPoint
    End With
    
    rstSearchResult.Close
    Set rstSearchResult = Nothing
'
'********* Code Ends Here **************
'
   Exit Sub
'
UpdateDAO_Error:
    ErrorRaise "SearchResult.UpdateDAO"
End Sub

Public Sub Read(ByVal plngSearchResultKey As Long)

    Dim strSQL              As String
    Dim strWhereClause      As String
    Dim rstSearchResult           As DAO.Recordset
    '
'***************************************
' Module/Form Name   : SearchResult
'
' Procedure Name     : Read
'
' Purpose            :
'
' Date Created       : 13/06/2006 23:33
'
' Author             : Gareth
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
    On Error GoTo ErrorProc

    strSQL = "SELECT SEARCH_RESULT.* "
    strSQL = strSQL & "FROM search_result "
    strSQL = strSQL & "WHERE search_result.search_result_key = " & CStr(plngSearchResultKey)
    '
    Set rstSearchResult = db.OpenRecordset(strSQL, , dbOpenSnapshot)
    With rstSearchResult
        mvarSearchResultKey = !search_result_key
        mvarSearchNo = !search_no
        mvarCustomerNo = !customer_no
        mvarBatchNo = !batch_no
        mvarPhotographNo = !photograph_no
        mvarDateReturned = NVL(!date_returned, NVLNumeric)
        mvarBookedInBy = !booked_in_by & ""
        mvarConfirmed = !Confirmed
        mvarDigitalStatus = !DigitalStatus & ""
    End With
    rstSearchResult.Close
    Set rstSearchResult = Nothing

    On Error GoTo 0
    Exit Sub
'
'********* Code Ends Here **************
'
ErrorProc:
    ErrorRaise "SearchResult.Read"
End Sub


