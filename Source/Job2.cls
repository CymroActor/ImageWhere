VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Job2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mvarJobNo                       As Long
Private mvarDeliveryNoteNo              As Long
Private mvarReference                   As String
Private mvarDateTaken                   As Date
Private mvarTakenBy                     As String
Private mvarCustomerNo                  As Integer
Private mvarRequestedBy                 As String
Private mvarRequestedByID               As Long
Private mvarRequestorsTelNo             As String
Private mvarDeliveryAddressLine1        As String
Private mvarDeliveryAddressLine2        As String
Private mvarDeliveryAddressLine3        As String
Private mvarDeliveryCountyState         As String
Private mvarDeliveryCountry             As String
Private mvarDeliveryPostCode            As String
Private mvarDateRequired                As Date
Private mvarRequestDetails              As String
Private mvarDeliveryCustomerName        As String
Private mvarRequestorsEmail             As String
Private mvarWebStatus                   As String
Private mvarWebDatePosted               As Date
Private mvarNoPhotos                    As Integer
Private mvarNoConfirmed                 As Integer
Private mvarNoConfirmedTrans            As Integer
Private mvarNoConfirmedDigital          As Integer
Private mvarDispatchTransparency        As String
Private mvarDispatchDigital             As String
Private mvarServiceFee                  As Currency
Private mvarNoPhotosRequested           As Integer
Private mvarPreferredPhotoType          As String
Private mvarDigitalSearchesFolder       As String
'
Private Const DIGITAL_SEARCHES_FOLDER   As String = "DigitalSearches"
'
Private mvarSearches As Searches
'
'   FTP Variables and Dlls
'
''Private Const CONST_USER = "clivenichols"
''Private Const CONST_PASSWORD = "clin88"
''Private Const CONST_SERVER = "ftp.transcom.com"
'
'   Constants
'
Private Const CONST_FILENAME_WEBJOBS = "webjobs.txt"
Private Const CONST_FILENAME_WEBSEARCHES = "websearches.txt"
Private Const CONST_FILENAME_WEBPHOTOGRAPHS = "webphotographs.txt"
Private mvarActivities          As Activities
Private mvarJobNotes            As Activities
''
''Public Property Get NoOutstandingPictures() As Long
''    Dim strSQL As String
''    Dim rstJob As DAO.Recordset
''
''    strSQL = "SELECT * FROM JOB WHERE JOB_NO = " & mvarJobNo
''    Set rstJob = db.OpenRecordset(strSQL, dbOpenSnapshot)
''
''    NoOutstandingPictures = Val(rstJob!NoOutstandingPictures & "")
''
''    Set rstJob = Nothing
''End Property

Public Property Let JobNo(vData As Long)
    mvarJobNo = vData
End Property

Public Property Get JobNo() As Long
    JobNo = mvarJobNo
End Property

Public Property Let DeliveryNoteNo(vData As Long)
    mvarDeliveryNoteNo = vData
End Property

Public Property Get DeliveryNoteNo() As Long
    DeliveryNoteNo = mvarDeliveryNoteNo
End Property

Public Property Let reference(vData As String)
    mvarReference = vData
End Property

Public Property Get reference() As String
    reference = mvarReference
End Property

Public Property Let DateTaken(vData As Date)
    mvarDateTaken = vData
End Property

Public Property Get DateTaken() As Date
    DateTaken = mvarDateTaken
End Property

Public Property Let TakenBy(vData As String)
    mvarTakenBy = vData
End Property

Public Property Get TakenBy() As String
    TakenBy = mvarTakenBy
End Property

Public Property Let CustomerNo(vData As Integer)
    mvarCustomerNo = vData
End Property

Public Property Get CustomerNo() As Integer
    CustomerNo = mvarCustomerNo
End Property

Public Property Let RequestedBy(vData As String)
    mvarRequestedBy = vData
End Property

Public Property Get RequestedBy() As String
    RequestedBy = mvarRequestedBy
End Property

Public Property Let RequestedByID(vData As Long)
    mvarRequestedByID = vData
End Property

Public Property Get RequestedByID() As Long
    RequestedByID = mvarRequestedByID
End Property

Public Property Get RequestedByFirstName() As String
    RequestedByFirstName = IIf(InStr(Trim(mvarRequestedBy), " ") = 0, Format(Left(mvarRequestedBy, 1), ">") & Format(Mid(Trim(mvarRequestedBy), 2), "<"), Format(Left(mvarRequestedBy, 1), ">") & Format(Mid(Left(Trim(mvarRequestedBy), InStr(Trim(mvarRequestedBy), " ")), 2), "<"))
End Property

Public Property Let RequestorsTelNo(vData As String)
    mvarRequestorsTelNo = vData
End Property

Public Property Get RequestorsTelNo() As String
    RequestorsTelNo = mvarRequestorsTelNo
End Property

Public Property Let RequestorsEmail(vData As String)
    mvarRequestorsEmail = vData
End Property

Public Property Get RequestorsEmail() As String
    RequestorsEmail = mvarRequestorsEmail
End Property

Public Property Let DeliveryAddressLine1(vData As String)
    mvarDeliveryAddressLine1 = vData
End Property

Public Property Get DeliveryAddressLine1() As String
    DeliveryAddressLine1 = mvarDeliveryAddressLine1
End Property

Public Property Let DeliveryAddressLine2(vData As String)
    mvarDeliveryAddressLine2 = vData
End Property

Public Property Get DeliveryAddressLine2() As String
    DeliveryAddressLine2 = mvarDeliveryAddressLine2
End Property

Public Property Let DeliveryAddressLine3(vData As String)
    mvarDeliveryAddressLine3 = vData
End Property

Public Property Get DeliveryAddressLine3() As String
    DeliveryAddressLine3 = mvarDeliveryAddressLine3
End Property

Public Property Let DeliveryCountyState(vData As String)
    mvarDeliveryCountyState = vData
End Property

Public Property Get DeliveryCountyState() As String
    DeliveryCountyState = mvarDeliveryCountyState
End Property

Public Property Let DeliveryCountry(vData As String)
    mvarDeliveryCountry = vData
End Property

Public Property Get DeliveryCountry() As String
    DeliveryCountry = mvarDeliveryCountry
End Property

Public Property Let DeliveryPostCode(vData As String)
    mvarDeliveryPostCode = vData
End Property

Public Property Get DeliveryPostCode() As String
    DeliveryPostCode = mvarDeliveryPostCode
End Property

Public Property Let DateRequired(vData As Date)
    mvarDateRequired = vData
End Property

Public Property Get DateRequired() As Date
    DateRequired = mvarDateRequired
End Property

Public Property Let RequestDetails(vData As String)
    mvarRequestDetails = vData
End Property

Public Property Get RequestDetails() As String
    RequestDetails = mvarRequestDetails
End Property

Public Property Let DeliveryCustomerName(vData As String)
    mvarDeliveryCustomerName = vData
End Property

Public Property Get DeliveryCustomerName() As String
    DeliveryCustomerName = mvarDeliveryCustomerName
End Property

Public Property Get WebStatus() As String
    WebStatus = mvarWebStatus
End Property

Friend Property Let WebStatus(vData As String)
    mvarWebStatus = vData
End Property

''Public Property Let NoSearches(vData As Integer)
''    mvarNoSearches = vData
''End Property
''
''Public Property Get NoSearches() As Integer
''    NoSearches = mvarNoSearches
''End Property
''
Public Property Let NoPhotos(vData As Integer)
    mvarNoPhotos = vData
End Property

Public Property Get NoPhotos() As Integer
    NoPhotos = mvarNoPhotos
End Property

Public Property Let NoConfirmed(vData As Integer)
    mvarNoConfirmed = vData
End Property

Public Property Get NoConfirmed() As Integer
    NoConfirmed = mvarNoConfirmed
End Property

Public Property Let NoConfirmedTrans(vData As Integer)
    mvarNoConfirmedTrans = vData
End Property

Public Property Get NoConfirmedTrans() As Integer
    NoConfirmedTrans = mvarNoConfirmedTrans
End Property

Public Property Let NoConfirmedDigital(vData As Integer)
    mvarNoConfirmedDigital = vData
End Property

Public Property Get NoConfirmedDigital() As Integer
    NoConfirmedDigital = mvarNoConfirmedDigital
End Property

Public Property Get WebDatePosted() As Date
    WebDatePosted = mvarWebDatePosted
End Property

Friend Property Let WebDatePosted(vData As Date)
    mvarWebDatePosted = vData
End Property

Public Property Get WebDatePostedFormatted() As String
    If mvarWebDatePosted = 0 Then
        WebDatePostedFormatted = ""
    Else
        WebDatePostedFormatted = Format(mvarWebDatePosted, "dd/mm/yyyy")
    End If
End Property

Public Property Get DispatchTransparency() As String
    DispatchTransparency = mvarDispatchTransparency
End Property

Friend Property Let DispatchTransparency(vData As String)
    mvarDispatchTransparency = vData
End Property

Public Property Get DispatchDigital() As String
    DispatchDigital = mvarDispatchDigital
End Property

Friend Property Let DispatchDigital(vData As String)
    mvarDispatchDigital = vData
End Property

Public Property Let ServiceFee(vData As Currency)
    mvarServiceFee = vData
End Property

Public Property Get NoPhotosRequested() As Integer
    NoPhotosRequested = mvarNoPhotosRequested
End Property

Public Property Let NoPhotosRequested(vData As Integer)
    mvarNoPhotosRequested = vData
End Property

Public Property Get Searches() As Searches
    If mvarSearches Is Nothing Then
        Set mvarSearches = New Searches
        mvarSearches.JobNo = mvarJobNo
    End If

    Set Searches = mvarSearches
End Property

Public Property Set Searches(vData As Searches)
    Set mvarSearches = vData
End Property

Private Function GetDigitalSearchesFolder() As String
    'GetDigitalSearchesFolder = App.Path & "\" & DIGITAL_SEARCHES_FOLDER
    GetDigitalSearchesFolder = goSystemConfig.ServerLocation & "\" & DIGITAL_SEARCHES_FOLDER
End Function

Public Property Get DigitalSearchesFolder() As String
    DigitalSearchesFolder = GetDigitalSearchesFolder
End Property

Public Property Get DigitalDeliveryNoteFolder() As String
    If mvarDeliveryNoteNo = 0 Then
        DigitalDeliveryNoteFolder = ""
    Else
        DigitalDeliveryNoteFolder = DigitalSearchesFolder & "\DNote '" & CStr(mvarDeliveryNoteNo) & "'"
    End If
End Property

Public Property Get DigitalDeliveryNote() As String
    If mvarDeliveryNoteNo = 0 Then
        DigitalDeliveryNote = ""
    Else
        DigitalDeliveryNote = DigitalDeliveryNoteFolder & "\" & CStr(mvarDeliveryNoteNo) & ".pdf"
    End If
End Property

Private Function Refresh(ByRef rstRecord As DAO.Recordset)
'***************************************
' Module/Form Name   : Job2
'
' Procedure Name     : Refresh
'
' Purpose            :
'
' Date Created       : 05/04/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : rstRecord - ADODB.Recordset
'
' Returns            :
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 05/04/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo Refresh_Error
    '
    '******** Code Starts Here *************
    '
    With rstRecord
        mvarJobNo = NVL(!job_no, NVLNumeric)
        mvarDeliveryNoteNo = NVL(!delivery_note_no, NVLNumeric)
        mvarReference = NVL(!job_reference, NVLString)
        mvarDateTaken = NVL(!date_taken, NVLNumeric)
        mvarTakenBy = NVL(!taken_by, NVLString)
        mvarCustomerNo = NVL(!customer_no, NVLNumeric)
        mvarRequestedBy = NVL(!requested_by, NVLString)
        mvarRequestedByID = NVL(!RequestedByID, NVLNumeric)
        mvarRequestorsTelNo = NVL(!requestors_telno, NVLString)
        mvarDeliveryAddressLine1 = NVL(!del_address_line_1, NVLString)
        mvarDeliveryAddressLine2 = NVL(!del_address_line_2, NVLString)
        mvarDeliveryAddressLine3 = NVL(!del_address_line_3, NVLString)
        mvarDeliveryCountyState = NVL(!del_county_or_state, NVLString)
        mvarDeliveryCountry = NVL(!del_country, NVLString)
        mvarDeliveryPostCode = NVL(!del_post_code, NVLString)
        mvarDateRequired = NVL(!date_required, NVLNumeric)
        mvarRequestDetails = NVL(!request_details, NVLString)
        mvarDeliveryCustomerName = NVL(!del_customer_name, NVLString)
        mvarRequestorsEmail = NVL(!RequestorsEmail, NVLString)
        mvarWebStatus = NVL(!WebStatus, NVLString)
        mvarWebDatePosted = NVL(!WebDatePosted, NVLNumeric)
        mvarNoPhotosRequested = NVL(!NoPhotosRequested, NVLNumeric)
        mvarPreferredPhotoType = NVL(!PreferredPhotoType, NVLString)
    End With
    '
    '********* Code Ends Here **************
    '
    Exit Function
    '
Refresh_Error:
    ErrorRaise "Job2.Refresh"
End Function

Public Sub Read(Optional ByVal lngJobNo As Long = 0, _
                Optional ByVal lngDNoteNo As Long = 0)
    '***************************************
    ' Module/Form Name   : Job2
    '
    ' Procedure Name     : Read
    '
    ' Purpose            :
    '
    ' Date Created       : 05/04/2002
    '
    ' Author             : GARETH SAUNDERS
    '
    ' Parameters         : JobNo - Long
    '
    ' Amendment History  : Date       Author    Description
    '                    : --------------------------------
    '                    : 05/04/2002 GARETH SAUNDERS
    '
    '***************************************
    '
    On Error GoTo Read_Error
    '
    '******** Code Starts Here *************
    '

    Dim strSQL As String
    Dim rstJob As DAO.Recordset
    
    If lngJobNo = 0 And lngDNoteNo = 0 Then
        Err.Raise vbObjectError + 13, , "Internal software error: JobNo and DNoteNo cannot both be zero."
    End If
    
    strSQL = "SELECT * FROM JOB WHERE "
    If lngJobNo <> 0 Then
        strSQL = strSQL & "JOB_NO = " & lngJobNo
    Else
        strSQL = strSQL & "DELIVERY_NOTE_NO = " & lngDNoteNo
    End If
    Set rstJob = db.OpenRecordset(strSQL, dbOpenSnapshot)
    
    Refresh rstJob
    '
    '   Set up the Job Stats.
    '
    SetJobStats
    '
    '********* Code Ends Here **************
    '
    Exit Sub
    '
Read_Error:
    ErrorRaise "Job2.Read"
End Sub

Private Sub SetJobStats()
'***************************************
' Module/Form Name   : Job2
'
' Procedure Name     : SetJobStats
'
' Purpose            :
'
' Date Created       : 01/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 01/12/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo SetJobStats_Error
'
'******** Code Starts Here *************
'
    Dim strSQL As String
''    Dim rstJobStats As ADODB.Recordset
    Dim rstJobStats As DAO.Recordset
    '
    '   Total Number of Photos on Job.
    '
''    Set rstJobStats = New ADODB.Recordset
    strSQL = "SELECT count(search_result.search_result_key) AS NoPhotos "
    strSQL = strSQL & "FROM search, search_result, photograph "
    strSQL = strSQL & "WHERE search.job_no = " & CStr(mvarJobNo) & " "
    strSQL = strSQL & "and search.search_no = search_result.search_no and photograph.photograph_key = search_result.photograph_key and photograph.deleted = false"
''    rstJobStats.Open strSQL, gdbADO, adOpenStatic, adLockOptimistic
    Set rstJobStats = db.OpenRecordset(strSQL, dbOpenSnapshot)
    mvarNoPhotos = rstJobStats!NoPhotos
    rstJobStats.Close
    Set rstJobStats = Nothing
    '
    '   Number of Confirmed Photos on Job.
    '
''    Set rstJobStats = New ADODB.Recordset
    strSQL = "SELECT count(search_result.search_result_key) AS NoConfirmed, "
    strSQL = strSQL & "sum(iif(search.phototype='T',1,0)) AS NoConfirmedTrans, "
    strSQL = strSQL & "sum(iif(search.phototype='D',1,0)) AS NoConfirmedDigital "
    strSQL = strSQL & "FROM search, search_result, photograph "
    strSQL = strSQL & "WHERE search.job_no = " & CStr(mvarJobNo) & " "
    strSQL = strSQL & "and search.search_no = search_result.search_no and photograph.photograph_key = search_result.photograph_key and photograph.deleted = false "
    strSQL = strSQL & "and search_result.confirmed = true"
''    rstJobStats.Open strSQL, gdbADO, adOpenStatic, adLockOptimistic
    Set rstJobStats = db.OpenRecordset(strSQL, dbOpenSnapshot)
    mvarNoConfirmed = NVL(rstJobStats!NoConfirmed, NVLNumeric)
    mvarNoConfirmedTrans = NVL(rstJobStats!NoConfirmedTrans, NVLNumeric)
    mvarNoConfirmedDigital = NVL(rstJobStats!NoConfirmedDigital, NVLNumeric)
    rstJobStats.Close
    Set rstJobStats = Nothing
    '
    '********* Code Ends Here **************
    '
    Exit Sub
    '
SetJobStats_Error:
    ErrorRaise "Job2.SetJobStats"
End Sub

Public Sub update()
'***************************************
' Module/Form Name   : Job2
'
' Procedure Name     : update
'
' Purpose            :
'
' Date Created       : 05/04/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 05/04/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo update_Error
'
'******** Code Starts Here *************
'
    Dim strSQL      As String
    Dim rstJob      As DAO.Recordset
    
    strSQL = "SELECT * FROM JOB WHERE JOB_NO = " & mvarJobNo
    
    Set rstJob = db.OpenRecordset(strSQL, dbOpenDynaset)
''    With mrstJob
    With rstJob
        .Edit
        !delivery_note_no = mvarDeliveryNoteNo
        !job_reference = mvarReference
        !date_taken = mvarDateTaken
        !taken_by = mvarTakenBy
        !customer_no = mvarCustomerNo
        !requested_by = mvarRequestedBy
        !RequestedByID = mvarRequestedByID
        !requestors_telno = mvarRequestorsTelNo
        !del_address_line_1 = mvarDeliveryAddressLine1
        !del_address_line_2 = mvarDeliveryAddressLine2
        !del_address_line_3 = mvarDeliveryAddressLine3
        !del_county_or_state = mvarDeliveryCountyState
        !del_country = mvarDeliveryCountry
        !del_post_code = mvarDeliveryPostCode
        !date_required = mvarDateRequired
        !request_details = mvarRequestDetails
        !del_customer_name = mvarDeliveryCustomerName
        !RequestorsEmail = mvarRequestorsEmail
        !NoPhotosRequested = mvarNoPhotosRequested
        !WebStatus = mvarWebStatus
        !WebDatePosted = mvarWebDatePosted
        !PreferredPhotoType = mvarPreferredPhotoType
        .update
''        Set .ActiveConnection = gdbADO
''        On Error Resume Next
''        .UpdateBatch
''        If Err.Number = -2147217864 Then
''            On Error GoTo Update_Error
''            Err.Raise vbObjectError + 12, , "Job has been updated since first retrieving it."
''        ElseIf Err.Number = 0 Then
''            On Error GoTo Update_Error
''        Else
''            ErrorSave
''            On Error GoTo Update_Error
''            ErrorRestore
''        End If
''        Set .ActiveConnection = Nothing
    End With
    goLog.WriteDetailedLog CStr(mvarJobNo), "Job Updated"
    '
    '********* Code Ends Here **************
    '
    Exit Sub
        '
update_Error:
    ErrorRaise "Job2.update"
End Sub

Public Function Delete() As Boolean
'***************************************
' Module/Form Name   : Job2
'
' Procedure Name     : Delete
'
' Purpose            :
'
' Date Created       : 09/04/2002
'
' Author             : GARETH SAUNDERS
'
' Returns            : Boolean
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 11/04/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo Delete_Error
'
'******** Code Starts Here *************
'
    Dim strSQL As String
    Dim lngRecordsAffected As Long
    Dim recDelete       As DAO.Recordset
    
    '   Initialise the result.
    Delete = False

'    strSQL = "DELETE * FROM JOB " _
'             & "WHERE JOB_NO = " & CStr(mvarJobNo)
    strSQL = "SELECT * FROM JOB " _
             & "WHERE JOB_NO = " & CStr(mvarJobNo)
        
    Set recDelete = db.OpenRecordset(strSQL, dbOpenDynaset)
    If recDelete.EOF Then
        Err.Raise vbObjectError + 3, "Job: Delete", "Job does not Exist. It may have been deleted by another user."
    Else
        On Error Resume Next
        recDelete.Delete
        If Err.Number = 3200 Then
            On Error GoTo Delete_Error
            Err.Raise vbObjectError + 4, , "Job has an active history and cannot be deleted"
        ElseIf Err.Number <> 0 Then
            ErrorSave
            On Error GoTo Delete_Error
            ErrorRestore
        Else
            On Error GoTo Delete_Error
        End If
    End If
    
    '   Set the result.
    Delete = True
    goLog.WriteDetailedLog CStr(mvarJobNo), "Job Deleted"
    Exit Function
    '
    '********* Code Ends Here **************
    '
    Exit Function
    '
Delete_Error:
    ErrorRaise "Job2.Delete"
End Function

Public Sub Create(strDeliveryNoteNo As Long, _
                  strReference As String, _
                  dteDateTaken As Date, _
                  strTakenBy As String, _
                  intCustomerNo As Integer, _
                  poContact As Contact, _
                  strRequestorsTelNo As String, _
                  strDeliveryAddressLine1 As String, _
                  strDeliveryAddressLine2 As String, _
                  strDeliveryAddressLine3 As String, _
                  strDeliveryCountyState As String, _
                  strDeliveryCountry As String, _
                  strDeliveryPostCode As String, _
                  dteDateRequired As Date, _
                  strRequestDetails As String, _
                  strDeliveryCustomerName As String, _
                  strRequestorsEmail As String)
'***************************************
' Module/Form Name   : Job2
'
' Procedure Name     : create
'
' Purpose            :
'
' Date Created       : 08/04/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : strDeliveryNoteNo - Long
'                    : strReference - String
'                    : dteDateTaken - Date
'                    : strTakenBy - String
'                    : intCustomerNo - Integer
'                    : strRequestedBy - String
'                    : strRequestorsTelNo - String
'                    : strDeliveryAddressLine1 - String
'                    : strDeliveryAddressLine2 - String
'                    : strDeliveryAddressLine3 - String
'                    : strDeliveryCountyState - String
'                    : strDeliveryCountry - String
'                    : strDeliveryPostCode - String
'                    : dteDateRequired - Date
'                    : strRequestDetails - String
'                    : strDeliveryCustomerName - String
'                    : strRequestorsEmail - String
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 08/04/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo Create_Error
'
'******** Code Starts Here *************
'
    Dim strSQL      As String
    Dim rstJob      As DAO.Recordset
    Dim oCustomer   As Customer2
    '
    '   This technique is required to retrieve the Job Number afterwards.
    '
    strSQL = "SELECT TOP 1 * FROM JOB WHERE 1 = 2"
''    Set rstJob = New ADODB.Recordset
''    rstJob.Open "Job", gdbADO, adOpenKeyset, adLockPessimistic
    Set rstJob = db.OpenRecordset(strSQL, dbOpenDynaset)
''    With rstJob
    With rstJob
        .AddNew
        !delivery_note_no = strDeliveryNoteNo
        !job_reference = strReference
        !date_taken = dteDateTaken
        !taken_by = strTakenBy
        !customer_no = intCustomerNo
        !requested_by = poContact.Name
        !RequestedByID = poContact.ID
        !requestors_telno = strRequestorsTelNo
        !del_address_line_1 = strDeliveryAddressLine1
        !del_address_line_2 = strDeliveryAddressLine2
        !del_address_line_3 = strDeliveryAddressLine3
        !del_county_or_state = strDeliveryCountyState
        !del_country = strDeliveryCountry
        !del_post_code = strDeliveryPostCode
        !date_required = dteDateRequired
        !request_details = strRequestDetails
        !del_customer_name = strDeliveryCustomerName
        !RequestorsEmail = strRequestorsEmail
        !WebStatus = ""
        !WebDatePosted = Null
        !PreferredPhotoType = mvarPreferredPhotoType
        On Error Resume Next
        .update
    End With
    If Err.Number = 3022 Then
        Set oCustomer = New Customer2
        oCustomer.Read intCustomerNo
        Err.Raise vbObjectError + 1, , "Job number " & strReference & " already exists for Customer: " & oCustomer.CustomerName
    ElseIf Err.Number <> 0 Then
        ErrorSave
        On Error GoTo Create_Error
        ErrorRestore
    Else
        On Error GoTo Create_Error
        rstJob.Bookmark = rstJob.LastModified
    End If
    '
    '   Set up the internal values.
    '
    mvarJobNo = rstJob!job_no
    mvarDeliveryNoteNo = strDeliveryNoteNo
    mvarReference = strReference
    mvarDateTaken = dteDateTaken
    mvarTakenBy = strTakenBy
    mvarCustomerNo = intCustomerNo
    mvarRequestedBy = poContact.Name
    mvarRequestedByID = poContact.ID
    mvarRequestorsTelNo = strRequestorsTelNo
    mvarDeliveryAddressLine1 = strDeliveryAddressLine1
    mvarDeliveryAddressLine2 = strDeliveryAddressLine2
    mvarDeliveryAddressLine3 = strDeliveryAddressLine3
    mvarDeliveryCountyState = strDeliveryCountyState
    mvarDeliveryCountry = strDeliveryCountry
    mvarDeliveryPostCode = strDeliveryPostCode
    mvarDateRequired = dteDateRequired
    mvarRequestDetails = strRequestDetails
    mvarDeliveryCustomerName = strDeliveryCustomerName
    mvarRequestorsEmail = strRequestorsEmail
    mvarWebStatus = ""
    mvarWebDatePosted = 0
    mvarNoPhotosRequested = 0
    rstJob.Close
    Set rstJob = Nothing
    goLog.WriteDetailedLog CStr(mvarJobNo), "Job Created"
    '
    '********* Code Ends Here **************
    '
    Exit Sub
    '
Create_Error:
    ErrorRaise "Job2.create"
End Sub

Public Sub PostToWeb()
'***************************************
' Module/Form Name   : Job2
'
' Procedure Name     : PostToWeb
'
' Purpose            :
'
' Date Created       : 18/04/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 18/04/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo PostToWeb_Error
'
'******** Code Starts Here *************
'
''    Dim oDNCFTP As DNCFunctions.FTP
    Dim oDNCFTP As Object
    '
    '   Check there is a connection to the internet first.
    '
''    Set oDNCFTP = New DNCFunctions.FTP
    Set oDNCFTP = CreateObject("DNCFunctions.FTP")
    If Not oDNCFTP.IsNetConnectionAlive Then
        Set oDNCFTP = Nothing
        Err.Raise vbObjectError + 11, , "There is no current connection to the internet. The job cannot be posted"
    End If
    Set oDNCFTP = Nothing
    '
    '   Refresh the Job Record.
    '
    Read mvarJobNo
    '
''    gdbADO.BeginTrans
    begin_trans
        '
        '   Web Status must be Unset.
        '
''        If Format(mrstJob!WebStatus) <> "" Then
        If Format(mvarWebStatus) <> "" Then
''            Err.Raise vbObjectError + 6, , "Web Status is " & mrstJob!WebStatus & ". Job cannot be posted to Web"
            Err.Raise vbObjectError + 6, , "Web Status is " & mvarWebStatus & ". Job cannot be posted to Web"
        End If
        '
        '   Web Posted Date also must not be set.
        '
''        If mrstJob!WebDatePosted <> 0 Then
        If mvarWebDatePosted <> 0 Then
''            Err.Raise vbObjectError + 7, , "Web Posted Date is " & Format(mrstJob!WebDatePosted, "DD/MM/YYYY") & ". Job cannot be posted to Web"
            Err.Raise vbObjectError + 7, , "Web Posted Date is " & Format(mvarWebDatePosted, "DD/MM/YYYY") & ". Job cannot be posted to Web"
        End If
        '
        mvarWebStatus = "P"
        mvarWebDatePosted = Now
        update
        '
        '   Output the details.
        '
        OutputWebDetails
        '
''    gdbADO.CommitTrans
    commit_trans
    '
    '********* Code Ends Here **************
    '
    Exit Sub
    '
PostToWeb_Error:
    ErrorRaise "Job2.PostToWeb"
End Sub

Private Sub OutputWebDetails()
    '    GHS (26/02/2005) - 4.3 Enhancements Section 2.5
'***************************************
' Module/Form Name   : Job2
'
' Procedure Name     : OutputWebDetails
'
' Purpose            :
'
' Date Created       : 25/10/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo OutputWebDetails_Error
'
'******** Code Starts Here *************
'
    Dim strJobsFileWithPath As String
    Dim strSearchesFileWithPath As String
    Dim strPhotographsFileWithPath As String
    Dim strHeading As String
    Dim strDetail As String
''    Dim rstPhotographs As ADODB.Recordset
''    Dim rstSearches As ADODB.Recordset
    Dim rstPhotographs As DAO.Recordset
    Dim rstSearches As DAO.Recordset
    Dim strSQL As String
''    Dim oDNCFTP As DNCFunctions.FTP
    Dim oDNCFTP As Object
    Dim strWebDestination As String
    Dim oSearch         As Search2
    '
    '   Output the Headings.
    '
    'strJobsFileWithPath = App.Path & "\Web Jobs\" & CONST_FILENAME_WEBJOBS
    strJobsFileWithPath = goSystemConfig.ServerLocation & "\Web Jobs\" & CONST_FILENAME_WEBJOBS
    Close #1
    Open strJobsFileWithPath For Output As #1
    strHeading = """JobNo"",""Reference"",""UseByDate"",""CustomerName"",""EmailAddress"""
    Print #1, strHeading
    Close #1
    '
    'strSearchesFileWithPath = App.Path & "\Web Jobs\" & CONST_FILENAME_WEBSEARCHES
    strSearchesFileWithPath = goSystemConfig.ServerLocation & "\Web Jobs\" & CONST_FILENAME_WEBSEARCHES
    Close #1
    Open strSearchesFileWithPath For Output As #1
    strHeading = """JobNo"",""SearchNo"",""Description"",""PhotoType"",""UseByDate"""
    Print #1, strHeading
    Close #1
    '
    'strPhotographsFileWithPath = App.Path & "\Web Jobs\" & CONST_FILENAME_WEBPHOTOGRAPHS
    strPhotographsFileWithPath = goSystemConfig.ServerLocation & "\Web Jobs\" & CONST_FILENAME_WEBPHOTOGRAPHS
    Close #1
    Open strPhotographsFileWithPath For Output As #1
    strHeading = """SearchNo"",""PhotoNo"",""Description"",""Credit"""
    Print #1, strHeading
    Close #1
    '
    '   Output the Job details.
    '
    Close #1
    Open strJobsFileWithPath For Append As #1
    strDetail = CStr(mvarJobNo) & ",""" & _
                mvarReference & """," & _
                Format(WebSearchUseByDate, "dd-mmm-yyyy") & ",""" & _
                mvarDeliveryCustomerName & """," & _
                goSystemConfig.WebSearchTestEmail & """"
    Print #1, strDetail
    Close #1
    '
''    '   Output all the Searches for the Job.
''    '
''    strSQL = "SELECT "
''    strSQL = strSQL & "Job.Job_No , Search.Search_No, Search.Search_Description, Search.PhotoType "
''    strSQL = strSQL & "FROM Job INNER JOIN Search ON Job.job_no = Search.job_no "
''    strSQL = strSQL & "WHERE job.job_no = " & mvarJobNo
''    '
''    Set rstSearches = db.OpenRecordset(strSQL, dbOpenSnapshot)
''    With rstSearches
    Open strSearchesFileWithPath For Append As #1
    For Each oSearch In mvarSearches
        strDetail = CStr(mvarJobNo) & "," & _
                    CStr(oSearch.SearchNo) & ",""" & _
                    oSearch.Description & """,""" & _
                    oSearch.PhotoType & """," & _
                    Format(oSearch.WebSearchUseByDate(mvarWebDatePosted), "dd-mmm-yyyy")
        Print #1, strDetail
    Next oSearch
        
''        While Not .EOF
''            strDetail = CStr(mvarJobNo) & "," & _
''                        CStr(!search_no) & ",""" & _
''                        !Search_Description & """,""" & _
''                        !PhotoType & """"
''            Print #1, strDetail
''            .MoveNext
''        Wend
    Close #1
''    End With
    '
    '   Output all the Photographs for the Job.
    '
    strSQL = "SELECT "
    strSQL = strSQL & "Search.Search_No, search_result.batch_no, Batch.description, Batch.Credit "
    strSQL = strSQL & "FROM ((Job INNER JOIN Search ON Job.job_no = Search.job_no) INNER JOIN "
    strSQL = strSQL & "Search_result ON Search.search_no = Search_result.search_no) INNER JOIN "
    strSQL = strSQL & "Batch ON Search_result.batch_no = Batch.batch_no "
    strSQL = strSQL & "WHERE job.job_no = " & mvarJobNo
    '
''    Set rstPhotographs = New ADODB.Recordset
''    rstPhotographs.Open strSQL, gdbADO, adOpenForwardOnly
    Set rstPhotographs = db.OpenRecordset(strSQL, dbOpenSnapshot)
    With rstPhotographs
        Open strPhotographsFileWithPath For Append As #1
        While Not .EOF
            strDetail = CStr(!search_no) & "," & _
                        CStr(!batch_no) & ",""" & _
                        !Description & """,""" & _
                        !Credit & """"
            Print #1, strDetail
            .MoveNext
        Wend
        Close #1
    End With
    '
    '   Move the Files to the Web.
    '
''    Set oDNCFTP = New DNCFunctions.FTP
    Set oDNCFTP = CreateObject("DNCFunctions.FTP")
    With oDNCFTP
        .OpenConnection goSystemConfig.FTPServer, goSystemConfig.FTPUser, goSystemConfig.FTPPassword
        If goSystemConfig.WebSearchTestEmail = "" Then      '   This is production!
            strWebDestination = "WebSearches/jobs/"
        Else
            strWebDestination = "WebSearches/test/jobs/"    '   This is a test job.
        End If
        .PostWebFile strJobsFileWithPath, CONST_FILENAME_WEBJOBS, strWebDestination
        .PostWebFile strSearchesFileWithPath, CONST_FILENAME_WEBSEARCHES
        .PostWebFile strPhotographsFileWithPath, CONST_FILENAME_WEBPHOTOGRAPHS
        .CloseConnection
    End With
    '
    '   Call an XML method to import the data to the database on the web server.
    '
    ImportDataToWebDatabase mvarJobNo
'
'********* Code Ends Here **************
'
   Exit Sub
'
OutputWebDetails_Error:
    ErrorRaise "Job2.OutputWebDetails"
End Sub

Private Sub ImportDataToWebDatabase(Optional JobNo As Long = 0)
'***************************************
' Module/Form Name   : Job2
'
' Procedure Name     : ImportDataToWebDatabase
'
' Purpose            :
'
' Date Created       : 26/10/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : JobNo - Long
'                    : = -
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo ImportDataToWebDatabase_Error
'
'******** Code Starts Here *************
'

    Dim objHTTP As New MSXML2.XMLHTTP
    Dim strEnvelope As String
    Dim strReturn As String
    Dim objReturn As New MSXML2.DOMDocument30
    Dim objNode As MSXML2.IXMLDOMNode
    Dim strQuery As String
    Dim strSuccess As String
    Dim fError As frmHTMLError
    Dim strError As String
    Dim strImportLocation    As String
    
    If goSystemConfig.WebSearchTestEmail = "" Then
        strImportLocation = "http://" & goSystemConfig.PostWebAddress & "/websearches/importdata.asp"
    Else
        strImportLocation = "http://" & goSystemConfig.PostWebAddress & "/websearches/test/importdata.asp"
    End If
    
    strEnvelope = "<SOAP:Envelope xmlns:SOAP=""urn:schemas-xmlsoap-org:soap.v1"">" & _
                  "<SOAP:Body>" & _
                  "<m:IWImportData xmlns:m=""urn:imagewhere/soap:ImportData"">" & _
                  "<JobNo>" & CStr(JobNo) & "</JobNo>" & _
                  "<WebSearchesEmailTo>" & CStr(goSystemConfig.WebSearchesEmailTo) & "</WebSearchesEmailTo>" & _
                  "<URL>" & CStr(goSystemConfig.PostWebAddress) & "</URL>" & _
                  "<WebSearchesEmailFrom>" & CStr(goSystemConfig.WebSearchesEmailFrom) & "</WebSearchesEmailFrom>" & _
                  "</m:IWImportData>" & _
                  "</SOAP:Body>" & _
                  "</SOAP:Envelope>"

    With objHTTP
        .Open "post", strImportLocation, False
        .setRequestHeader "Content-Type", "text/xml"
        .setRequestHeader "SOAPMethodName", "urn:imagewhere/soap:ImportData#IWImportData"
        .Send strEnvelope
        strReturn = .responseText
    End With

    objReturn.loadXML strReturn

    strQuery = "SOAP:Envelope/SOAP:Body/m:IWImportData/Result"

    strSuccess = "Fail"
    On Error Resume Next
    strSuccess = objReturn.selectSingleNode(strQuery).Text
    On Error GoTo ImportDataToWebDatabase_Error
    If strSuccess <> "Success" Then
        Set fError = New frmHTMLError
        fError.Display strReturn
        strError = fError.ErrorText
        Set fError = Nothing
        Screen.MousePointer = vbHourglass
        Err.Raise vbObjectError + 8, , "Web Server error importing Search Data for job: " & CStr(mvarJobNo) & " - " & strError
    End If
'
'********* Code Ends Here **************
'
   Exit Sub
'
ImportDataToWebDatabase_Error:
    ErrorRaise "Job2.ImportDataToWebDatabase"
End Sub

Public Sub ReceiveWebFeedback(intNoPhotosRequested As Integer)
    '***************************************
    ' Module/Form Name   : Job2
    '
    ' Procedure Name     : ReceiveWebFeedback
    '
    ' Purpose            :
    '
    ' Date Created       : 09/06/2002
    '
    ' Author             : GARETH SAUNDERS
    '
    ' Parameters         : intNoPhotosRequested - Integer
    '
    ' Amendment History  : Date       Author    Description
    '                    : --------------------------------
    '                    : 09/06/2002 GARETH SAUNDERS
    '
    '***************************************
    '
    On Error GoTo ReceiveWebFeedback_Error
    '
    '******** Code Starts Here *************
    '
    Dim strSQL As String

    strSQL = "UPDATE JOB SET "
    strSQL = strSQL & "WEBSTATUS = 'R', "
    strSQL = strSQL & "NOPHOTOSREQUESTED = " & CStr(intNoPhotosRequested) & " "
    strSQL = strSQL & "WHERE JOB_NO = " & CStr(mvarJobNo)
    '
''    gdbADO.Execute strSQL
    db.Execute strSQL

    mvarWebStatus = "R"
    mvarNoPhotosRequested = intNoPhotosRequested
    '
    '********* Code Ends Here **************
    '
    Exit Sub
    '
ReceiveWebFeedback_Error:
    ErrorRaise "Job2.ReceiveWebFeedback"
End Sub

Public Property Get Contact() As Contact
'***************************************
' Module/Form Name   : Job2
'
' Procedure Name     : Contact
'
' Purpose            :
'
' Date Created       : 21/04/2002
'
' Author             : GARETH SAUNDERS
'
' Returns            : Contact
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
'******** Code Starts Here *************
'
    Dim oContact As Contact
    Dim blnFound As Boolean
    Dim oCustomer As Customer2
    
'    oContact.Read mvarCustomerNo, mvarRequestedBy
    '
    '   ID may not be set up for historical records or if the contact no long exists.
    '
    If mvarRequestedByID <> 0 Then
        On Error Resume Next
        Set oContact = New Contact
        oContact.Read mvarRequestedByID
        If Err.Number = 0 Then
            On Error GoTo 0
            Set Contact = oContact
        ElseIf Err.Number <> vbObjectError + 5 Then
            ErrorSave
            On Error GoTo 0
            ErrorRestore
        Else
            On Error GoTo 0
        End If
    Else
        Set oCustomer = New Customer2
        oCustomer.Read mvarCustomerNo
        blnFound = False
        For Each oContact In oCustomer.Contacts
            If LCase(oContact.Name) = LCase(mvarRequestedBy) Then
                blnFound = True
                Exit For
            End If
        Next oContact
        Set oCustomer = Nothing
        '
        If blnFound Then
            Set Contact = oContact
        Else
            Err.Raise vbObjectError + 9, , "Contact not found"
        End If
    End If
    '
'
'********* Code Ends Here **************
'
End Property

Public Property Get CancellationHTMLEmail() As String
    Dim strEmailBody As String
    
    strEmailBody = goSystemConfig.CancellationHTMLEmail
    '
    '   Create the body of the email.
    '
    strEmailBody = Replace(strEmailBody, "<Reference>", mvarReference)
    strEmailBody = Replace(strEmailBody, "<RequestedBy>", RequestedByFirstName)
    strEmailBody = Replace(strEmailBody, "<TakenBy>", mvarTakenBy)
    strEmailBody = Replace(strEmailBody, "<Telephone>", goCompanyInfo.TelNo)
    strEmailBody = Replace(strEmailBody, "<Email>", goCompanyInfo.Email)
    
    CancellationHTMLEmail = strEmailBody
End Property
''
''Public Property Get ConfirmationEmail() As String
''          Dim strEmailBody As String
''
''          strEmailBody = "Reference: " & mvarReference & vbCrLf & vbCrLf
''          strEmailBody = strEmailBody & "Dear " & IIf(InStr(Trim(mvarRequestedBy), " ") = 0, Format(Left(mvarRequestedBy, 1), ">") & Format(Mid(Trim(mvarRequestedBy), 2), "<"), Format(Left(mvarRequestedBy, 1), ">") & Format(Mid(Left(Trim(mvarRequestedBy), InStr(Trim(mvarRequestedBy), " ")), 2), "<")) & vbCrLf & vbCrLf
''          strEmailBody = strEmailBody & "We have now selected the transparencies for you and will dispatch " & mvarNoConfirmed & " today."
''          strEmailBody = strEmailBody & "We apologise if we have been unable to fulfil your request "
''          strEmailBody = strEmailBody & "completely.  Please do contact us if you wish us to look further for you.  (Should you wish to "
''          strEmailBody = strEmailBody & "download screen resolution images for layout purposes, the site " & WebSearchURL & " "
''          strEmailBody = strEmailBody & "will be available for you to review until " & Format(WebSearchUseByDate, "dd/mm/yyyy") & "."
''          strEmailBody = strEmailBody & vbCrLf & vbCrLf
''          strEmailBody = strEmailBody & "The selected transparencies are being dispatched today by " & mvarDispatchTransparency & " to " & FormattedDeliveryAddress & "."
''          strEmailBody = strEmailBody & "Please advise us immediately if this is incorrect."
''          strEmailBody = strEmailBody & vbCrLf & vbCrLf
''          strEmailBody = strEmailBody & "There is a service fee of " & ", which includes postage."
''          strEmailBody = strEmailBody & vbCrLf & vbCrLf
''          strEmailBody = strEmailBody & "With best wishes"
''          strEmailBody = strEmailBody & vbCrLf & vbCrLf
''          strEmailBody = strEmailBody & mvarTakenBy
''
''          ConfirmationEmail = strEmailBody
''End Property

Public Property Get PostedHTMLEmail() As String
'***************************************
' Module/Form Name   : Job2
'
' Procedure Name     : PostedHTMLEmail
'
' Purpose            :
'
' Date Created       : 21/05/2002
'
' Author             : GARETH SAUNDERS
'
' Returns            : String
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
'******** Code Starts Here *************
'
    Dim strEmailBody As String

    On Error GoTo ErrorProc

    strEmailBody = goSystemConfig.PostedHTMLEmail
    '
    '   Create the body of the email.
    '
    strEmailBody = Replace(strEmailBody, "<Reference>", mvarReference)
    strEmailBody = Replace(strEmailBody, "<RequestedBy>", RequestedByFirstName)
    strEmailBody = Replace(strEmailBody, "<WebSearchUseByDate>", Format(WebSearchUseByDate, "Long Date"))
    strEmailBody = Replace(strEmailBody, "<WebSearchURL>", WebSearchURL)
    strEmailBody = Replace(strEmailBody, "<TakenBy>", mvarTakenBy)
    strEmailBody = Replace(strEmailBody, "<WebAddress>", goCompanyInfo.WebSite)
    strEmailBody = Replace(strEmailBody, "<Telephone>", goCompanyInfo.TelNo)
    strEmailBody = Replace(strEmailBody, "<Email>", goCompanyInfo.Email)

    PostedHTMLEmail = strEmailBody
'
'********* Code Ends Here **************
'
    Exit Sub
ErrorProc:
    ErrorRaise "Job2.PostedHTMLEmail"
End Property

Public Property Get ConfirmationHTMLEmail() As String
    Dim strEmailBody As String

    strEmailBody = goSystemConfig.ConfirmationHTMLEmail
    '
    '   Create the body of the email.
    '
    strEmailBody = Replace(strEmailBody, "<Reference>", mvarReference)
    strEmailBody = Replace(strEmailBody, "<RequestedBy>", RequestedByFirstName)
    strEmailBody = Replace(strEmailBody, "<NoConfirmed>", NoConfirmed)
    strEmailBody = Replace(strEmailBody, "<NoConfirmedTrans>", NoConfirmedTrans)
    strEmailBody = Replace(strEmailBody, "<NoConfirmedDigital>", NoConfirmedDigital)
    strEmailBody = Replace(strEmailBody, "<NoSelected>", mvarNoPhotosRequested)
    strEmailBody = Replace(strEmailBody, "<WebSearchUseByDate>", Format(WebSearchUseByDate, "Long Date"))
    strEmailBody = Replace(strEmailBody, "<WebSearchURL>", WebSearchURL)
    strEmailBody = Replace(strEmailBody, "<DispatchTransparency>", mvarDispatchTransparency)
    strEmailBody = Replace(strEmailBody, "<DeliveryNameAddress>", mvarRequestedBy & ", " & FormattedDeliveryAddress)
    strEmailBody = Replace(strEmailBody, "<ServiceFee>", Format(mvarServiceFee, "£###,##0.00"))
    strEmailBody = Replace(strEmailBody, "<TakenBy>", mvarTakenBy)
    strEmailBody = Replace(strEmailBody, "<WebAddress>", goCompanyInfo.WebSite)
    strEmailBody = Replace(strEmailBody, "<Telephone>", goCompanyInfo.TelNo)
    strEmailBody = Replace(strEmailBody, "<Email>", goCompanyInfo.Email)

    ConfirmationHTMLEmail = strEmailBody
End Property

Public Property Get BundledHTMLEmail() As String
    Dim strEmailBody As String

    strEmailBody = goSystemConfig.BundledHTMLEmail
    '
    '   Create the body of the email.
    '
    strEmailBody = Replace(strEmailBody, "<Reference>", mvarReference)
    strEmailBody = Replace(strEmailBody, "<RequestedBy>", RequestedByFirstName)
    strEmailBody = Replace(strEmailBody, "<NoConfirmed>", NoConfirmed)
    strEmailBody = Replace(strEmailBody, "<NoConfirmedTrans>", NoConfirmedTrans)
    strEmailBody = Replace(strEmailBody, "<NoConfirmedDigital>", NoConfirmedDigital)
    strEmailBody = Replace(strEmailBody, "<NoSelected>", mvarNoPhotosRequested)
    strEmailBody = Replace(strEmailBody, "<WebSearchUseByDate>", Format(WebSearchUseByDate, "Long Date"))
    strEmailBody = Replace(strEmailBody, "<WebSearchURL>", WebSearchURL)
    strEmailBody = Replace(strEmailBody, "<DispatchTransparency>", mvarDispatchTransparency)
    strEmailBody = Replace(strEmailBody, "<DeliveryNameAddress>", mvarRequestedBy & ", " & FormattedDeliveryAddress)
    strEmailBody = Replace(strEmailBody, "<ServiceFee>", Format(mvarServiceFee, "£###,##0.00"))
    strEmailBody = Replace(strEmailBody, "<TakenBy>", mvarTakenBy)
    strEmailBody = Replace(strEmailBody, "<WebAddress>", goCompanyInfo.WebSite)
    strEmailBody = Replace(strEmailBody, "<Telephone>", goCompanyInfo.TelNo)
    strEmailBody = Replace(strEmailBody, "<Email>", goCompanyInfo.Email)

    BundledHTMLEmail = strEmailBody
End Property

Public Property Get WebConvertedReference() As String
    Dim strReference As String
    Dim strConvertedReference As String
    Dim strChar As String
    Dim intChar As Integer
    Dim intCount As Integer
    '
    '   Ignore all characters outside the ranges below 48-57, 65-90 and 97-122 in the ASCII code sequence.
    '   Ignore ampersand and colon however as these mean something in the URL.
    '
    strConvertedReference = ""
    strReference = Left(mvarReference, 20)
    For intCount = 1 To Len(strReference)
        strChar = Mid(strReference, intCount, 1)
        intChar = Asc(strChar)
        If ((intChar >= 48 And intChar <= 57) Or _
            (intChar >= 65 And intChar <= 90) Or _
            (intChar >= 97 And intChar <= 122)) Then
            strConvertedReference = strConvertedReference & strChar
        End If
    Next intCount
    WebConvertedReference = LCase(strConvertedReference)
End Property

Public Property Get WebSearchURL() As String
    WebSearchURL = "http://" & goSystemConfig.PostWebAddress & goSystemConfig.WebLocation & "/?jobref=" & WebConvertedReference & "&jobno=" & CStr(mvarJobNo)
End Property

Public Property Get WebSearchFile() As String
    WebSearchFile = "/websearches/?jobref=" & LCase(Replace(mvarReference, " ", "%20")) & "?jobno=" & CStr(mvarJobNo)
End Property

Public Property Get WebSearchUseByDate() As Date
    Dim intUseByPeriod  As Integer
    Dim oSearch         As Search2
    Dim blnFound        As Boolean
    '
    '   If the job has a Digital Search then the Job's period is the digital one
    '   otherwise it is the Transparency one.
    '
    blnFound = False
    For Each oSearch In mvarSearches
        If oSearch.PhotoType = "D" Then
            blnFound = True
            Exit For
        End If
    Next oSearch
    '
    If blnFound Then
        intUseByPeriod = goSystemConfig.WebUseByPeriodDigital
    Else
        intUseByPeriod = goSystemConfig.WebUseByPeriod
    End If
    '
    WebSearchUseByDate = DateAdd("d", intUseByPeriod, mvarWebDatePosted)
End Property

Public Property Get FormattedDeliveryAddress() As String
    Dim strAddress As String
    
    strAddress = mvarDeliveryCustomerName
    
    If mvarDeliveryAddressLine1 <> "" Then strAddress = strAddress & ", " & mvarDeliveryAddressLine1
    If mvarDeliveryAddressLine2 <> "" Then strAddress = strAddress & ", " & mvarDeliveryAddressLine2
    If mvarDeliveryAddressLine3 <> "" Then strAddress = strAddress & ", " & mvarDeliveryAddressLine3
    If mvarDeliveryCountyState <> "" Then strAddress = strAddress & ", " & mvarDeliveryCountyState
    If mvarDeliveryPostCode <> "" Then strAddress = strAddress & ", " & mvarDeliveryPostCode
    If mvarDeliveryCountry <> "" Then strAddress = strAddress & ", " & mvarDeliveryCountry
  
    FormattedDeliveryAddress = strAddress
End Property

Public Property Get Activities() As Activities
    If mvarActivities Is Nothing Then
        Set mvarActivities = New Activities
    End If

    Set Activities = mvarActivities
End Property

Public Property Set Activities(vData As Activities)
    Set mvarActivities = vData
End Property

Public Property Let PreferredPhotoType(ByVal vData As String)
    mvarPreferredPhotoType = vData
End Property

Public Property Get PreferredPhotoType() As String
    If mvarPreferredPhotoType = "" Then
        PreferredPhotoType = "T"
    Else
        PreferredPhotoType = mvarPreferredPhotoType
    End If
End Property

Public Property Get JobNotes() As Activities
    If mvarJobNotes Is Nothing Then
        Set mvarJobNotes = New Activities
        mvarJobNotes.ActivityType = "JNOT"
        mvarJobNotes.JobNo = mvarJobNo
    End If

    Set JobNotes = mvarJobNotes
End Property

Public Property Set JobNotes(vData As Activities)
    Set mvarJobNotes = vData
End Property

Public Sub PrepareDigitalSearchesForDelivery()
'***************************************
' Module/Form Name   : Job2
'
' Procedure Name     : PrepareDigitalSearchesForDelivery
'
' Purpose            :
'
' Date Created       : 05/07/2005
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo PrepareDigitalSearchesForDelivery_Error
'
'******** Code Starts Here *************
'
    Dim oFSO                    As FileSystemObject
    Dim strImageFolder          As String
    Dim oSearch                 As Search2
    Dim oSearchResult           As SearchResult
    Dim oDigitalImage           As DigitalImage
    Dim blnDigitalImagesExist   As Boolean
    '
    '   Copy images to a Digital Searches folder.
    '
    Set oFSO = New Scripting.FileSystemObject
    '
    '   Does the folder exist.
    '
    'strImageFolder = App.Path & "\DigitalSearches"
    strImageFolder = goSystemConfig.ServerLocation & "\DigitalSearches"
    If Not oFSO.FolderExists(strImageFolder) Then
        oFSO.CreateFolder (strImageFolder)
    End If
    '
    '   Create the Delivery Note folder.
    '
    blnDigitalImagesExist = False
    '
    '   Create the Search Folders.
    '
    For Each oSearch In mvarSearches
        If oSearch.PhotoType = "D" Then
            strImageFolder = RemoveIllegalFileAndFolderCharacters(oSearch.DisplayDescription)
            strImageFolder = oFSO.BuildPath(DigitalDeliveryNoteFolder, "Search '" & strImageFolder & "'")
            '
            '   Copy the images.
            '
            For Each oSearchResult In oSearch.SearchResults
                blnDigitalImagesExist = True
                If Not oFSO.FolderExists(DigitalDeliveryNoteFolder) Then
                    oFSO.CreateFolder (DigitalDeliveryNoteFolder)
                End If
                '
                If Not oFSO.FolderExists(strImageFolder) Then
                    oFSO.CreateFolder (strImageFolder)
                End If
                '
                Set oDigitalImage = New DigitalImage
                oDigitalImage.BatchNo = oSearchResult.BatchNo
                oFSO.CopyFile oDigitalImage.FileLocation("A4"), strImageFolder & "\", True
                Set oDigitalImage = Nothing
            Next oSearchResult
        End If
    Next oSearch
    '
    Set oFSO = Nothing
    If blnDigitalImagesExist Then
        mvarDigitalSearchesFolder = DigitalDeliveryNoteFolder
    End If
'
'********* Code Ends Here **************
'
   Exit Sub
'
PrepareDigitalSearchesForDelivery_Error:
    ErrorRaise "Job2.PrepareDigitalSearchesForDelivery"
End Sub

Private Sub Class_Initialize()
    mvarDigitalSearchesFolder = ""
End Sub
