VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DeliveryNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mvarDNoteNo                 As Long
Private mvarCustomerNo              As Integer
Private mvarCustomerName            As String
Private mvarJobReference            As String
Private mvarStatus                  As String
Private mvarJobNo                   As Long
Private mvarDatePrinted             As Date
Private mvarReturnByDate            As Date
Private mvarTotalOutstandingTrans   As Integer
Private mvarTotalOutstandingDigital As Integer
Private mvarTotalPictures           As Integer
Private mvarServiceFee              As Currency
Private mvarWeeklyHoldingFee        As Currency
Private mvarLossDamageFee           As Currency
Private mvarMemo                    As String
Private mvarDispatchTransparency    As String
Private mvarDispatchDigital         As String
Private mvarSelected                As Boolean
Private mvarSearchResults           As SearchResults

Public Property Get SearchResults() As SearchResults
    If mvarSearchResults Is Nothing Then
        Set mvarSearchResults = New SearchResults
    End If

    Set SearchResults = mvarSearchResults
End Property

Public Property Set SearchResults(vData As SearchResults)
    Set mvarSearchResults = vData
End Property

Public Property Let DNoteNo(ByVal vData As Long)
    mvarDNoteNo = vData
End Property

Public Property Get DNoteNo() As Long
    DNoteNo = mvarDNoteNo
End Property

Public Property Let CustomerNo(ByVal vData As Integer)
    mvarCustomerNo = vData
End Property

Public Property Get CustomerNo() As Integer
    CustomerNo = mvarCustomerNo
End Property

Public Property Let CustomerName(ByVal vData As String)
    mvarCustomerName = vData
End Property

Public Property Get CustomerName() As String
    CustomerName = mvarCustomerName
End Property

Public Property Let JobReference(ByVal vData As String)
    mvarJobReference = vData
End Property

Public Property Get JobReference() As String
    JobReference = mvarJobReference
End Property

Public Property Let Status(ByVal vData As String)
    mvarStatus = vData
End Property

Public Property Get Status() As String
    Status = mvarStatus
End Property

Public Property Let JobNo(ByVal vData As Long)
    mvarJobNo = vData
End Property

Public Property Get JobNo() As Long
    JobNo = mvarJobNo
End Property

Public Property Let DatePrinted(ByVal vData As Date)
    mvarDatePrinted = vData
End Property

Public Property Get DatePrinted() As Date
    DatePrinted = mvarDatePrinted
End Property

Public Property Let ReturnByDate(ByVal vData As Date)
    mvarReturnByDate = vData
End Property

Public Property Get ReturnByDate() As Date
    ReturnByDate = mvarReturnByDate
End Property

Public Property Let TotalOutstandingTrans(ByVal vData As Integer)
    mvarTotalOutstandingTrans = vData
End Property

Public Property Get TotalOutstandingTrans() As Integer
    TotalOutstandingTrans = mvarTotalOutstandingTrans
End Property

Public Property Let TotalOutstandingDigital(ByVal vData As Integer)
    mvarTotalOutstandingDigital = vData
End Property

Public Property Get TotalOutstandingDigital() As Integer
    TotalOutstandingDigital = mvarTotalOutstandingDigital
End Property

Public Property Let TotalPictures(ByVal vData As Integer)
    mvarTotalPictures = vData
End Property

Public Property Get TotalPictures() As Integer
    TotalPictures = mvarTotalPictures
End Property

Public Property Let ServiceFee(ByVal vData As Currency)
    mvarServiceFee = vData
End Property

Public Property Get ServiceFee() As Currency
    ServiceFee = mvarServiceFee
End Property

Public Property Let WeeklyHoldingFee(ByVal vData As Currency)
    mvarWeeklyHoldingFee = vData
End Property

Public Property Get WeeklyHoldingFee() As Currency
    WeeklyHoldingFee = mvarWeeklyHoldingFee
End Property

Public Property Let LossDamageFee(ByVal vData As Currency)
    mvarLossDamageFee = vData
End Property

Public Property Get LossDamageFee() As Currency
    LossDamageFee = mvarLossDamageFee
End Property

Public Property Let Selected(ByVal vData As Boolean)
    mvarSelected = vData
End Property

Public Property Get Selected() As Boolean
    Selected = mvarSelected
End Property

Public Property Let Memo(ByVal vData As String)
    mvarMemo = vData
End Property

Public Property Get Memo() As String
    Memo = mvarMemo
End Property

Public Property Let DispatchDigital(ByVal vData As String)
    mvarDispatchDigital = vData
End Property

Public Property Get DispatchDigital() As String
    DispatchDigital = mvarDispatchDigital
End Property

Public Property Let DispatchTransparency(ByVal vData As String)
    mvarDispatchTransparency = vData
End Property

Public Property Get DispatchTransparency() As String
    DispatchTransparency = mvarDispatchTransparency
End Property

Public Sub Read(ByVal lngDeliveryNoteNo As Long)

    Dim strSQL As String
    Dim rstDeliveryNote As DAO.Recordset
    
'***************************************
' Module/Form Name   : DeliveryNote
'
' Procedure Name     : Read
'
' Purpose            :
'
' Date Created       : 08/06/2006 00:18
'
' Author             : Gareth
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
    On Error GoTo ErrorProc

    strSQL = "SELECT delivery_note.delivery_note_no, date_printed, returnbydate, customer.customer_name, customer.customer_no, job.job_reference, job_no, Service_Fee, weekly_holding_fee, loss_damage_fee, memo, DispatchDigital, DispatchTransparency, "
    strSQL = strSQL & "(SELECT count(*) FROM search_result, photograph WHERE delivery_note.delivery_note_no = search_result.delivery_note_no and photograph.photograph_key = search_result.photograph_key and photograph.deleted = false) as TotalPictures, "
    strSQL = strSQL & "sum(iif(search_result.photograph_no <> 0 and (search_result.date_returned is null or search_result.date_returned = 0) , 1,0)) as TotalOutstandingTrans "
    strSQL = strSQL & "FROM delivery_note, customer, search_result, job, photograph, Chaser "
    strSQL = strSQL & "WHERE delivery_note.delivery_note_no = " & CStr(lngDeliveryNoteNo) & " AND "
    strSQL = strSQL & "customer.customer_no = delivery_note.customer_no and delivery_note.delivery_note_no = search_result.delivery_note_no and job.delivery_note_no = delivery_note.delivery_note_no and "
    strSQL = strSQL & "photograph.photograph_key = search_result.photograph_key and photograph.deleted = false AND Chaser.DeliveryNoteNo = Delivery_Note.Delivery_Note_No and Chaser.Action = 'None' and Chaser.NextAction = 'None' "
    strSQL = strSQL & "GROUP BY delivery_note.delivery_note_no , customer_name, customer.customer_no, job_no, job.job_reference, date_printed, ReturnByDate, Service_Fee, weekly_holding_fee, loss_damage_fee, memo, DispatchDigital, DispatchTransparency"
    
    Set rstDeliveryNote = db.OpenRecordset(strSQL, dbOpenSnapshot)
    
    With rstDeliveryNote
        mvarDNoteNo = !delivery_note_no
        mvarCustomerNo = !customer_no
        mvarCustomerName = !customer_name
        mvarJobReference = !job_reference
        mvarJobNo = !job_no
        mvarDatePrinted = !date_printed
        mvarReturnByDate = !ReturnByDate
        mvarTotalOutstandingTrans = !TotalOutstandingTrans
        mvarTotalPictures = !TotalPictures
        mvarServiceFee = !service_fee
        mvarWeeklyHoldingFee = !weekly_holding_fee
        mvarLossDamageFee = !loss_damage_fee
        mvarMemo = !Memo
        mvarDispatchDigital = !DispatchDigital
        mvarDispatchTransparency = !DispatchTransparency
    End With
    
    rstDeliveryNote.Close
    Set rstDeliveryNote = Nothing

    On Error GoTo 0
    Exit Sub
'
'********* Code Ends Here **************
'
ErrorProc:
    ErrorRaise "DeliveryNote.Read"
End Sub

