Attribute VB_Name = "search_result"
Option Explicit
Private SQL As String
Private abort As Boolean
Private dyna_search_result As DAO.Recordset
Public Function read_outstanding(dyna_photograph, abort) As DAO.Recordset
  
    On Error GoTo error_proc
    
    abort = True
    
    Set read_outstanding = Nothing
    SQL = "SELECT * " _
        & "FROM search_result " _
        & "WHERE search_result.batch_no = " + Str(dyna_photograph!batch_no) + _
          " and search_result.photograph_no = " + Str(dyna_photograph!photograph_no) + _
          " and (search_result.date_returned is null or search_result.date_returned = 0)"
  
    Set read_outstanding = db.OpenRecordset(SQL, dbOpenDynaset)
    
    abort = False
    
    Exit Function
    
error_proc:
    MsgBox Str(Err.Number) + " - " + Err.Description, vbCritical
    
    On Error Resume Next
    
    read_outstanding.Close

End Function

Public Function Confirmed(search_result_key) As Boolean
    
    Set dyna_search_result = Read(search_result_key, dbOpenDynaset)
    Confirmed = dyna_search_result!Confirmed
    
End Function


Public Function Read(search_result_key, rsettype) As DAO.Recordset
    
    SQL = "SELECT * " _
        & "FROM search_result " _
        & "WHERE search_result_key = " + Str(search_result_key) + ";"

    Set Read = db.OpenRecordset(SQL, rsettype)

End Function

Public Sub confirm(search_result_key As Long, Confirmed As Boolean)
    
    Set dyna_search_result = Read(search_result_key, dbOpenDynaset)
    With dyna_search_result
        .Edit
        If !Confirmed = True Then
            !Confirmed = False
            Confirmed = False
        Else
            !Confirmed = True
            Confirmed = True
        End If
        .update
    End With
    
    dyna_search_result.Close
    
    Exit Sub

error_proc:
    MsgBox Str(Err.Number) + " - " + Err.Description, vbCritical
    
End Sub

Public Function customers_pending_photographs(customer_no) As Long
    
'***************************************
' Module/Form Name   : search_result
'
' Procedure Name     : customers_pending_photographs
'
' Purpose            :
'
' Date Created       : 20/04/2006 11:12
'
' Author             : Gareth
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                      2006-04-20 Gareth    Ignore digital images and allow for a zero date returned field.
'***************************************
'
    On Error GoTo ErrorProc

    SQL = "SELECT count(*) as customers_pending_photographs " _
        & "FROM search_result " _
        & "WHERE (date_returned is null or date_returned = 0) and customer_no = " + Str(customer_no) _
        & " and delivery_note_no is null and photograph_no > 0;"

    Set dyna_search_result = db.OpenRecordset(SQL, dbOpenDynaset)
    
    customers_pending_photographs = dyna_search_result!customers_pending_photographs

    On Error GoTo 0
    Exit Function
'
'********* Code Ends Here **************
'
ErrorProc:
    ErrorRaise "search_result.customers_pending_photographs"
End Function

Public Sub cancel_return_photo(search_result_key, abort)

    On Error GoTo error_proc
    
    abort = True
'
'   Read search_result
'
    SQL = "SELECT * " _
        & "FROM search_result " _
        & "WHERE search_result.search_result_key = " + Str(search_result_key) + ";"
    
    Set dyna_search_result = db.OpenRecordset(SQL, dbOpenDynaset)
'
'   Cancel the return of the photo.
'
    photograph.cancel_return_photo dyna_search_result!photograph_key, abort
    If abort Then
        dyna_search_result.Close
        Exit Sub
    End If
'
'   Nullify the customer return no.
'
    With dyna_search_result
        .Edit
        !customer_return_no = Null
        .update
    End With
    
    dyna_search_result.Close
    
    abort = False
    
    Exit Sub
    
error_proc:
    MsgBox Str(Err.Number) + " - " + Err.Description, vbCritical
    
    On Error Resume Next
    dyna_search_result.Close

End Sub

Public Function customers_outstanding_photographs(customer_no) As Long
    
'***************************************
' Module/Form Name   : search_result
'
' Procedure Name     : customers_outstanding_photographs
'
' Purpose            :
'
' Date Created       : 20/04/2006 11:05
'
' Author             : Gareth
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                      2006-04-20 Gareth    Ignore digital images and allow for a zero date returned field.
'***************************************
'
    On Error GoTo ErrorProc

    SQL = "SELECT count(*) as customers_outstanding_photographs " _
        & "FROM search_result " _
        & "WHERE (date_returned is null or date_returned = 0) and customer_no = " + Str(customer_no) _
        & " and delivery_note_no <> null and photograph_no > 0;"

    Set dyna_search_result = db.OpenRecordset(SQL, dbOpenDynaset)
    
    customers_outstanding_photographs = dyna_search_result!customers_outstanding_photographs

    On Error GoTo 0
    Exit Function
'
'********* Code Ends Here **************
'
ErrorProc:
    ErrorRaise "search_result.customers_outstanding_photographs"
End Function

Public Function read_all_delivery_note_photos(delivery_note_no, rset_type) As DAO.Recordset
    
    SQL = "SELECT search_result.*, batch.description " _
        & "FROM (search_result INNER JOIN photograph " _
        & "ON search_result.photograph_key = photograph.photograph_key) " _
        & "INNER JOIN batch ON photograph.batch_no = batch.batch_no " _
        & "WHERE search_result.delivery_note_no = " + Str(delivery_note_no) _
        & " AND photograph.deleted=False;"
    
    Set read_all_delivery_note_photos = db.OpenRecordset(SQL, rset_type)

End Function

Public Sub return_photo(search_result_key, _
                        date_returned, _
                        booked_in_by, _
                        abort)

    On Error GoTo error_proc
    
    abort = True
'
'   Read search_result
'
    SQL = "SELECT * " _
        & "FROM search_result " _
        & "WHERE search_result.search_result_key = " + Str(search_result_key) + ";"
    
    Set dyna_search_result = Nothing
    Set dyna_search_result = db.OpenRecordset(SQL, dbOpenDynaset)
'
'   Update return data.
'
    With dyna_search_result
        .Edit
        !date_returned = date_returned
        !booked_in_by = booked_in_by
        .update
        .Bookmark = .LastModified
    End With
'
'   Photograph status updated to 'I'
'
    photograph.return_photo dyna_search_result!photograph_key, abort
    If abort Then
        dyna_search_result.Close
        Exit Sub
    End If
    
    dyna_search_result.Close
    
    abort = False
    
    Exit Sub
    
error_proc:
    MsgBox Str(Err.Number) + " - " + Err.Description, vbCritical
    
    On Error Resume Next
    dyna_search_result.Close

End Sub


