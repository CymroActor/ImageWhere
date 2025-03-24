Attribute VB_Name = "delivery_note"
Option Explicit
''Private dyna_search_result As ADODB.Recordset
Private dyna_search_result      As DAO.Recordset
Private mod_delivery_note_no    As Long

Private SQL As String

Public Function Read(delivery_note_no) As DAO.Recordset
'***************************************
' Module/Form Name   : delivery_note
'
' Procedure Name     : Read
'
' Purpose            :
'
' Date Created       : 28/06/2004
'
' Author             : ADMINISTRATOR
'
' Parameters         : delivery_note_no - DAO.Recordset
'
' Returns            :
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo Read_Error
'
'******** Code Starts Here *************
'
    
    SQL = "SELECT top 1 Chaser.ReturnByDate as return_date, delivery_note.* " _
        & "FROM delivery_note INNER JOIN Chaser ON delivery_note.delivery_note_no = Chaser.DeliveryNoteNo " _
        & "Where delivery_note.delivery_note_no = " & CStr(delivery_note_no) & " " _
        & "ORDER BY Chaser.Id DESC;"

    Set Read = db.OpenRecordset(SQL, dbOpenSnapshot)
'
'********* Code Ends Here **************
'
   Exit Function
'
Read_Error:
    ErrorRaise "delivery_note.Read"
End Function

Private Sub get_job_dyna_search_result(job_no)
'***************************************
' Module/Form Name   : delivery_note
'
' Procedure Name     : get_job_dyna_search_result
'
' Purpose            :
'
' Date Created       : 24/11/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : job_no -
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 24/11/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo get_job_dyna_search_result_Error
'
'******** Code Starts Here *************
'
    Set dyna_search_result = Nothing
''    Set dyna_search_result = New ADODB.Recordset
    
    SQL = "SELECT  search_result.*,photograph.status, photograph.scanned as photo_scanned, search.phototype " _
    & "FROM job, search, search_result, photograph, batch, " _
    & "job INNER JOIN search ON job.job_no = search.job_no, " _
    & "search INNER JOIN search_result ON search.search_no = search_result.search_no, " _
    & "search_result INNER JOIN photograph ON search_result.photograph_key = photograph.photograph_key, " _
    & "photograph INNER JOIN batch ON photograph.batch_no = batch.batch_no " _
    & "WHERE job.job_no = " + Str(job_no) + ";"
    
''    dyna_search_result.Open SQL, gdbADO, adOpenDynamic, adLockPessimistic
    Set dyna_search_result = db.OpenRecordset(SQL, dbOpenDynaset)
    
    Exit Sub
'
'********* Code Ends Here **************
'
    Exit Sub
    '
get_job_dyna_search_result_Error:
    ErrorRaise "delivery_note.get_job_dyna_search_result"
End Sub



Public Sub deliver_job_search_results(in_delivery_note_no, job_no, abort As Boolean)
'***************************************
' Module/Form Name   : delivery_note
'
' Procedure Name     : deliver_job_search_results
'
' Purpose            :
'
' Date Created       : 29/07/2001
'
' Author             : GARETH
'
' Parameters         : in_delivery_note_no - abort
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo deliver_job_search_results_Error
'
'******** Code Starts Here *************
'
    
    abort = True
    
    get_job_dyna_search_result (job_no)
    
    If dyna_search_result.RecordCount = 0 Then
        close_delivery_recordsets
        Exit Sub
    End If
    
    dyna_search_result.MoveFirst
    
    With dyna_search_result
        While Not .EOF
            .Edit
''            If NVL(!PhotoType, NVLString) = "T" Then
            If NVL(!PhotoType, NVLString) = "T" And Not !Status = "P" Then
                MsgBox "Photograph " + Str(!batch_no) + "/" + Str(!photograph_no) + _
                       " is not at Status 'Pending'." + Chr(10) + "Aborting transaction", vbCritical
                close_delivery_recordsets
                Exit Sub
            End If
            If !Confirmed Then
                If NVL(!PhotoType, NVLString) = "T" Then
                    !Status = "O"
                    !photo_scanned = True
                End If
                !delivery_note_no = in_delivery_note_no
            Else
                '
                '   Set date returned and booked in by to show a cancellation.
                '
                !Status = "I"
                !date_returned = Date
                !booked_in_by = "ABORT DNOTE: " + Str(in_delivery_note_no)
            End If
            .update
''            End If
            dyna_search_result.MoveNext
        Wend
    End With
    
    close_delivery_recordsets

    abort = False
'
'********* Code Ends Here **************
'
   Exit Sub
'
deliver_job_search_results_Error:
    ErrorSave
    On Error Resume Next
    close_delivery_recordsets
    On Error GoTo 0
    ErrorRestore "delivery_note.deliver_job_search_results"
End Sub






Private Sub close_delivery_recordsets()
    On Error GoTo error_proc
    
    dyna_search_result.Close
    
    Exit Sub
    
error_proc:
    Select Case Err.Number
        Case Is = 91
            Resume Next
        Case Is = 3420
            Resume Next
        Case Else
            MsgBox Str(Err.Number) + " - " + Err.Description, vbCritical
        End Select
End Sub


Public Sub read_customers_outstanding_delivery_notes(filter_customer, _
                                                     filter_photo, _
                                                     return_SQL As String, _
                                                     Optional ByVal pblnOutstandingOnly As Boolean = True)
    Dim filter_customer_text As String
    Dim filter_photo_text As String
    Dim filter_image_text As String
    Dim filter_delivery_notes As String
    Dim filter_delivery_notes_recordset As DAO.Recordset
    Dim photograph_no As Long
    Dim image_no As Long
'
'   Extract the photograph and image number from the entry.
'
    photograph_no = CLng(characters_before_decimal(filter_photo))
    image_no = CLng(characters_after_decimal(filter_photo))
'
    If filter_customer = 0 Then
        filter_customer_text = ""
    Else
        filter_customer_text = " dnote.customer_no = " & Str(filter_customer) & " AND "
    End If
  
    If photograph_no = 0 Then
        filter_photo_text = ""
    Else
        If image_no = 0 Then
            filter_image_text = ""
        Else
            filter_image_text = " and search_result.photograph_no = " & Str(image_no)
        End If
        SQL = "SELECT delivery_note.delivery_note_no " _
            & "FROM delivery_note, search_result " _
            & "WHERE delivery_note.delivery_note_no = search_result.delivery_note_no " _
            & "and search_result.batch_no = " & Str(photograph_no) _
            & filter_image_text & ";"
        Set filter_delivery_notes_recordset = db.OpenRecordset(SQL, dbOpenSnapshot)
        filter_delivery_notes = ""
        If Not filter_delivery_notes_recordset.EOF Then
            With filter_delivery_notes_recordset
                .MoveFirst
                While Not .EOF
                    filter_delivery_notes = filter_delivery_notes & _
                                            CStr(!delivery_note_no) & ","
                    .MoveNext
                Wend
            End With
        End If
        If filter_delivery_notes = "" Then
            filter_delivery_notes = "0"
        Else
            filter_delivery_notes = Left(filter_delivery_notes, Len(filter_delivery_notes) - 1)
        End If
        filter_photo_text = " dnote.delivery_note_no IN (" & filter_delivery_notes & ") AND "
        '
        filter_delivery_notes_recordset.Close
        Set filter_delivery_notes_recordset = Nothing
    End If
    
'    SQL = "SELECT delivery_note.delivery_note_no, date_printed, returnbydate, " _
        & "customer.customer_name, customer.customer_no, job.job_reference, " _
        & "(SELECT count(*) FROM search_result, photograph WHERE delivery_note.delivery_note_no = search_result.delivery_note_no and photograph.photograph_key = search_result.photograph_key and photograph.deleted = false) as total_pictures, " _
        & "count(search_result.search_result_key) as total_outstanding " _
        & "FROM delivery_note, customer, search_result, job, photograph, Chaser " _
        & "WHERE " + filter_customer_text + " customer.customer_no = delivery_note.customer_no and delivery_note.delivery_note_no = search_result.delivery_note_no and search_result.date_returned is null " _
        & "and job.delivery_note_no = delivery_note.delivery_note_no " + filter_photo_text + " and photograph.photograph_key = search_result.photograph_key and photograph.deleted = false AND " _
        & "Chaser.DeliveryNoteNo = Delivery_Note.Delivery_Note_No and Chaser.Action = 'None' and Chaser.NextAction = 'None' " _
        & "GROUP BY delivery_note.delivery_note_no, customer_name, customer.customer_no, job.job_reference, date_printed, " _
        & "returnbydate " _
        & "ORDER BY delivery_note.delivery_note_no desc;"
  
    SQL = "SELECT dnote.delivery_note_no, dnote.Date_Printed, Chaser.ReturnByDate, customer.customer_name, customer.customer_no, job.job_reference, "
    SQL = SQL & "Count(search_result.search_result_key) As total_pictures, job.job_no, "
    SQL = SQL & "Sum(iif(search_result.photograph_no <> 0 and (search_result.date_returned Is Null or search_result.date_returned =0),1,0)) as  total_outstanding_trans, "
    SQL = SQL & "Sum(iif(search_result.photograph_no = 0 and (search_result.date_returned Is Null or search_result.date_returned =0) and (search_result.digitalstatus is null or search_result.digitalstatus = ''),1,0)) as  total_outstanding_digital, "
    SQL = SQL & "(select top 1 action from chaser where chaser.deliverynoteno = dnote.delivery_note_no and chaser.dateamended = (select max(dateamended) from chaser where chaser.deliverynoteno = dnote.delivery_note_no)) as DNStatus "
    SQL = SQL & "FROM (search_result INNER JOIN photograph ON search_result.photograph_key = photograph.photograph_key) INNER JOIN (job INNER JOIN ((delivery_note as dnote INNER JOIN customer ON dnote.customer_no = customer.customer_no) INNER JOIN Chaser ON dnote.delivery_note_no = Chaser.DeliveryNoteNo) ON job.delivery_note_no = dnote.delivery_note_no) ON search_result.delivery_note_no = dnote.delivery_note_no "
    SQL = SQL & "Where "
    SQL = SQL & filter_customer_text
    SQL = SQL & filter_photo_text
    SQL = SQL & "photograph.deleted=False AND "
    SQL = SQL & "Chaser.Action='None' AND "
    SQL = SQL & "Chaser.NextAction='None' "
    SQL = SQL & "Group By "
    SQL = SQL & "dnote.delivery_note_no, "
    SQL = SQL & "dnote.Date_Printed, "
    SQL = SQL & "Chaser.ReturnByDate, "
    SQL = SQL & "customer.customer_name, "
    SQL = SQL & "customer.customer_no, "
    SQL = SQL & "job.job_reference, "
    SQL = SQL & "job.job_no "
    If pblnOutstandingOnly Then
        SQL = SQL & "Having Sum (iif(search_result.photograph_no <> 0 and (search_result.date_returned Is Null or search_result.date_returned =0),1,0)) <> 0 "
    End If
    SQL = SQL & "ORDER BY dnote.delivery_note_no DESC"

'    Set delivery_note_recordset = db.OpenRecordset(SQL, recordset_type)
    return_SQL = SQL

End Sub


