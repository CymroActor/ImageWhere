Attribute VB_Name = "photograph"
Option Explicit
Private snap_batch              As DAO.Recordset
Private snap_batch_counts       As DAO.Recordset
Private dyna_batch              As DAO.Recordset
Private dyna_photograph         As DAO.Recordset
Private SQL                     As String
Private abort                   As Boolean
Private mstrCriteria()              As String
Public mblnMultipleAddition     As Boolean

Public Sub set_up_date_controls(current_date, cmb_month, cmb_year)
'
'   Set up the date combo boxes.
'
    Dim month_no As Integer
    Dim year_no As Integer
    For year_no = 1985 To 2020
        cmb_year.AddItem LTrim(Str(year_no))
        If LTrim(Str(year_no)) = Format(current_date, "yyyy") Then
            cmb_year.ListIndex = cmb_year.NewIndex
        End If
    Next year_no
'
    For month_no = 1 To 12
        cmb_month.AddItem Format(DateSerial(1, month_no, 1), "mmm")
        If cmb_month.List(cmb_month.NewIndex) = Format(current_date, "mmm") Then
            cmb_month.ListIndex = cmb_month.NewIndex
        End If
    Next month_no

End Sub

Private Sub set_and_clauses(where_clause, j, blnKeywordSearch As Boolean)
    Dim i                   As Integer
    Dim strSearchField      As String
    
'***************************************
' Module/Form Name   : photograph
'
' Procedure Name     : set_and_clauses
'
' Purpose            :
'
' Date Created       : 27/07/2005
'
' Author             : Gareth Saunders
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                      27/07/2005 Gareth Saunders
'
'***************************************
'
    On Error GoTo ErrorHandler:
'
'******** Code Starts Here *************
'
    If blnKeywordSearch Then
        strSearchField = "BatchKeywords.Keyword"
    Else
        strSearchField = "Batch.Description"
    End If
    
    where_clause = where_clause & "(" & strSearchField & " like '*" + mstrCriteria(0, j) + "*'"
    For i = 1 To 4
        If mstrCriteria(i, j) = "" Then
            Exit For
        End If
        where_clause = where_clause + " and " & strSearchField & " like '*" + mstrCriteria(i, j) + "*'"
    Next i
    
    where_clause = where_clause + ")"
'
'********* Code Ends Here **************
'
    On Error GoTo 0
    Exit Sub
'
ErrorHandler:
    ErrorRaise "photograph.set_and_clauses"
End Sub


Private Sub validate_criteria(ByVal criteria_no As Integer, ByVal criteria_text As String, abort As Boolean)
    Dim finish As Integer
    Dim no_of_criteria_elements
    
'***************************************
' Module/Form Name   : photograph
'
' Procedure Name     : validate_criteria
'
' Purpose            :
'
' Date Created       : 27/07/2005
'
' Author             : Gareth Saunders
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                      27/07/2005 Gareth Saunders
'
'***************************************
'
    On Error GoTo ErrorHandler:
'
'******** Code Starts Here *************
'
    abort = True
    
    finish = InStr(2, criteria_text, " AND ")
    no_of_criteria_elements = 1
    Do While finish <> 0
        If no_of_criteria_elements > 4 Then
            MsgBox "There cannot be greater than 5 'ADD' clauses within anyone criteria", vbExclamation
            Exit Sub
        End If
  
        If Not finish = 0 Then
            mstrCriteria(no_of_criteria_elements - 1, criteria_no) = _
                        Trim(Mid(criteria_text, 1, finish - 1))
            criteria_text = Trim(Mid(criteria_text, finish + 5))
        End If
  
        no_of_criteria_elements = no_of_criteria_elements + 1
        finish = InStr(2, criteria_text, " AND ")
    Loop
    
    mstrCriteria(no_of_criteria_elements - 1, criteria_no) = criteria_text
  
    abort = False
'
'********* Code Ends Here **************
'
    On Error GoTo 0
    Exit Sub
'
ErrorHandler:
    ErrorRaise "photograph.validate_criteria"
End Sub


Private Sub validate_where_clause(where_clause_string As String, abort As Boolean)
  
    On Error GoTo error_proc
    
    abort = True
    
    Dim no_of_criteria As Integer
    Dim start As Integer
    Dim finish As Integer
'
'   Remove leading and trailing spaces.
'
    where_clause_string = LTrim(RTrim(where_clause_string))
    If Len(where_clause_string) = 0 Then
        Exit Sub
    End If
'
    no_of_criteria = 0
    start = 1
    If Mid(where_clause_string, start, 1) = "," Then
        MsgBox "Invalid search criteria", vbExclamation
        Exit Sub
    End If
    
    If Mid(where_clause_string, Len(where_clause_string), 1) = "," Then
        MsgBox "Invalid search criteria", vbExclamation
        Exit Sub
    End If
    
    Do
        If no_of_criteria > 29 Then
            MsgBox "There cannot be greater than 30 sets of search criteria", vbExclamation, "Print Photograph Labels"
            Exit Sub
        End If
  
        finish = InStr(start + 1, where_clause_string, ",")
  
        If Not finish = 0 Then
            no_of_criteria = no_of_criteria + 1
            ReDim Preserve mstrCriteria(4, 1 To no_of_criteria) ' Resize to an additional element.
            mstrCriteria(0, no_of_criteria) = Mid(where_clause_string, start, finish - start)
            start = finish + 1
            If Mid(where_clause_string, start, 1) = "," Then
                MsgBox "Invalid search criteria", vbExclamation
                Exit Sub
            End If
        End If
  
    Loop While finish <> 0
    
    If Len(where_clause_string) <> 0 Then
'
'   Store last criteria.
'
            If Mid(where_clause_string, start, 1) = "," Then
                MsgBox "Invalid search criteria", vbExclamation
                Exit Sub
            End If
            no_of_criteria = no_of_criteria + 1
            ReDim Preserve mstrCriteria(4, 1 To no_of_criteria) ' Resize to an additional element.
            mstrCriteria(0, no_of_criteria) = Mid(where_clause_string, start)
    End If
    
    Dim i
    For i = 1 To no_of_criteria
         validate_criteria i, mstrCriteria(0, i), abort
         If abort Then
             Exit Sub
         End If
    Next i
    
    abort = False
    
    Exit Sub

error_proc:

    MsgBox Str(Err.Number) + " - " + Err.Description, vbCritical
 
End Sub

Public Function read_counts(batch_no) As DAO.Recordset
    
    SQL = "SELECT (SELECT count(*) FROM photograph " _
        & "WHERE photograph.batch_no = batch.batch_no) as total_no_of_photos, " _
        & "(SELECT count(*) FROM photograph " _
        & "WHERE photograph.batch_no = batch.batch_no " _
        & "and photograph.status = 'I') as no_of_photos_in_stock, " _
        & "(SELECT max(photograph_no) FROM photograph " _
        & "WHERE photograph.batch_no = batch.batch_no) as max_image_no " _
        & "FROM batch " _
        & "WHERE batch_no = " + Str(batch_no) + ";"
    
    Set read_counts = db.OpenRecordset(SQL, dbOpenSnapshot)

End Function

Public Function count_images(batch_no As Long) As Integer
    Dim SQL As String
    Dim count_images_rset As DAO.Recordset
    
    SQL = "SELECT count(*) as no_of_images " _
        & "FROM photograph " _
        & "WHERE photograph.batch_no = " + Str(batch_no) _
        & " and deleted = false;"
    
    Set count_images_rset = db.OpenRecordset(SQL, dbOpenSnapshot)
    count_images = count_images_rset!no_of_images

End Function


Public Function get_last_number_used() As Long
    Dim SQL As String
    Dim get_photograph_no As DAO.Recordset
    
    On Error GoTo ErrorProc
    
    SQL = "SELECT * " _
        & "FROM batchno;"

    Set get_photograph_no = db.OpenRecordset(SQL, dbOpenSnapshot)

    get_last_number_used = get_photograph_no!next_no - 1

    Exit Function
ErrorProc:
    Err.Raise Err.Number
End Function

Private Function where_clause(txt_description As String, _
                              blnKeywordSearch As Boolean, _
                              abort As Boolean) As String
'***************************************
' Module/Form Name   : photograph
'
' Procedure Name     : where_clause
'
' Purpose            :
'
' Date Created       : 29/06/2004
'
' Author             : GARETH SAUNDERS
'
' Parameters         : txt_description - String
'                    : blnKeywordSearch - Boolean
'                    : abort - Boolean
'
' Returns            : String
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo where_clause_Error
'
'******** Code Starts Here *************
'
  
    Dim i, j As Integer
    abort = True
    where_clause = ""
'
'   Validate the string initially
'
    validate_where_clause Duplicate_character(txt_description, "'"), abort
    If abort Then
        Exit Function
    Else
        abort = True
    End If
    
    set_and_clauses where_clause, 1, blnKeywordSearch
    For j = 2 To UBound(mstrCriteria, 2)
        where_clause = where_clause + "OR"
        set_and_clauses where_clause, j, blnKeywordSearch
    Next j
    
    abort = False
    
    Exit Function
'
'********* Code Ends Here **************
'
   Exit Function
'
where_clause_Error:
    ErrorRaise "photograph.where_clause"
End Function


Private Sub delete_images(batch_no, abort)
    
    msg_title = "Delete Images"
    
    abort = True

    On Error GoTo error_proc
    
    begin_trans
    
        SQL = "SELECT * " _
            & "FROM photograph " _
            & "WHERE batch_no = " + Str(batch_no) + ";"
    
        Set dyna_photograph = db.OpenRecordset(SQL, dbOpenDynaset)

        If Not dyna_photograph.RecordCount = 0 Then
            dyna_photograph.MoveFirst
            While Not dyna_photograph.EOF
                dyna_photograph.Delete
                dyna_photograph.MoveFirst
            Wend
        End If
        abort = False
        dyna_photograph.Close
    
    commit_trans
    
    abort = False
    Exit Sub

error_proc:

    Select Case Err.Number
        Case Is = 3200
            MsgBox "Photograph has images with active histories and therefore cannot be deleted", vbExclamation, msg_title
        Case Else
            ErrorRaise "photograph.delete_images"
''            MsgBox Str(Err.Number) + " - " + Err.Description, vbCritical, msg_title
    End Select
    
    On Error Resume Next
    dyna_photograph.Close
    roll_back
End Sub

Public Sub Delete(batch_no As Long, abort As Boolean)
    
    msg_title = "Delete Photograph"
    
    abort = True

    On Error GoTo error_proc
    
    begin_trans
    
    SQL = "SELECT * " _
        & "FROM batch " _
        & "WHERE batch_no = " + Str(batch_no) + ";"
    
    Set dyna_batch = db.OpenRecordset(SQL, dbOpenDynaset)

    If dyna_batch.RecordCount = 0 Then
        MsgBox "photograph not found", vbExclamation, msg_title
        abort = False
        dyna_batch.Close
        roll_back
        Exit Sub
    End If
    
    With dyna_batch
        .Edit
        msg = "Are you certain you wish to Delete Photograph '" + Str(batch_no) + "'"
        style = vbYesNo + vbQuestion + vbDefaultButton2
  
        response = MsgBox(msg, style, msg_title)

        If response = vbNo Then
            dyna_batch.Close
            roll_back
            Exit Sub
        End If
        delete_images batch_no, abort
        If abort Then
            dyna_batch.Close
            roll_back
            Exit Sub
        End If
        .Delete
    End With

    MsgBox "Photograph '" + Str(batch_no) + "'" + " has been deleted.", vbInformation, msg_title
    
    dyna_batch.Close
    
    commit_trans
    
    abort = False
    Exit Sub

error_proc:

    Select Case Err.Number
  
        Case Is = 3200
            MsgBox "Photograph represents a Batch and therefore cannot be deleted", vbExclamation, msg_title
        Case Else
            ErrorRaise "photograph.Delete"
''            MsgBox Str(Err.Number) + " - " + Err.Description, vbCritical, "Delete Job"
    End Select
    On Error Resume Next
    dyna_batch.Close
    roll_back
End Sub

Public Function Read(ByVal plngBatchNo As Long, _
                     ByVal pintRsetType As Integer, _
                     Optional ByVal plngPhotographNo As Variant) As DAO.Recordset
'***************************************
' Module/Form Name   : photograph
'
' Procedure Name     : Read
'
' Purpose            :
'
' Date Created       : 24/06/2004
'
' Author             : GARETH SAUNDERS
'
' Parameters         : plngBatchNo - Long
'                    : pintRsetType - Integer
'                    : plngPhotographNo - Variant
'
' Returns            : DAO.Recordset
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
    SQL = "SELECT * "
    SQL = SQL & "FROM batch "
    SQL = SQL & "WHERE Batch_No = " & CStr(plngBatchNo)
    If Not IsMissing(plngPhotographNo) Then
        SQL = SQL & " AND Phtograp_No = " & CStr(plngPhotographNo)
    End If
    
    Set Read = db.OpenRecordset(SQL, pintRsetType)
'
'********* Code Ends Here **************
'
   Exit Function
'
Read_Error:
    ErrorRaise "photograph.Read"
End Function

Public Sub cancel_return_photo(photograph_key, abort)
'***************************************
' Module/Form Name   : photograph
'
' Procedure Name     : cancel_return_photo
'
' Purpose            :
'
' Date Created       : 29/06/2004
'
' Author             : GARETH SAUNDERS
'
' Parameters         : photograph_key -
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo cancel_return_photo_Error
'
'******** Code Starts Here *************
'
    
    abort = True
'
'   Retrieve photograph
'
    SQL = "SELECT * " _
        & "FROM photograph " _
        & "WHERE photograph_key = " + Str(photograph_key) + ";"

    Set dyna_photograph = db.OpenRecordset(SQL, dbOpenDynaset)
'
'   Update status back to 'O'.
'
    With dyna_photograph
        .Edit
        If !Status = "P" Then
            MsgBox "Image: " + Str(!batch_no) + "/" + Str(!photograph_no) + _
                   " has been used on another Job Sheet", vbExclamation
            dyna_photograph.Close
            Exit Sub
        Else
            If !Status = "O" Then
                MsgBox "Image: " + Str(!batch_no) + "/" + Str(!photograph_no) + _
                       " has been delivered to another customer", vbExclamation
                dyna_photograph.Close
                Exit Sub
            End If
        End If
        !Status = "O"
        .update
    End With
'
'   Batch numbers to be updated.
'
    batch.cancel_return_photo dyna_photograph!batch_no, abort
    If abort Then
        dyna_photograph.Close
        Exit Sub
    End If
    
    dyna_photograph.Close
 
    abort = False
    
'
'********* Code Ends Here **************
'
   Exit Sub
'
cancel_return_photo_Error:
    ErrorRaise "photograph.cancel_return_photo"
End Sub


Public Function photos_in_stock(batch_no As Long) As Boolean
'***************************************
' Module/Form Name   : photograph
'
' Procedure Name     : photos_in_stock
'
' Purpose            :
'
' Date Created       : 29/06/2004
'
' Author             : GARETH SAUNDERS
'
' Parameters         : batch_no - Long
'
' Returns            : Boolean
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo photos_in_stock_Error
'
'******** Code Starts Here *************
'
    
    Dim snap_photograph As DAO.Recordset
    
    photos_in_stock = False
    '
    '   Don't include deleted or digital (Photograph Number is zero) images.
    '
    SQL = "SELECT TOP 1 * "
    SQL = SQL & "From photograph "
    SQL = SQL & "Where photograph.batch_no = " & CStr(batch_no) & " "
    SQL = SQL & "And photograph.status = 'I' "
    SQL = SQL & "And photograph.deleted = false "
    SQL = SQL & "And photograph.photograph_no <> 0"
    
    Set snap_photograph = db.OpenRecordset(SQL, dbOpenDynaset)
    
    If snap_photograph.RecordCount = 0 Then
        snap_photograph.Close
        Exit Function
    End If
    
    photos_in_stock = True
'
'********* Code Ends Here **************
'
   Exit Function
'
photos_in_stock_Error:
    ErrorRaise "photograph.photos_in_stock"
End Function

Private Sub close_photograph_recordsets()
    
    On Error GoTo error_proc
    
    dyna_batch.Close
    snap_batch_counts.Close
    dyna_photograph.Close
    
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

Public Sub make_pending(ByRef poSearchResult As SearchResult, _
                        ByRef pblnCancel As Boolean, _
                        ByRef plngPhotographSwapNo As Long)
'***************************************
' Module/Form Name   : photograph
'
' Procedure Name     : make_pending
'
' Purpose            :
'
' Date Created       : 24/06/2004
'
' Author             : GARETH SAUNDERS
'
' Parameters         : poSearchresult - SearchResult
'                    : pblnCancel - Boolean
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo make_pending_Error
'
'******** Code Starts Here *************
'
    Dim oSearch                 As Search2
    Dim rstSearchResult         As DAO.Recordset
    
    plngPhotographSwapNo = 0
    pblnCancel = True
    '
    '   If this is a Digital Image then process it in a separate routine.
    '
    If poSearchResult.PhotoType = "D" Then
        If Not MakeDigitalPending(poSearchResult) Then Exit Sub
        GoTo photo_end
    End If
    '
    '   Photograph: if an image number has not been entered then get the first
    '   in stock else attempt to get the precise image.
    '
    If poSearchResult.PhotographNo <> 0 Then
        GoTo precise_image
    End If
    
process_first_in_stock:
    
    Set dyna_photograph = photo_image.get_next_available_image(poSearchResult.BatchNo, _
                                                               dbOpenDynaset, _
                                                               False, _
                                                               abort)
    
    If abort Then
        close_photograph_recordsets
        Exit Sub
    End If
    
    With dyna_photograph
        .Edit
        If Not dyna_photograph!Status = "I" Then
            MsgBox "No photographs found in stock", vbExclamation
            dyna_photograph.Close
            Exit Sub
        End If
        !Status = "P"
        .update
    End With
'
'   Update the photograph no to pass back to the calling procedure.
'
    poSearchResult.PhotographNo = dyna_photograph!photograph_no

    GoTo photo_end
    
precise_image:

    Set dyna_photograph = photo_image.Read(poSearchResult.BatchNo, poSearchResult.PhotographNo, 0, False, dbOpenDynaset)
    With dyna_photograph
        If .RecordCount = 0 Then
            MsgBox "Image number does not exist for Photograph", vbExclamation
            dyna_photograph.Close
            Exit Sub
        End If
        .Edit
        Select Case !Status
            Case Is = "I"
                .Edit
                !Status = "P"
                .update
            Case Is = "P"
                Set rstSearchResult = search_result.read_outstanding(dyna_photograph, abort)
''                If search_result.read_outstanding(dyna_photograph, abort)!Confirmed = True Then
                If Not rstSearchResult.EOF Or abort Then
                    If rstSearchResult!Confirmed = True Then
                        Set oSearch = New Search2
                        oSearch.Read search_result.read_outstanding(dyna_photograph, abort)!search_no
                        MsgBox "Image is already pending and confirmed on job: '" & _
                                search_result.read_outstanding(dyna_photograph, abort)!job_description + _
                                "'; search: '" & _
                                oSearch.DisplayDescription & _
                                "'", vbExclamation
                        dyna_photograph.Close
                        Set oSearch = Nothing
                        Exit Sub
                    Else
                        photo_image.take_image_number dyna_photograph, _
                                                      abort, _
                                                      plngPhotographSwapNo
                        If abort Then
                            dyna_photograph.Close
                            Exit Sub
                        End If
                        .Edit
                        !Status = "P"
                        .update
                    End If
                End If
            Case Is = "O"
                If !scanned Then
                    MsgBox "Image is currently out", vbExclamation
                    dyna_photograph.Close
                    Exit Sub
                Else
'                    msg = "Image '" + LTrim(batch_no) + "." + LTrim(Str(photograph_no)) + _
                          "' is at status '" + dyna_photograph!status + "'. Do you wish to swap its image number " + _
                          "with the next available?"
'                    style = vbYesNo + vbQuestion + vbDefaultButton2
'                    response = MsgBox(msg, style, msg_title)
'                    If response = vbNo Then
'                        dyna_photograph.Close
'                        Exit Sub
'                    Else
                        photo_image.take_image_number dyna_photograph, _
                                                      abort, _
                                                      plngPhotographSwapNo
                        If abort Then
                            dyna_photograph.Close
                            Exit Sub
                        End If
                        .Edit
                        !Status = "P"
                        .update
'                    End If
                End If
        End Select
    End With

photo_end:
'
'   Update the photograph key to pass back to the calling procedure.
'
    If poSearchResult.PhotoType = "T" Then
        poSearchResult.PhotographKey = dyna_photograph!photograph_key
        dyna_photograph.Close
    End If
    '
    pblnCancel = False
'
'********* Code Ends Here **************
'
   Exit Sub
'
make_pending_Error:
    ErrorRaise "photograph.make_pending"
End Sub

Public Sub get_all_photographs_in_stock(batch_no As Long)
'***************************************
' Module/Form Name   : photograph
'
' Procedure Name     : get_all_photographs_in_stock
'
' Purpose            :
'
' Date Created       : 29/06/2004
'
' Author             : GARETH SAUNDERS
'
' Parameters         : batch_no - Long
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo get_all_photographs_in_stock_Error
'
'******** Code Starts Here *************
'
    
    SQL = "SELECT * " _
        & "FROM photograph " _
        & "WHERE photograph.batch_no = " + Str(batch_no) _
        & " And photograph.status = 'I' " _
        & "ORDER by photograph_no desc;"
    
    Set dyna_photograph = db.OpenRecordset(SQL, dbOpenDynaset)
    
'
'********* Code Ends Here **************
'
   Exit Sub
'
get_all_photographs_in_stock_Error:
    ErrorRaise "photograph.get_all_photographs_in_stock"
End Sub

Private Function MakeDigitalPending(ByRef poSearchResult As SearchResult) As Boolean
'***************************************
' Module/Form Name   : photograph
'
' Procedure Name     : MakeDigitalPending
'
' Purpose            :
'
' Date Created       : 24/06/2004
'
' Author             : GARETH SAUNDERS
'
' Parameters         : poSearchResult - SearchResult
'
' Returns            : Boolean
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo MakeDigitalPending_Error
'
'******** Code Starts Here *************
'
    Dim oPhotograph     As Photograph2
    Dim oDigitalImage   As DigitalImage
    Dim strMessage      As String
    '
    MakeDigitalPending = False
    '
    '   If a digital Photograph record doesn't exist then one must be created.
    '
    Set oPhotograph = New Photograph2
    oPhotograph.Read poSearchResult.BatchNo, 0
    If oPhotograph.BatchNo = 0 Then
        With oPhotograph
            .BatchNo = poSearchResult.BatchNo
            .PhotographNo = 0
            .Status = "I"
            .Create
        End With
    End If
    '
    Set oDigitalImage = New DigitalImage
    With oDigitalImage
        .BatchNo = poSearchResult.BatchNo
        If oDigitalImage.FileLocation("A4") = "" Then
            strMessage = "There is no 'A4' image of photograph '" & poSearchResult.BatchNo & "'."
            If mblnMultipleAddition Then
                strMessage = strMessage & vbCrLf & "This photograph will not therefore be included in any Searches."
            End If
''            MsgBox strMessage, vbExclamation
            Err.Raise vbObjectError + 1, , strMessage
            Exit Function
        End If
    End With
    Set oDigitalImage = Nothing
    '
    '   Finally set up the Photograph Key.
    '
    poSearchResult.PhotographKey = oPhotograph.PhotographKey
    MakeDigitalPending = True
'
'********* Code Ends Here **************
'
   Exit Function
'
MakeDigitalPending_Error:
    ErrorRaise "photograph.MakeDigitalPending"
End Function

Public Sub amend_no_of_photographs(snap_batch As DAO.Recordset, new_total As Integer, abort As Boolean)
'***************************************
' Module/Form Name   : photograph
'
' Procedure Name     : amend_no_of_photographs
'
' Purpose            :
'
' Date Created       : 29/06/2004
'
' Author             : GARETH SAUNDERS
'
' Parameters         : snap_batch - DAO.Recordset
'                    : new_total - Integer
'                    : abort - Boolean
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo amend_no_of_photographs_Error
'
'******** Code Starts Here *************
'
    
    Dim i As Integer
    Dim photograph_no As Integer
    
    abort = True
    
    get_all_photographs_in_stock snap_batch!batch_no
    
    Set snap_batch_counts = read_counts(snap_batch!batch_no)
    
    Select Case snap_batch_counts!total_no_of_photos
        Case Is > new_total
            With dyna_photograph
                For i = 1 To snap_batch_counts!total_no_of_photos - new_total
                    .MoveFirst
                    .Edit
                    .Delete
                Next i
            End With
    
        Case Is < new_total
            With dyna_photograph
                photograph_no = snap_batch_counts!total_no_of_photos
                For i = 1 To new_total - snap_batch_counts!total_no_of_photos
                    photograph_no = photograph_no + 1
                    .AddNew
                    !batch_no = snap_batch!batch_no
                    !photograph_no = photograph_no
                    !Status = "I"
                    .update
                Next i
            End With
    
        Case Else
            Exit Sub
    End Select
    
    close_photograph_recordsets
    
    abort = False
'
'********* Code Ends Here **************
'
   Exit Sub
'
amend_no_of_photographs_Error:
    ErrorRaise "photograph.amend_no_of_photographs"
End Sub


Public Sub return_photo(photograph_key, abort)
'***************************************
' Module/Form Name   : photograph
'
' Procedure Name     : return_photo
'
' Purpose            :
'
' Date Created       : 29/06/2004
'
' Author             : GARETH SAUNDERS
'
' Parameters         : photograph_key -
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo return_photo_Error
'
'******** Code Starts Here *************
'
    
    abort = True
'
'   Retrieve photograph
'
    SQL = "SELECT * " _
        & "FROM photograph " _
        & "WHERE photograph_key = " + Str(photograph_key) + ";"

    Set dyna_photograph = db.OpenRecordset(SQL, dbOpenDynaset)
'
'   Update status to 'I'
'
    With dyna_photograph
        .Edit
        If Not !Status = "O" Then
            MsgBox "Data error: status of Photograph " + Str(!batch_no) + "/" + _
                    Str(!photograph_no) + " is " + !Status, vbCritical
            dyna_photograph.Close
            Exit Sub
        End If
        !Status = "I"
        .update
    End With
'
'   Batch numbers to be updated.
'
    dyna_photograph.Close
 
    abort = False
'
'********* Code Ends Here **************
'
   Exit Sub
'
return_photo_Error:
    ErrorRaise "photograph.return_photo"
End Sub

Public Sub search(search_string As String, _
                  photograph_recordset As DAO.Recordset, _
                  recordset_type As String, _
                  abort As Boolean, _
                  Optional blnKeywordSearch As Boolean = False)
'***************************************
' Module/Form Name   : photograph
'
' Procedure Name     : search
'
' Purpose            :
'
' Date Created       : 29/06/2004
'
' Author             : GARETH SAUNDERS
'
' Parameters         : search_string - String
'                    : photograph_recordset - DAO.Recordset
'                    : recordset_type - String
'                    : abort - Boolean
'                    : blnKeywordSearch - Boolean
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo search_Error
'
'******** Code Starts Here *************
'
    Dim SQL As String
'
'   Initialise the criteria array.
'
    Erase mstrCriteria
'''
'''   This has been put in because NT takes a great deal of time reading the first
'''   time through.
'''
''    SQL = "SELECT  batch_no, description, Credit " _
''    & "From batch " _
''    & "WHERE description = 'XXXXXXXXXX';"
''
''    Screen.MousePointer = vbHourglass
''
''    Set photograph_recordset = db.OpenRecordset(SQL, recordset_type)
''    photograph_recordset.Close
    
    If blnKeywordSearch Then
        SQL = CreateKeywordSQL(UCase(search_string), abort)
    Else
        SQL = "SELECT  batch_no, description, Credit, Exclusive "
        SQL = SQL & "FROM batch "
        SQL = SQL & "WHERE " & where_clause(UCase(search_string), blnKeywordSearch, abort) & " "
        SQL = SQL & "ORDER BY batch_no"
    End If
    
    If abort Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    DoEvents
    Set photograph_recordset = db.OpenRecordset(SQL, recordset_type)
    Screen.MousePointer = vbDefault
'
'********* Code Ends Here **************
'
   Exit Sub
'
search_Error:
    ErrorRaise "photograph.search (SQL = '" & SQL & "')"
End Sub

Private Function CreateKeywordSQL(ByVal strSearchString As String, abort As Boolean) As String
    Dim strSQL      As String
    Dim intIndex    As Integer
    '
    '   Construct the Criteria array.
    '
    validate_where_clause Replace(strSearchString, "'", "''"), abort
    If abort Then
        Exit Function
    Else
        abort = True
    End If
    '
    '   First dimension is the AND words and the second is the OR words.
    '
    strSQL = ""
    For intIndex = 1 To UBound(mstrCriteria, 2)
        strSQL = strSQL & IIf(intIndex > 1, " UNION ", "") & CreateKeywordSQL_OR(intIndex)
    Next intIndex

    CreateKeywordSQL = strSQL
    
    abort = False
End Function

Private Function CreateKeywordSQL_OR(ByVal intIndex As Integer)
    Dim strSQL              As String
    Dim intNoAddKeywords    As Integer
    Dim intANDIndex         As Integer
        
    For intNoAddKeywords = 1 To 5
        If Trim(mstrCriteria(intNoAddKeywords - 1, intIndex)) = "" Then
            Exit For
        End If
    Next intNoAddKeywords
    
    intNoAddKeywords = intNoAddKeywords - 1
    
    strSQL = "SELECT B0.BATCH_NO, B0.DESCRIPTION, B0.CREDIT, B0.EXCLUSIVE FROM "
    strSQL = strSQL & String(IIf(intNoAddKeywords > 2, intNoAddKeywords - 2, 0), "(")
    '
    '   Construct the AND SQL.
    '
    For intANDIndex = 0 To intNoAddKeywords - 1
        strSQL = strSQL & "(SELECT DISTINCT Batch.* FROM Batch INNER JOIN BatchKeywords ON Batch.batch_no = BatchKeywords.BatchNo WHERE BatchKeywords.Keyword "
        If goSystemConfig.FuzzyKeywordSearch Then
            strSQL = strSQL & "LIKE '*"
        Else
            strSQL = strSQL & "= '"
        End If
        strSQL = strSQL & mstrCriteria(intANDIndex, intIndex)
        If goSystemConfig.FuzzyKeywordSearch Then
            strSQL = strSQL & "*' "
        Else
            strSQL = strSQL & "' "
        End If
        strSQL = strSQL & "ORDER BY Batch.batch_no) AS b" & CStr(intANDIndex)
        If intANDIndex <> 0 Then
            strSQL = strSQL & " ON b0.batch_no = b" & CStr(intANDIndex) & ".batch_no" & IIf(intNoAddKeywords > 1 And intANDIndex <> intNoAddKeywords - 1, ")", "")
        End If
        strSQL = strSQL & IIf(intANDIndex = intNoAddKeywords - 1, "", " INNER JOIN ")
    Next intANDIndex
    
    CreateKeywordSQL_OR = strSQL
    
End Function

Private Function CreateKeywordSQL_AND(ByVal intIndex As Integer)

End Function
