Attribute VB_Name = "photo_image"
Option Explicit
Private SQL As String
Private dyna_photo_image As DAO.Recordset
Private snap_photo_image As DAO.Recordset
Public Sub undelete(batch_no, photograph_no, abort As Boolean, Optional SuppressQuestion As Boolean)
    
    msg_title = "Undelete Image"
    
    abort = True

    On Error GoTo error_proc
    
    begin_trans
    
        Set dyna_photo_image = Read(batch_no, photograph_no, 0, True, dbOpenDynaset)
    
        If dyna_photo_image.RecordCount = 0 Then
            MsgBox "image not found", vbExclamation, msg_title
            abort = False
            close_recordset dyna_photo_image
            roll_back
            Exit Sub
        End If
    
        With dyna_photo_image
            .Edit
            If Not SuppressQuestion Then
                msg = "Are you certain you wish to undelete Image '" + Str(photograph_no) + "'"
                style = vbYesNo + vbQuestion + vbDefaultButton2

                response = MsgBox(msg, style, msg_title)
    
                If response = vbNo Then
                    close_recordset dyna_photo_image
                    roll_back
                    Exit Sub
                End If
            End If
            !Deleted = False
            .update
        End With

        If Not SuppressQuestion Then
            MsgBox "Image '" + Str(photograph_no) + "'" + " has been undeleted.", vbInformation, msg_title
        End If
  
        close_recordset dyna_photo_image
    
    commit_trans
    
    abort = False
    
    Exit Sub

error_proc:

    MsgBox Str(Err.Number) + " - " + Err.Description, vbCritical, msg_title
    
    close_recordset dyna_photo_image
    
    roll_back

End Sub

Public Sub Image_Display(batch_no, SQL, Archived_Images As Boolean)
'***************************************
' Module/Form Name   : photo_image
'
' Procedure Name     : Image_Display
'
' Purpose            :
'
' Date Created       : 16/05/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : batch_no - Archived_Images
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo Image_Display_Error
'
'******** Code Starts Here *************
'
    
    Dim Deleted_String As String
    
    If Archived_Images Then
        Deleted_String = "photograph.deleted = true"
    Else
        Deleted_String = "photograph.deleted = false"
    End If

    SQL = "SELECT "
    SQL = SQL & "search_result.photograph_no as ImageNo, "
    SQL = SQL & "search_result.delivery_note_no as Delivery, "
    SQL = SQL & "photograph.status as Status, "
    SQL = SQL & "customer.customer_name as [Customer Name], "
    SQL = SQL & "customer.address_line_1 as [Address Line], "
    SQL = SQL & "photograph.scanned as [Barcode Read], "
    SQL = SQL & "customer.customer_no, "
    SQL = SQL & "(select search.job_no from search where search_no = search_result.search_no) as job_no, "
    SQL = SQL & "search_result.search_no "
    SQL = SQL & "From "
    SQL = SQL & "(photograph INNER JOIN Search_result ON photograph.photograph_key = Search_result.photograph_key) INNER JOIN Customer ON Search_result.customer_no = Customer.customer_no "
    SQL = SQL & "Where "
    SQL = SQL & "photograph.batch_no = " & CStr(batch_no) & " and "
    SQL = SQL & "photograph.status <> 'I' and "
    SQL = SQL & "(search_result.date_returned is null or search_result.date_returned =0) and "
    SQL = SQL & "photograph.photograph_no <> 0 and "
    SQL = SQL & Deleted_String & " "
    SQL = SQL & "Union "
    SQL = SQL & "SELECT "
    SQL = SQL & "photograph.photograph_no, "
    SQL = SQL & "'' as Delivery, "
    SQL = SQL & "photograph.status, "
    SQL = SQL & "'' as [Customer Name], "
    SQL = SQL & "'' as [Address Line], "
    SQL = SQL & "photograph.scanned as [Barcode Read], "
    SQL = SQL & "0 as customer_no, "
    SQL = SQL & "0 as job_no, "
    SQL = SQL & "0 as search_no "
    SQL = SQL & "From "
    SQL = SQL & "photograph "
    SQL = SQL & "Where "
    SQL = SQL & "photograph.batch_no = " & CStr(batch_no) & " and "
    SQL = SQL & "photograph.status = 'I' and "
    SQL = SQL & "photograph.photograph_no <> 0 and "
    SQL = SQL & Deleted_String & " "
    SQL = SQL & "UNION SELECT "
    SQL = SQL & "search_result.photograph_no as ImageNo, "
    SQL = SQL & "search_result.delivery_note_no as Delivery, "
    SQL = SQL & "Iif(search_result.delivery_note_no is null, 'P', search_result.DigitalStatus) as Status, "
    SQL = SQL & "customer.customer_name as [Customer Name], "
    SQL = SQL & "customer.address_line_1 as [Address Line], "
    SQL = SQL & "false as [Barcode Read], "
    SQL = SQL & "customer.customer_no, "
    SQL = SQL & "(select search.job_no from search where search_no = search_result.search_no) as job_no, "
    SQL = SQL & "search_result.search_no "
    SQL = SQL & "From "
    SQL = SQL & "(photograph INNER JOIN Search_result ON photograph.photograph_key = Search_result.photograph_key) INNER JOIN Customer ON Search_result.customer_no = Customer.customer_no "
    SQL = SQL & "Where "
    SQL = SQL & "photograph.batch_no = " & CStr(batch_no) & " and "
    SQL = SQL & "photograph.photograph_no = 0 and "
    SQL = SQL & Deleted_String & " "
    
''    SQL = "SELECT "
''    SQL = SQL & "search_result.photograph_no as ImageNo, "
''    SQL = SQL & "search_result.delivery_note_no as Delivery, "
''    SQL = SQL & "photograph.status as Status, "
''    SQL = SQL & "customer.customer_name as [Customer Name], "
''    SQL = SQL & "customer.address_line_1 as [Address Line], "
''    SQL = SQL & "photograph.scanned as [Barcode Read], "
''    SQL = SQL & "customer.customer_no, "
''    SQL = SQL & "(select search.job_no from search where search_no = search_result.search_no) as job_no, "
''    SQL = SQL & "search_result.search_no "
''    SQL = SQL & "From "
''    SQL = SQL & "(photograph INNER JOIN Search_result ON photograph.photograph_key = Search_result.photograph_key) INNER JOIN Customer ON Search_result.customer_no = Customer.customer_no "
''    SQL = SQL & "Where "
''    SQL = SQL & "photograph.batch_no = " & CStr(batch_no) & " and "
''    SQL = SQL & "photograph.status <> 'I' and "
''    SQL = SQL & "photograph.deleted = false and "
''    SQL = SQL & "(search_result.date_returned is null or search_result.date_returned =0) "
''    SQL = SQL & Deleted_String & " "
''    SQL = SQL & "Union "
''    SQL = SQL & "SELECT "
''    SQL = SQL & "photograph.photograph_no, "
''    SQL = SQL & "search_result.delivery_note_no, "
''    SQL = SQL & "photograph.status, "
''    SQL = SQL & "customer.customer_name as [Customer Name], "
''    SQL = SQL & "customer.address_line_1 as [Address Line], "
''    SQL = SQL & "photograph.scanned, "
''    SQL = SQL & "customer.customer_no, "
''    SQL = SQL & "(select search.job_no from search where search_no = search_result.search_no) as job_no, "
''    SQL = SQL & "search_result.search_no "
''    SQL = SQL & "From "
''    SQL = SQL & "(photograph INNER JOIN Search_result ON photograph.photograph_key = Search_result.photograph_key) INNER JOIN Customer ON Search_result.customer_no = Customer.customer_no "
''    SQL = SQL & "Where "
''    SQL = SQL & "photograph.batch_no = 10088 and "
''    SQL = SQL & "photograph.photograph_no = 0 "
''    SQL = SQL & Deleted_String & " "
'
'********* Code Ends Here **************
'
   Exit Sub
'
Image_Display_Error:
    ErrorRaise "photo_image.Image_Display"
End Sub

Public Function Read(batch_no, _
                     photograph_no, _
                     photograph_key, _
                     archived_image As Boolean, _
                     rsettype) As DAO.Recordset
'***************************************
' Module/Form Name   : photo_image
'
' Procedure Name     : Read
'
' Purpose            :
'
' Date Created       : 25/03/2005
'
' Author             : GARETH SAUNDERS
'
' Parameters         : batch_no - photograph_key
'                    : archived_image - Boolean
'                    : rsettype - DAO.Recordset
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
    Dim Deleted_String As String
    
    If archived_image Then
        Deleted_String = " and photograph.deleted = true;"
    Else
        Deleted_String = " and photograph.deleted = false;"
    End If
    
    If batch_no = 0 Or photograph_no = 0 Then
        SQL = "SELECT * " _
            & "FROM photograph " _
            & "WHERE photograph_key = " + Str(photograph_key) + Deleted_String
    Else
        SQL = "SELECT * " _
            & "FROM photograph " _
            & "WHERE batch_no = " + Str(batch_no) + " and " _
            & " photograph_no = " + Str(photograph_no) + Deleted_String
    End If
    
    Set Read = db.OpenRecordset(SQL, rsettype)
'
'********* Code Ends Here **************
'
   Exit Function
'
Read_Error:
    ErrorRaise "photo_image.Read"
End Function

Public Sub Delete(batch_no, photograph_no, abort As Boolean, Optional SuppressQuestion As Boolean)
    '***************************************
    ' Module/Form Name   : photo_image
    '
    ' Procedure Name     : delete
    '
    ' Purpose            :
    '
    ' Date Created       : 23/04/2002
    '
    ' Author             : GARETH SAUNDERS
    '
    ' Parameters         : batch_no - abort
    '
    ' Amendment History  : Date       Author    Description
    '                    : --------------------------------
    '                    : 23/04/2002 GARETH SAUNDERS
    '
    '***************************************
    '
    On Error GoTo Delete_Error
    '
    '******** Code Starts Here *************
    '
    Dim OutMessage As String

    msg_title = "Delete Image"

    abort = True

    begin_trans

        Set dyna_photo_image = Nothing
        Set dyna_photo_image = Read(batch_no, photograph_no, 0, False, dbOpenDynaset)
    
        If dyna_photo_image.RecordCount = 0 Then
            MsgBox "image not found", vbExclamation, msg_title
            abort = False
            close_recordset dyna_photo_image
            roll_back
            Exit Sub
        End If
    
        With dyna_photo_image
            .Edit
            Select Case !Status
                Case Is = "O"
                    OutMessage = "Image is currently out on a Delivery Note." & Chr(10)
                Case Is = "P"
                    roll_back
                    Err.Raise vbObjectError + 1, , "Image must be removed from job before deleting."
                Case Else
                    OutMessage = ""
            End Select
            If Not SuppressQuestion Or _
                !Status = "O" Then
                msg = OutMessage & "Are you certain you wish to delete Image '" + Str(photograph_no) + "'"
                style = vbYesNo + vbQuestion + vbDefaultButton2
    
                response = MsgBox(msg, style, msg_title)
    
                If response = vbNo Then
                    close_recordset dyna_photo_image
                    roll_back
                    Exit Sub
                End If
            End If
            On Error Resume Next
            .Delete
            If Err.Number = 3200 Then
                .Edit
                !Deleted = True
                .update
                OutMessage = "archived"
            Else
                If Err.Number <> 0 Then
                    GoTo Delete_Error
                End If
                OutMessage = "deleted"
            End If
            On Error GoTo Delete_Error
        End With
    
        If Not SuppressQuestion Then
            MsgBox "Image '" + Str(photograph_no) + "'" + " has been " & OutMessage & ".", vbInformation, msg_title
        End If
        close_recordset dyna_photo_image

    commit_trans

    abort = False

    '
    '********* Code Ends Here **************
    '
Exit Sub
    '
Delete_Error:
    ErrorRaise "photo_image.delete"
End Sub

Private Sub close_photo_image_recordsets()
    
    On Error GoTo error_proc
    
    dyna_photo_image.Close
    snap_photo_image.Close
'    next_dyna_photo_image.Close
'    out_search_result.Close
    
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

Public Function get_next_available_image(batch_no, rsettype, scanned As Boolean, abort) As DAO.Recordset
'
'   Get next available image that is not scanned! (GS 25/04/96)
'
    On Error GoTo error_proc
  
    Dim scanned_string As String
    
    abort = True
    
    If scanned Then
        scanned_string = " and photograph.scanned = false"
    Else
        scanned_string = ""
    End If
    
    SQL = "SELECT TOP 1 * "
    SQL = SQL & "From photograph "
    SQL = SQL & "Where photograph.batch_no = " & Str(batch_no) & " "
    SQL = SQL & "And photograph.deleted = false "
    SQL = SQL & "And photograph.status = 'I' " & scanned_string & " "
    SQL = SQL & "And photograph.photograph_no <> 0 "
    SQL = SQL & "ORDER BY photograph.photograph_no"
    
    Set get_next_available_image = db.OpenRecordset(SQL, rsettype)
    
    If get_next_available_image.RecordCount = 0 Then
        If scanned Then
            MsgBox "No unscanned images to swap found in stock for Photograph: " + Str(batch_no), vbExclamation
        Else
            MsgBox "No images to swap found in stock for Photograph: " + Str(batch_no), vbExclamation
        End If
        Exit Function
    End If
    
    abort = False
    
    Exit Function
    
error_proc:
    MsgBox Str(Err.Number) + " - " + Err.Description, vbCritical
    
    On Error Resume Next
    
End Function

Public Sub take_image_number(ByRef from_dyna_photo_image, _
                             ByRef abort, _
                             ByRef plngPhotographSwapNo As Long)
'***************************************
' Module/Form Name   : photo_image
'
' Procedure Name     : take_image_number
'
' Purpose            :
'
' Date Created       : 08/07/2004
'
' Author             : GARETH SAUNDERS
'
' Parameters         : from_dyna_photo_image - abort
'                    : pblnImageSwapped - Boolean
'                    : plngPhotographNo - Long
'                    : plngPhotographKey - Long
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo take_image_number_Error
'
'******** Code Starts Here *************
'
    Dim next_dyna_photo_image As DAO.Recordset
    Dim out_search_result As DAO.Recordset
    
    abort = True
    plngPhotographSwapNo = 0
    '
    '   The entered recordset's photograph must not be scanned if it is OUT and not
    '   confirmed if it is PENDING.
    '
    If from_dyna_photo_image!Status = "O" Then
        If from_dyna_photo_image!scanned = True Then
            Exit Sub
        End If
    End If
    '
    '   Get next available image number for recordset's photograph.
    '
    Set next_dyna_photo_image = _
        photo_image.get_next_available_image(from_dyna_photo_image!batch_no, _
                                             dbOpenDynaset, _
                                             IIf(from_dyna_photo_image!Status = "O", True, False), _
                                             abort)
    
    If abort Then
        close_photo_image_recordsets
        Exit Sub
    End If
    '
    '   Get recordset's 'OUT'/'PENDING' search_result.
    '
    Set out_search_result = search_result.read_outstanding(from_dyna_photo_image, abort)
    If abort Then
        close_photo_image_recordsets
        Exit Sub
    End If
    '
    '   The entered recordset's photograph must not be confirmed if it is PENDING.
    '
    If from_dyna_photo_image!Status = "P" Then
        If out_search_result!Confirmed = True Then
            Exit Sub
        End If
    End If
    '
    '   Make the next available image number the one that is out.
    '
    With out_search_result
        .Edit
        !photograph_no = next_dyna_photo_image!photograph_no
        !photograph_key = next_dyna_photo_image!photograph_key
        .update
        On Error Resume Next
        goLog.WriteDetailedLog CStr(!search_no) & "/" & CStr(!batch_no) & "/" & CStr(!photograph_no), "Search Result Updated (Dynaset)"
        On Error GoTo take_image_number_Error
    End With

    With next_dyna_photo_image
        .Edit
        !Status = from_dyna_photo_image!Status  '   The status of the image swapped.
        .update
    End With
    '
    plngPhotographSwapNo = next_dyna_photo_image!photograph_no
    '
    abort = False
    
'
'********* Code Ends Here **************
'
   Exit Sub
'
take_image_number_Error:
    DisplayError , "photo_image.take_image_number", vbExclamation
    'To pass the error back to the calling proc., use the following line
'    ErrorRaise "photo_image.take_image_number"
End Sub

Public Sub change_image_number(ByVal from_search_result_key As Long, _
                               ByVal to_photograph_image_no As String, _
                               ByVal search_no As Long, _
                               ByRef poSearchResult As SearchResult, _
                               ByRef abort As Boolean)
'***************************************
' Module/Form Name   : photo_image
'
' Procedure Name     : change_image_number
'
' Purpose            :
'
' Date Created       : 21/07/2004
'
' Author             : GARETH SAUNDERS
'
' Parameters         : from_search_result_key - Long
'                    : to_photograph_image_no - Long
'                    : search_no - Long
'                    : poSearchResult - SearchResult
'                    : abort - Boolean
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo change_image_number_Error
'
'******** Code Starts Here *************
'
    
    Dim dummy_key               As Long
    Dim Cancel                  As Boolean
    Dim oSearchResult           As New SearchResult
    Dim lngPhotographSwapNo     As Long
    Dim from_search_result As DAO.Recordset
    Dim from_snap_photo_image As DAO.Recordset
    Dim to_snap_photo_image As DAO.Recordset
    
    abort = True
'
'   Get from search_result.
'
    SQL = "SELECT * " _
        & "FROM search_result " _
        & "WHERE search_result_key = " + Str(from_search_result_key) + ";"
    
    Set from_search_result = db.OpenRecordset(SQL, dbOpenDynaset)
'
'   Get photo images.
'
    Dim to_batch_no As Long
    Dim to_photograph_no As Long
    to_batch_no = characters_before_decimal(to_photograph_image_no)
    to_photograph_no = characters_after_decimal(to_photograph_image_no)
    
    Set to_snap_photo_image = Read(to_batch_no, to_photograph_no, 0, False, dbOpenDynaset)
    If to_snap_photo_image.RecordCount = 0 Then
        MsgBox "Image number does not exist"
        to_snap_photo_image.Close
        Exit Sub
    End If
    
    Set from_snap_photo_image = Read(0, 0, from_search_result!photograph_key, False, dbOpenDynaset)
'
'   Swap numbers.
'
begin_trans
    '
    '   Make Pending by calling photograph procedure.
    '
    With oSearchResult
        .BatchNo = to_batch_no
        .PhotographNo = to_photograph_no
        .PhotoType = "T"
    End With
    photograph.make_pending oSearchResult, Cancel, lngPhotographSwapNo
    If Cancel Then
        roll_back
        from_search_result.Close
        from_snap_photo_image.Close
        to_snap_photo_image.Close
        Exit Sub
    End If

    With from_search_result
        .Edit
        !photograph_no = to_photograph_no
        !photograph_key = to_snap_photo_image!photograph_key
        .update
    End With
    '
    With poSearchResult
        .PhotographNo = to_photograph_no
        .PhotographKey = to_snap_photo_image!photograph_key
    End With
'
'   Set the status of the from image.
'
    With from_snap_photo_image
        .Edit
        !Status = "I"
        .update
    End With

commit_trans
    
    from_search_result.Close
    from_snap_photo_image.Close
    to_snap_photo_image.Close

    abort = False
'
'********* Code Ends Here **************
'
   Exit Sub
'
change_image_number_Error:
    ErrorRaise "photo_image.change_image_number"
End Sub


Public Function scanned(photograph_image_no As String) As Boolean
    
    Dim batch_no As Long
    Dim photograph_no As Long
    
    batch_no = characters_before_decimal(photograph_image_no)
    photograph_no = characters_after_decimal(photograph_image_no)
    
    SQL = "SELECT * " _
        & "FROM photograph " _
        & "WHERE batch_no = " + Str(batch_no) _
        & " and photograph_no = " + Str(photograph_no) + ";"
    
    Set dyna_photo_image = db.OpenRecordset(SQL, dbOpenSnapshot)
    scanned = dyna_photo_image!scanned

End Function


Public Sub create_new_images(batch_no, no_of_images, abort)
    Dim snap_image_counts As DAO.Recordset
    Dim table_image As DAO.Recordset
    Dim i As Integer
    Dim photograph_no As Long
    
    On Error GoTo error_proc
    abort = True
'
'   Determine maximum image number before creating new ones.
'
    Set snap_image_counts = read_counts(batch_no)
'
'   Create new images.
'
    Set table_image = db.OpenRecordset("photograph", dbOpenTable)
  
    With table_image
        If IsNull(snap_image_counts!max_image_no) Then
            photograph_no = 0
        Else
            photograph_no = snap_image_counts!max_image_no
        End If
        For i = 1 To no_of_images
            photograph_no = photograph_no + 1
            .AddNew
            !batch_no = batch_no
            !photograph_no = photograph_no
            !Status = "I"
            .update
        Next i
    End With
    
    close_recordset snap_image_counts
    close_recordset table_image
    
    abort = False
    
    Exit Sub
    
error_proc:
    MsgBox Str(Err.Number) + " - " + Err.Description, vbCritical
    
    close_recordset snap_image_counts
    close_recordset table_image

End Sub

