Attribute VB_Name = "swapping_bugs"
Option Explicit
Public Sub IN_bug()

End Sub


Public Sub out_bug()
          Dim SQL As String
          Dim snap_error_srs As DAO.Recordset
          Dim dyna_sr As DAO.Recordset
          Dim next_dyna_photo_image As DAO.Recordset
    
10        On Error GoTo error_proc
      '
      '   Get errored Images/Search Results
      '
20        ws.BeginTrans
30        SQL = "SELECT Photograph.batch_no, photograph.photograph_no, " _
              & "Count(photograph.photograph_key) AS ['Count 0f Nulls'] " _
              & "FROM Photograph INNER JOIN Search_result " _
              & "ON Photograph.photograph_key = Search_result.photograph_key " _
              & "Where ((photograph.status) = 'O') And ((search_result.date_returned) Is Null) " _
              & "GROUP BY Photograph.batch_no, photograph.photograph_no " _
              & "HAVING Count(photograph.photograph_key) > 1;"
    
40        Set snap_error_srs = db.OpenRecordset(SQL, dbOpenSnapshot)
    
50        If snap_error_srs.RecordCount = 0 Then
60            MsgBox "There are none in error!"
70            snap_error_srs.Close
80            ws.Rollback
90            Exit Sub
100       End If
      '
      '   For each Image, fix the erroneous Search Result.
      '
110       With snap_error_srs
120           .MoveFirst
130           While Not .EOF
140               SQL = "SELECT DISTINCTROW Search_result.* " _
                      & "FROM Photograph INNER JOIN Search_result " _
                      & "ON Photograph.photograph_key = Search_result.photograph_key " _
                      & "WHERE Search_result.batch_no = " + Str(!batch_no) _
                      & " AND Search_result.photograph_no = " + Str(!photograph_no) _
                      & " ORDER BY delivery_note_no;"
      
150               Set dyna_sr = db.OpenRecordset(SQL, dbOpenDynaset)
160               If dyna_sr.RecordCount <> 2 Then
170                   MsgBox "Batch: " + !batch_no + "/Photograph: " + !photograph_no + _
                             " has " + Str(dyna_sr.RecordCount) + " erroneous search results. " + _
                             vbCr + "Aborting Fix."
180                   dyna_sr.Close
190                   ws.Rollback
200                   Exit Sub
210               End If
      '
      '   Get next available non-scanned image number for the erroneous search result.
      '
220               Set next_dyna_photo_image = _
                      photo_image.get_next_available_image(!batch_no, _
                                                           dbOpenDynaset, _
                                                           True, _
                                                           abort)
    
230               If abort Then
240                   snap_error_srs.Close
250                   next_dyna_photo_image.Close
260                   ws.Rollback
270                   Exit Sub
280               End If
      '
      '   Amend the lower delivery note number, ie. the first.
      '
290               With dyna_sr
300                   .MoveFirst
310                   .Edit
320                   !photograph_no = next_dyna_photo_image!photograph_no
330                   !photograph_key = next_dyna_photo_image!photograph_key
340                   .update
350               End With
      '
      '   Make the next available image number the one that is out.
      '
360               With next_dyna_photo_image
370                   .Edit
380                   !Status = "O"
390                   .update
400               End With
    
410           Wend
420       End With
430       ws.CommitTrans
    
440       snap_error_srs.Close
450       dyna_sr.Close
460       next_dyna_photo_image.Close
    
470       MsgBox "All errors fixed!"
    
480       Exit Sub
    
error_proc:
490       MsgBox Str(Err.Number) + " - " + Err.Description, vbCritical
500       On Error Resume Next
    
510       snap_error_srs.Close
520       dyna_sr.Close
530       next_dyna_photo_image.Close

End Sub



