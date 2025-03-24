Attribute VB_Name = "batch"
Option Explicit
Private dyna_batch As DAO.Recordset
Private SQL As String
Private abort As Boolean
Public Sub cancel_return_photo(batch_no, abort)
      '
      '   Decrements the count on the batch record to indicate that a
      '   photograph has been returned.
      '
10        On Error GoTo error_proc
    
20        abort = True
      '
      '   Retrieve batch
      '
30        SQL = "SELECT * " _
              & "FROM batch " _
              & "WHERE batch_no = " + Str(batch_no) + ";"

40        Set dyna_batch = db.OpenRecordset(SQL, dbOpenDynaset)
      '
      '   Decrement the 'number of photographs in stock' by 1.
      '
50        With dyna_batch
60            .Edit
70            If !total_no_of_photos = 0 Then
80                MsgBox "Data error: Photograph " + Str(!batch_no) + _
                          " - All known photographs are currently out to customers" + vbCr + "ABORTING TRANSACTION", vbCritical
90                dyna_batch.Close
100               Exit Sub
110           End If
120           !no_of_photos_in_stock = !no_of_photos_in_stock - 1
130           .update
140       End With
    
150       dyna_batch.Close
 
160       abort = False
    
170       Exit Sub

error_proc:
180       MsgBox Str(Err.Number) + " - " + Err.Description, vbCritical
190       On Error Resume Next
200       dyna_batch.Close

End Sub


Public Sub return_photo(batch_no, abort)
    
      '
      '   Increments the count on the batch record to indicate that a
      '   photograph has been returned.
      '
10        On Error GoTo error_proc
    
20        abort = True
      '
      '   Retrieve batch
      '
30        SQL = "SELECT * " _
              & "FROM batch " _
              & "WHERE batch_no = " + Str(batch_no) + ";"

40        Set dyna_batch = db.OpenRecordset(SQL, dbOpenDynaset)
      '
      '   Increment the 'number of photographs in stock' by 1.
      '
50        With dyna_batch
60            .Edit
70            If !total_no_of_photos = !no_of_photos_in_stock Then
80                MsgBox "Data error: Photograph " + Str(!batch_no) + _
                          " - All known photographs are currently in stock" + vbCr + "ABORTING TRANSACTION", vbCritical
90                dyna_batch.Close
100               Exit Sub
110           End If
120           !no_of_photos_in_stock = !no_of_photos_in_stock + 1
130           .update
140       End With
    
150       dyna_batch.Close
 
160       abort = False
    
170       Exit Sub

error_proc:
180       MsgBox Str(Err.Number) + " - " + Err.Description, vbCritical
190       On Error Resume Next
200       dyna_batch.Close

End Sub

