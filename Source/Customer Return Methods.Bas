Attribute VB_Name = "customer_return"
Option Explicit
Private abort As Boolean
Private dyna_cust_return As DAO.Recordset
Private snap_cust_return As DAO.Recordset
Private table_customer_return As DAO.Recordset
Private SQL As String

Public Sub cancel_return_photo(customer_return_no, no_of_photographs As Integer, abort As Boolean)
    
10        On Error GoTo error_proc
    
20        abort = True
      '
      '   Retrieve customer return
      '
30        SQL = "SELECT * " _
              & "FROM customer_return " _
              & "WHERE customer_return_no = " + Str(customer_return_no) + ";"

40        Set dyna_cust_return = db.OpenRecordset(SQL, dbOpenDynaset)
      '
      '   Update number of photographs.
      '
50        With dyna_cust_return
60            .Edit
70            !no_of_photographs = !no_of_photographs - no_of_photographs
80            .update
90        End With
    
100       dyna_cust_return.Close
 
110       abort = False
    
120       Exit Sub

error_proc:
130       MsgBox Str(Err.Number) + " - " + Err.Description, vbCritical
140       On Error Resume Next
150       dyna_cust_return.Close

End Sub

Public Sub Delete(customer_return_no As Long, abort)

10        On Error GoTo error_proc
20        abort = True
    
          Dim response As Integer
          Dim msg As String
          Dim style As String
          Dim customer_name As String
          Dim address_line_1 As String
   
30        SQL = "SELECT * " _
              & "FROM customer_return " _
              & "WHERE customer_return_no = " + Str(customer_return_no) + ";"
    
40        Set dyna_cust_return = db.OpenRecordset(SQL, dbOpenDynaset)

50        If dyna_cust_return.RecordCount = 0 Then
60            MsgBox "Customer Return not found", vbExclamation, "Delete Customer"
70            Exit Sub
80        End If
          
90        With dyna_cust_return
100           .Edit
110           msg = "Are you certain you wish to Delete Customer Return number:" + Str(dyna_cust_return!customer_return_no)
120           style = vbYesNo + vbQuestion + vbDefaultButton2
  
130           response = MsgBox(msg, style, "Delete Customer Return")

140           If response = vbNo Then
150               dyna_cust_return.Close
160               Exit Sub
170           End If
180           .Delete
190       End With

200       MsgBox "Customer Return number:" + Str(customer_return_no) + _
                 " has been deleted.", vbInformation, "Delete Customer Return"
    
210       dyna_cust_return.Close
    
220       abort = False
    
230       Exit Sub

error_proc:

240       Select Case Err.Number
  
              Case Is = 3200
250               MsgBox "Customer Return is active and cannot be deleted", vbInformation, "Delete Customer Return"
260           Case Else
270               MsgBox Str(Err.Number) + " - " + Err.Description, vbCritical, "Delete Customer"
280       End Select
    
290       On Error Resume Next
    
300       dyna_cust_return.Close

End Sub


Public Sub create(customer_return_no As Long, _
                  customer_no As Integer, _
                  customer_reference As String, _
                  date_returned As Date, _
                  no_of_photographs As Integer, _
                  received_by As String, _
                  abort As Boolean)

10        abort = True
    
20        On Error GoTo error_proc
    
30        Set table_customer_return = db.OpenRecordset("Customer_return", dbOpenTable)
    
40        With table_customer_return
50            .AddNew
60            !customer_no = customer_no
70            !customer_reference = customer_reference
80            !date_returned = date_returned
90            !received_by = received_by
100           .update
110           .Bookmark = .LastModified
      '
      '   Return the  customer return number.
      '
120           customer_return_no = !customer_return_no
130       End With
    
140       table_customer_return.Close
    
150       abort = False
    
160       Exit Sub
    
error_proc:
170       MsgBox Str(Err.Number) + " - " + Err.Description, vbCritical

180       On Error Resume Next
    
190       table_customer_return.Close
    
End Sub

Private Function customer_return_changed() As Boolean
    
          Dim i As Integer
    
10        customer_return_changed = False
    
20        For i = 0 To snap_cust_return.Fields.Count - 1
30            If Not snap_cust_return.Fields(i).Value = dyna_cust_return.Fields(i).Value Then
40                customer_return_changed = True
50                Exit For
60            End If
70        Next i
End Function

Private Sub get_dyna_cust_return(customer_return_no As Long)
    
10        SQL = "SELECT * " _
              & "FROM customer_return " _
              & "WHERE customer_return_no = " + Str(customer_return_no) + ";"
    
20        Set dyna_cust_return = db.OpenRecordset(SQL, dbOpenDynaset)

End Sub

Public Function Read(customer_return_no) As DAO.Recordset
    
10        SQL = "SELECT * " _
              & "FROM customer_return " _
              & "WHERE customer_return_no = " + Str(customer_return_no) + ";"

20        Set Read = db.OpenRecordset(SQL, dbOpenSnapshot)

End Function


Public Sub return_photo(customer_return_no, no_of_photographs As Integer, abort As Boolean)
    
10        On Error GoTo error_proc
    
20        abort = True
      '
      '   Retrieve customer return
      '
30        SQL = "SELECT * " _
              & "FROM customer_return " _
              & "WHERE customer_return_no = " + Str(customer_return_no) + ";"

40        Set dyna_cust_return = db.OpenRecordset(SQL, dbOpenDynaset)
      '
      '   Update number of photographs.
      '
50        With dyna_cust_return
60            .Edit
70            !no_of_photographs = !no_of_photographs + no_of_photographs
80            .update
90        End With
    
100       dyna_cust_return.Close
 
110       abort = False
    
120       Exit Sub

error_proc:
130       MsgBox Str(Err.Number) + " - " + Err.Description, vbCritical
140       On Error Resume Next
150       dyna_cust_return.Close

End Sub

Public Sub update(in_snap_cust_return As DAO.Recordset, _
                  customer_no As Integer, _
                  customer_reference As String, _
                  date_returned As Date, _
                  no_of_photographs As Integer, _
                  received_by As String, _
                  abort As Boolean)
            
10        abort = True
    
20        Set snap_cust_return = in_snap_cust_return.Clone
30        snap_cust_return.Bookmark = in_snap_cust_return.Bookmark
    
40        get_dyna_cust_return (snap_cust_return!customer_return_no)
50        With dyna_cust_return
60            .Edit
70            If customer_return_changed Then
80                MsgBox "Customer Return has been changed by another user." + vbCr + _
                         "Press OK to see the current details", vbExclamation
90                Exit Sub
100           End If
      '
      '   Update the Customer Return.
      '
110           !customer_no = customer_no
120           !customer_reference = customer_reference
130           !date_returned = date_returned
140           !no_of_photographs = no_of_photographs
150           !received_by = received_by
160           .update
170       End With
    
180       abort = False
190       Exit Sub
    
error_proc:
    
200       MsgBox Str(Err.Number) + " - " + Err.Description, vbCritical

End Sub




