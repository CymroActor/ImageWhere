Attribute VB_Name = "customer"
Option Explicit
Private SQL As String
Private dyna_customer As DAO.Recordset
Public Sub print_label(cdg_print As Control, crs_delivery_note As Control, customer_no As Long, abort As Boolean)
    
10        abort = True
    
20        On Error GoTo error_proc
    
30        With cdg_print
40            .DialogTitle = "Set printer for Customer Label"
50            .FLAGS = cdlPDPrintSetup
60            .CancelError = True
70            .ShowPrinter
80        End With
    
90        DoEvents
    
100       crs_delivery_note.Destination = 0
110       crs_delivery_note.CopiesToPrinter = cdg_print.Copies
120       crs_delivery_note.WindowTitle = "Print Customer Label: " + customer.Read(customer_no)!customer_name

      '   Set the Record Selection formula.

130       crs_delivery_note.SelectionFormula = "{Customer.customer_no} = " + Str(customer_no)

      '   Set location of the database file to the location from
      '   which the user ran the app

140       crs_delivery_note.DataFiles(0) = glo_dbname
'150       crs_delivery_note.ReportFileName = App.Path + "\Reports\Customer Label.rpt"
150       crs_delivery_note.ReportFileName = goSystemConfig.ServerLocation + "\Reports\Customer Label.rpt"

      '   EXECUTE PRINT CALL
    
160       crs_delivery_note.Action = 1
    
170       abort = False
    
180       Exit Sub

error_proc:

190       Select Case Err.Number
              Case Is = cdlCancel
200               abort = False
210               Exit Sub
220           Case Else
230               MsgBox Str(Err.Number) + " - " + Err.Description, vbCritical
240       End Select

End Sub

Public Function Read(customer_no) As DAO.Recordset
    
10        SQL = "SELECT * " _
              & "FROM customer " _
              & "WHERE customer_no = " + Str(customer_no) + ";"

20        Set Read = db.OpenRecordset(SQL, dbOpenSnapshot)

End Function

Private Sub close_customer_recordsets()
10        On Error GoTo error_proc
    
20        dyna_customer.Close
    
30        Exit Sub
    
error_proc:
40        Select Case Err.Number
              Case Is = 91
50                Resume Next
60            Case Is = 3420
70                Resume Next
80            Case Else
90                MsgBox Str(Err.Number) + " - " + Err.Description, vbCritical
100       End Select
End Sub



Public Sub Delete(customer_no As Integer)

10        On Error GoTo error_proc
    
          Dim response As Integer
          Dim msg As String
          Dim style As String
          Dim customer_name As String
          Dim address_line_1 As String
   
20        SQL = "SELECT * " _
              & "FROM customer " _
              & "WHERE customer_no = " + Str(customer_no) + ";"
    
30        Set dyna_customer = db.OpenRecordset(SQL, dbOpenDynaset)

40        If dyna_customer.RecordCount = 0 Then
50            MsgBox "Customer not found", vbExclamation, "Delete Customer"
60            dyna_customer.Close
70            Exit Sub
80        End If
          
90        With dyna_customer
100           .Edit
110           msg = "Are you certain you wish to Delete Customer '" + dyna_customer!customer_name + "'" + vbCr _
                  & "First line of address: " + dyna_customer!address_line_1 + vbCr
120           style = vbYesNo + vbQuestion + vbDefaultButton2
  
130           response = MsgBox(msg, style, "Delete Customer")

140           If response = vbNo Then
150               dyna_customer.Close
160               Exit Sub
170           End If
     
180           customer_name = dyna_customer!customer_name
190           address_line_1 = dyna_customer!address_line_1
  
200           .Delete
210       End With

220       MsgBox "Customer '" + customer_name + "'" + vbCr + "First line of address: " + address_line_1 + vbCr _
               & "has been deleted.", vbInformation, "Delete Customer"
    
230       close_customer_recordsets
    
240       Exit Sub

error_proc:

250       Select Case Err.Number
  
              Case Is = 3200
260               MsgBox "Customer has an active history and cannot be deleted", vbExclamation, "Delete Customer"
270           Case Else
280               MsgBox Str(Err.Number) + " - " + Err.Description, vbCritical, "Delete Customer"
290       End Select
    
300       close_customer_recordsets

End Sub


