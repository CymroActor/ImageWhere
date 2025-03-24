Attribute VB_Name = "job"
Option Explicit

Private table_delivery_note As DAO.Recordset
        
Private SQL As String
    
Public Enum JobDNoteReportFormat
    jdnOnLine
    jdnPDF
    jndOnLineAndPDF
End Enum


Public Sub print_job_sheet(crs_job As Control, _
                           job_no As Long, _
                           abort As Boolean)
'***************************************
' Module/Form Name   : job
'
' Procedure Name     : print_job_sheet
'
' Purpose            :
'
' Date Created       : 04/05/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : crs_job - Control
'                    : job_no - Long
'                    : abort - Boolean
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo print_job_sheet_Error
'
'******** Code Starts Here *************
'
    abort = True
    
    DoEvents
    
    With crs_job
        .Destination = 0
        .WindowTitle = "Job Sheet: " & CStr(job_no)
        .WindowState = 2
        .WindowBorderStyle = 1
    End With

'   Set the Record Selection formula.

    crs_job.SelectionFormula = "{Job.Job_no} = " + Str(job_no)

'   Set location of the database file to the location from
'   which the user ran the app

    crs_job.DataFiles(0) = glo_dbname
    crs_job.DataFiles(1) = glo_dbname
    crs_job.DataFiles(2) = glo_dbname
    crs_job.DataFiles(3) = glo_dbname
    crs_job.DataFiles(4) = glo_dbname
    crs_job.DataFiles(5) = glo_dbname
    crs_job.DataFiles(6) = glo_dbname
    'crs_job.ReportFileName = App.Path + "\Reports\Job.rpt"
    crs_job.ReportFileName = goSystemConfig.ServerLocation + "\Reports\Job.rpt"

'   EXECUTE PRINT CALL
    On Error GoTo print_job_sheet_Error
    
    crs_job.Action = 1
'
'   Maximize the window.
'
'    maximize_window
'
    abort = False
'
'********* Code Ends Here **************
'
   Exit Sub
'
print_job_sheet_Error:
    Select Case Err.Number
        Case Is = cdlCancel
            abort = False
        Case Else
            ErrorRaise "job.print_job_sheet"
    End Select
End Sub

Public Function Read(job_no) As DAO.Recordset
    
    SQL = "SELECT * " _
        & "FROM job " _
        & "WHERE job_no = " + Str(job_no) + ";"

    Set Read = db.OpenRecordset(SQL, dbOpenSnapshot)

End Function

Public Function read_for_delivery_note(delivery_note_no) As DAO.Recordset
    
    SQL = "SELECT * " _
        & "FROM job " _
        & "WHERE delivery_note_no = " + Str(delivery_note_no) + ";"

    Set read_for_delivery_note = db.OpenRecordset(SQL, dbOpenSnapshot)

End Function

Public Sub print_customer_label(cdg_print As Control, crs_delivery_note As Control, job_no As Long, abort As Boolean)
    
    abort = True
    
    On Error GoTo error_proc
    
    With cdg_print
        .FLAGS = cdlPDPrintSetup
        .CancelError = True
        .ShowPrinter
    End With
    
    crs_delivery_note.Destination = 1
    crs_delivery_note.CopiesToPrinter = cdg_print.Copies

'   Set the Record Selection formula.

    crs_delivery_note.SelectionFormula = "{Job.job_no} = " + Str(job_no)

'   Set location of the database file to the location from
'   which the user ran the app

    crs_delivery_note.DataFiles(0) = glo_dbname
    'crs_delivery_note.ReportFileName = App.Path + "\Reports\Delivery Note Customer Label.rpt"
    crs_delivery_note.ReportFileName = goSystemConfig.ServerLocation + "\Reports\Delivery Note Customer Label.rpt"

'   EXECUTE PRINT CALL
    
    crs_delivery_note.Action = 1
    
    abort = False
    
    Exit Sub

error_proc:

    Select Case Err.Number
        Case Is = cdlCancel
            abort = False
            Exit Sub
        Case Else
            MsgBox Str(Err.Number) + " - " + Err.Description, vbCritical
    End Select

End Sub

Private Sub close_delivery_recordsets()
    On Error GoTo error_proc
    
    table_delivery_note.Close
''    table_delivery_note_no.Close
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


Public Sub deliver(delivery_note_form As Form, _
                   oJob As Job2, _
                   delivery_note_no As Long, _
                   abort As Boolean)
    '***************************************
    ' Module/Form Name   : job
    '
    ' Procedure Name     : deliver
    '
    ' Purpose            :
    '
    ' Date Created       : 29/07/2001
    '
    ' Author             : GARETH
    '
    ' Parameters         : delivery_note_form - Form
    '                    : in_snap_job - DAO.Recordset
    '                    : delivery_note_no - Long
    '                    : abort - Boolean
    '
    ' Amendment History  : Date       Author    Description
    '                    : --------------------------------
    '                    : 21/04/2002 GARETH SAUNDERS
    '
    '***************************************
    '
    On Error GoTo deliver_Error
    '
    '******** Code Starts Here *************
    '

    Dim oChaser As Chaser
    Dim oBusinessType As BusinessType
    Dim oCustomer As Customer2

    abort = True
    '
    '   Lock the Delivery Note Number table to stop anyone from updating it while it is being
    '   used.
    '
    Set table_delivery_note = db.OpenRecordset("Delivery_note", dbOpenTable)
''    Set table_delivery_note = New ADODB.Recordset
''    table_delivery_note.Open "Delivery_note", gdbADO, adOpenDynamic, adLockPessimistic
    '
    '   Create the Delivery Note record
    '
    With table_delivery_note
        .AddNew
        !customer_no = oJob.CustomerNo
        !Memo = delivery_note_form.txt_memo.Text
        !date_created = Date
        !date_printed = Date
        '!return_date = delivery_note_form.txt_return_date.Text
        !DispatchTransparency = delivery_note_form.txtDispatchTransparency
        !DispatchDigital = delivery_note_form.txtDispatchDigital
        !service_fee = delivery_note_form.txt_service_fee.Text
        !weekly_holding_fee = delivery_note_form.txt_weekly_holding_fee.Text
        !loss_damage_fee = delivery_note_form.txt_loss_damage_fee.Text
        .update
        .Bookmark = .LastModified
    End With
    '
    '   Pass back the Delivery Note number to the calling procedure
    '
    delivery_note_no = table_delivery_note!delivery_note_no
    '
    '   Create the Chaser record.
    '
    Set oChaser = New Chaser
    oChaser.CreateDAO delivery_note_no, _
                      oJob.RequestedBy, _
                      "None", _
                      "", _
                      delivery_note_form.txt_return_date.Value, _
                      "None", _
                      "Initial Default Record"
    '
    'oChaser.Create delivery_note_no, _
                   oJob.RequestedBy, _
                   "None", _
                   "", _
                   delivery_note_form.txt_return_date.Value, _
                   "None", _
                   "Initial Default Record"
    '
    Set oCustomer = New Customer2
    oCustomer.Read oJob.CustomerNo
    Set oBusinessType = New BusinessType
    oBusinessType.Read oCustomer.BusinessType
    oChaser.CreateDAO delivery_note_no, _
                      oJob.RequestedBy, _
                      "None", _
                      "", _
                      DateAdd("d", oBusinessType.SL1ReturnPeriod, delivery_note_form.txt_return_date.Value), _
                      "SL1", _
                      "Initial Default Record"
    Set oBusinessType = Nothing
    Set oCustomer = Nothing
    '
    '   Update the Job record
    '
    With oJob
        .DeliveryNoteNo = table_delivery_note!delivery_note_no
        .update
    End With
    '
    '   Deliver each of the search_result photographs
    '
    delivery_note.deliver_job_search_results table_delivery_note!delivery_note_no, _
                                             oJob.JobNo, _
                                             abort
    If abort Then
        close_delivery_recordsets
        Exit Sub
    End If

    close_delivery_recordsets
    Set oChaser = Nothing

    '
    '********* Code Ends Here **************
    '
    Exit Sub
    '
deliver_Error:
    ErrorRaise "job.deliver"
End Sub


Public Sub print_all_delivery_notes(ByRef crs_delivery_note As Crystal.CrystalReport, _
                                    ByVal pReportFormat As JobDNoteReportFormat, _
                                    ByVal pstrOutputDestination As String, _
                                    ByVal abort As Boolean)
'***************************************
' Module/Form Name   : job
'
' Procedure Name     : print_delivery_note
'
' Purpose            :
'
' Date Created       : 04/05/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : crs_delivery_note - Crystal.CrystalReport
'                    : delivery_note_no - Long
'                    : customer_no - Integer
'                    : type_of_print - As
'                    : Boolean -
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo print_delivery_note_Error
'
'******** Code Starts Here *************
'
    Dim strSelectionFormula             As String
    Dim strReportFileName               As String
    Dim strWindowTitle                  As String
    Dim crxApp                          As New CRAXDRT.Application
    Dim crxRpt                          As CRAXDRT.Report
    Dim crxTab                          As CRAXDRT.DatabaseTable
    Dim crxSection                      As CRAXDRT.Section
    Dim ReportObject                    As Object
    Dim crxSubreport                    As CRAXDRT.Report
    Dim crxSubreportObj                 As CRAXDRT.SubreportObject
    Dim strSubReportSelectionFormula    As String
    
    abort = True
    
    DoEvents
    
    strWindowTitle = "All Delivery Notes between March 1st 2004 and 28th February 2005"
    'strReportFileName = App.Path & "\Reports\All Delivery Notes.rpt"
    strReportFileName = goSystemConfig.ServerLocation & "\Reports\All Delivery Notes.rpt"
    
    strSelectionFormula = "{Delivery_note.date_created} > Date(2004,02,29) and {Delivery_note.date_created} < Date (2004,03,02) AND {Chaser.Action} ='None' AND {Chaser.NextAction} ='None'"
    '
    '   Set up the selection formula for the sub report.
    '
    'strSubReportSelectionFormula = "{Delivery_note.delivery_note_no} = " & CStr(delivery_note_no)
    '
    '   What type of CR Report is it?
    '
    If pReportFormat = jdnOnLine Or pReportFormat = jndOnLineAndPDF Then
        With crs_delivery_note
            .Destination = 0
            .WindowState = 2
            .WindowControls = True
            .WindowControlBox = True
            '
            '   Set up the report parameters.
            '
            .WindowTitle = strWindowTitle
            .ReportFileName = strReportFileName
            .SelectionFormula = strSelectionFormula
            '
            '   Set location of the database file to the location from
            '   which the user ran the app
            '
'            .DataFiles(0) = glo_dbname
'            .DataFiles(1) = glo_dbname
'            .DataFiles(2) = glo_dbname
'            .DataFiles(3) = glo_dbname
'            .DataFiles(4) = glo_dbname
'            .DataFiles(5) = glo_dbname
'            .DataFiles(6) = glo_dbname
'            .DataFiles(7) = glo_dbname
'            .DataFiles(8) = glo_dbname
            '
            '   Set up the sub report.
            '
''            .SubreportToChange = "AbridgedTermsAndConditions.rpt"
''            .DataFiles(0) = glo_dbname
''            .SelectionFormula = strSubReportSelectionFormula
        End With
        '
        '   EXECUTE PRINT CALL
        '
        crs_delivery_note.Action = 1
        'Print/Export a report to a dummy file in the backgroud, the purpose of this section of code
        'is to just release first report from memory:
        With crs_delivery_note
            '.ReportFileName = App.Path & "\Reports\Dummy2.rpt"
            .ReportFileName = goSystemConfig.ServerLocation & "\Reports\Dummy2.rpt"
            .Destination = crptToFile
            .PrintFileType = crptCrystal
            .PrintFileName = "c:\temp\temp.rpt"
            'To turn off ProgressDialog so the user won't notice a report is printing on the background:
            .ProgressDialog = False
            On Error Resume Next
            .Action = 1
        End With
    ElseIf pReportFormat = jdnPDF Or pReportFormat = jndOnLineAndPDF Then
        '
        '   Print report to a PDF.
        '
        Set crxRpt = crxApp.OpenReport(strReportFileName)
    
        'Set the export options
        crxRpt.ExportOptions.DestinationType = crEDTDiskFile
        For Each crxTab In crxRpt.Database.Tables
            crxTab.Location = glo_dbname
        Next crxTab
        crxRpt.RecordSelectionFormula = strSelectionFormula
        '
        '   Now fix the sub report.
        '
        For Each crxSection In crxRpt.Sections
            For Each ReportObject In crxSection.ReportObjects
                If ReportObject.Kind = crSubreportObject Then
                    Set crxSubreportObj = ReportObject
                    Set crxSubreport = crxSubreportObj.OpenSubreport
                    For Each crxTab In crxSubreport.Database.Tables
                        crxTab.Location = glo_dbname
                    Next crxTab
                    crxSubreport.RecordSelectionFormula = strSubReportSelectionFormula
                End If
            Next ReportObject
        Next crxSection
        crxRpt.ExportOptions.PDFExportAllPages = True
        crxRpt.ExportOptions.FormatType = crEFTPortableDocFormat
        crxRpt.ExportOptions.DiskFileName = pstrOutputDestination
        
        'Export the report, without prompting the user
        'Another dialog will show with the status of the export
        'There is no way to suppress the status dialog.
        crxRpt.Export False
        '
        Set crxTab = Nothing
        Set crxSection = Nothing
        Set ReportObject = Nothing
        Set crxSubreport = Nothing
        Set crxSubreportObj = Nothing
        Set crxRpt = Nothing
        Set crxApp = Nothing
    
    End If
    '
    abort = False
'
'********* Code Ends Here **************
'
   Exit Sub
'
print_delivery_note_Error:
    Select Case Err.Number
        Case Is = cdlCancel
            abort = False
            Exit Sub
        Case Else
        ErrorRaise "job.print_delivery_note"
    End Select
End Sub

Public Sub print_delivery_note(ByRef crs_delivery_note As Crystal.CrystalReport, _
                               ByVal delivery_note_no As Long, _
                               ByVal customer_no As Integer, _
                               ByVal type_of_print, _
                               ByVal pReportFormat As JobDNoteReportFormat, _
                               ByVal pstrOutputDestination As String, _
                               ByVal abort As Boolean)
'***************************************
' Module/Form Name   : job
'
' Procedure Name     : print_delivery_note
'
' Purpose            :
'
' Date Created       : 04/05/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : crs_delivery_note - Crystal.CrystalReport
'                    : delivery_note_no - Long
'                    : customer_no - Integer
'                    : type_of_print - As
'                    : Boolean -
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo print_delivery_note_Error
'
'******** Code Starts Here *************
'
    Dim strSelectionFormula             As String
    Dim strReportFileName               As String
    Dim strWindowTitle                  As String
    Dim crxApp                          As New CRAXDRT.Application
    Dim crxRpt                          As CRAXDRT.Report
    Dim crxTab                          As CRAXDRT.DatabaseTable
    Dim crxSection                      As CRAXDRT.Section
    Dim ReportObject                    As Object
    Dim crxSubreport                    As CRAXDRT.Report
    Dim crxSubreportObj                 As CRAXDRT.SubreportObject
    Dim strSubReportSelectionFormula    As String
    
    abort = True
    
    DoEvents
    '
    '   Set up the selection formula.
    '
    Select Case type_of_print
        Case Is = "O"
            strSelectionFormula = "{Delivery_note.delivery_note_no} = " + Str(delivery_note_no)
            strSelectionFormula = strSelectionFormula & " and {Customer.customer_no} = " + Str(customer_no) + " and "
            strSelectionFormula = strSelectionFormula & "(isnull({Search_result.date_returned}) or {Search_result.date_returned}=Date(1899,12,30) or {Search_result.date_returned}=Date(0,0,0)) AND "
            strSelectionFormula = strSelectionFormula & "{Chaser.Action} ='None' AND {Chaser.NextAction} ='None'"
        Case Is = "A"
            strSelectionFormula = "{Delivery_note.delivery_note_no} = " + Str(delivery_note_no)
            strSelectionFormula = strSelectionFormula & " AND {Customer.customer_no} = " + Str(customer_no)
            strSelectionFormula = strSelectionFormula & " AND {Chaser.Action} ='None' AND {Chaser.NextAction} ='None'"
        Case Else
            MsgBox "Software problem: please contact your administrator", vbCritical
            Exit Sub
    End Select
    '
    '   What type of report is it?
    '
    Select Case type_of_print
        Case Is = "O"
            strWindowTitle = "Print Outstanding Pictures for Delivery Note: " & CStr(delivery_note_no)
            'strReportFileName = App.Path & "\Reports\Outstanding Pictures.rpt"
            strReportFileName = goSystemConfig.ServerLocation & "\Reports\Outstanding Pictures.rpt"
        Case Is = "A"
            strWindowTitle = "Print Delivery Note: " & CStr(delivery_note_no)
            'strReportFileName = App.Path & "\Reports\Delivery Note.rpt"
            strReportFileName = goSystemConfig.ServerLocation & "\Reports\Delivery Note.rpt"
        Case Else
            MsgBox "Software problem: please contact your administrator", vbCritical
            Exit Sub
    End Select
    '
    '   Set up the selection formula for the sub report.
    '
    strSubReportSelectionFormula = "{Delivery_note.delivery_note_no} = " & CStr(delivery_note_no)
    '
    '   What type of CR Report is it?
    '
    If pReportFormat = jdnOnLine Or pReportFormat = jndOnLineAndPDF Then
        With crs_delivery_note
            .Destination = 0
            .WindowState = 2
            .WindowControls = True
            .WindowControlBox = True
            '
            '   Set up the report parameters.
            '
            .WindowTitle = strWindowTitle
            .ReportFileName = strReportFileName
            .SelectionFormula = strSelectionFormula
            '
            '   Set location of the database file to the location from
            '   which the user ran the app
            '
            .DataFiles(0) = glo_dbname
            .DataFiles(1) = glo_dbname
            .DataFiles(2) = glo_dbname
            .DataFiles(3) = glo_dbname
            .DataFiles(4) = glo_dbname
            .DataFiles(5) = glo_dbname
            .DataFiles(6) = glo_dbname
            .DataFiles(7) = glo_dbname
            .DataFiles(8) = glo_dbname
            '
            '   Set up the sub report.
            '
            .SubreportToChange = "AbridgedTermsAndConditions.rpt"
            .DataFiles(0) = glo_dbname
            .SelectionFormula = strSubReportSelectionFormula
        End With
        '
        '   EXECUTE PRINT CALL
        '
        crs_delivery_note.Action = 1
        'Print/Export a report to a dummy file in the backgroud, the purpose of this section of code
        'is to just release first report from memory:
        With crs_delivery_note
            '.ReportFileName = App.Path & "\Reports\Dummy2.rpt"
            .ReportFileName = goSystemConfig.ServerLocation & "\Reports\Dummy2.rpt"
            .Destination = crptToFile
            .PrintFileType = crptCrystal
            .PrintFileName = "c:\temp\temp.rpt"
            'To turn off ProgressDialog so the user won't notice a report is printing on the background:
            .ProgressDialog = False
            On Error Resume Next
            .Action = 1
        End With
    ElseIf pReportFormat = jdnPDF Or pReportFormat = jndOnLineAndPDF Then
        '
        '   Print report to a PDF.
        '
        Set crxRpt = crxApp.OpenReport(strReportFileName)
    
        'Set the export options
        crxRpt.ExportOptions.DestinationType = crEDTDiskFile
        For Each crxTab In crxRpt.Database.Tables
            crxTab.Location = glo_dbname
        Next crxTab
        crxRpt.RecordSelectionFormula = strSelectionFormula
        '
        '   Now fix the sub report.
        '
        For Each crxSection In crxRpt.Sections
            For Each ReportObject In crxSection.ReportObjects
                If ReportObject.Kind = crSubreportObject Then
                    Set crxSubreportObj = ReportObject
                    Set crxSubreport = crxSubreportObj.OpenSubreport
                    For Each crxTab In crxSubreport.Database.Tables
                        crxTab.Location = glo_dbname
                    Next crxTab
                    crxSubreport.RecordSelectionFormula = strSubReportSelectionFormula
                End If
            Next ReportObject
        Next crxSection
        crxRpt.ExportOptions.PDFExportAllPages = True
        crxRpt.ExportOptions.FormatType = crEFTPortableDocFormat
        crxRpt.ExportOptions.DiskFileName = pstrOutputDestination & "\" & CStr(delivery_note_no) & ".pdf"
        
        crxRpt.Export (False)
        
        'Export the report, without prompting the user
        'Another dialog will show with the status of the export
        'There is no way to suppress the status dialog.
        ''crxRpt.Export False
        
        '
        Set crxTab = Nothing
        Set crxSection = Nothing
        Set ReportObject = Nothing
        Set crxSubreport = Nothing
        Set crxSubreportObj = Nothing
        Set crxRpt = Nothing
        Set crxApp = Nothing
    
    End If
    '
    abort = False
'
'********* Code Ends Here **************
'
   Exit Sub
'
print_delivery_note_Error:
    Select Case Err.Number
        Case Is = cdlCancel
            abort = False
            Exit Sub
        Case Else
        ErrorRaise "job.print_delivery_note"
    End Select
End Sub


Public Function read_job_images(job_no As Long) As DAO.Recordset
    
    SQL = "SELECT search_result.* " _
        & "FROM search_result, search, job " _
        & "WHERE search.search_no = search_result.search_no and job.job_no = search.job_no " _
        & "and job.job_no = " + Str(job_no) + ";"

    Set read_job_images = db.OpenRecordset(SQL, dbOpenSnapshot)
  
End Function


