VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReportRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mvarId As Long
Private mvarName As String
Private mvarParameter1 As String
Private mvarParameter2 As String
Private mvarParameter3 As String
'
'   ADO Recordset used for data access.
'
Private mrstReportRequest As ADODB.Recordset
'
'   ADO Recordset to hold a checkpointed copy of the data.
'
Private mrstCheckPoint As ADODB.Recordset

Public Property Let ID(ByVal vData As String)
10        mvarId = vData
End Property

Public Property Get ID() As String
10        ID = mvarId
End Property

Public Property Let Name(ByVal vData As String)
10        mvarName = vData
End Property

Public Property Get Name() As String
10        Name = mvarName
End Property

Public Property Let Parameter1(ByVal vData As String)
10        mvarParameter1 = vData
End Property

Public Property Get Parameter1() As String
10        Parameter1 = mvarParameter1
End Property

Public Property Let Parameter2(ByVal vData As String)
10        mvarParameter2 = vData
End Property

Public Property Get Parameter2() As String
10        Parameter2 = mvarParameter2
End Property

Public Property Let Parameter3(ByVal vData As String)
10        mvarParameter3 = vData
End Property

Public Property Get Parameter3() As String
10        Parameter3 = mvarParameter3
End Property

Public Function Read(Name As String) As Boolean
'***************************************
' Module/Form Name   : ReportRequest
'
' Procedure Name     : Read
'
' Purpose            :
'
' Date Created       : 29/07/2001
'
' Author             : GARETH
'
' Parameters         : Name - String
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
    Dim strSQL As String
    
    strSQL = "SELECT * " _
           & "FROM REPORTREQUEST " _
           & "WHERE NAME = '" & Replace(Name, "'", "''") & "'"
    
    Set mrstReportRequest = New ADODB.Recordset
    mrstReportRequest.Open strSQL, gdbADO, , , adCmdText
    If mrstReportRequest.EOF Then
        Read = False
    Else
        Refresh mrstReportRequest
        mrstReportRequest.Close
        Read = True
    End If
    Set mrstReportRequest = Nothing
'
'********* Code Ends Here **************
'
   Exit Function
'
Read_Error:
    ErrorRaise "ImageWhere.ReportRequest.Read"
End Function

Public Sub CheckPoint()
10        With mrstCheckPoint
20            !ID = mvarId
30            !Name = mvarName
40            !Parameter1 = mvarParameter1
50            !Parameter2 = mvarParameter2
60            !Parameter3 = mvarParameter3
70        End With
End Sub

Private Sub Class_Initialize()
    
10        Set mrstCheckPoint = New ADODB.Recordset
20        With mrstCheckPoint
30            .Fields.Append "Id", adInteger
40            .Fields.Append "Name", adVarWChar, 20
50            .Fields.Append "Parameter1", adVarWChar, 100
60            .Fields.Append "Parameter2", adVarWChar, 20
70            .Fields.Append "Parameter3", adVarWChar, 20
80        End With
90        With mrstCheckPoint
100           .Open
110           .AddNew
120           !ID = 0
130           !Name = ""
140           !Parameter1 = ""
150           !Parameter2 = ""
160           !Parameter3 = ""
170           .update
180       End With
End Sub

Private Function CheckRecordSame(rstRecord As ADODB.Recordset) As Boolean
          Dim fldCheckpoint As ADODB.Field
    
10        CheckRecordSame = False
          '
          '   Check if record has changed by comparing it with the data
          '   at last refresh.
          '
20        With rstRecord
30            For Each fldCheckpoint In rstRecord.Fields
40                If fldCheckpoint.Value <> mrstCheckPoint.Fields(fldCheckpoint.Name).Value Then
50                    Refresh rstRecord
60                    Err.Raise vbObjectError + 1, "ReportRequest: CheckRecordSame", "ReportRequest has been updated since first obtaining it."
70                End If
80            Next fldCheckpoint
90        End With
100       CheckRecordSame = True
End Function

Private Function Refresh(rstRecord As ADODB.Recordset)
10        With rstRecord
20            mvarId = !ID
30            mvarName = Format(!Name)
40            mvarParameter1 = Format(!Parameter1)
50            mvarParameter2 = Format(!Parameter2)
60            mvarParameter3 = Format(!Parameter3)
70        End With
80        CheckPoint
End Function

Public Sub update()
    Dim strSQL              As String
    Dim rstReportRequest    As DAO.Recordset
    
    On Error GoTo ErrorProc
    
    strSQL = "SELECT * FROM REPORTREQUEST " _
           & "WHERE NAME = '" & Replace(mvarName, "'", "''") & "'"
    
    Set rstReportRequest = db.OpenRecordset(strSQL, dbOpenDynaset)
    If rstReportRequest.RecordCount <> 1 Then
        Err.Raise vbObjectError + 2, , "Report Request does not Exist"
    End If
    
    With rstReportRequest
        .Edit
        !Parameter1 = mvarParameter1
        !Parameter2 = mvarParameter2
        !Parameter3 = mvarParameter3
        .update
    End With
    
    Exit Sub

ErrorProc:
    Err.Raise Err.Number, "ReportRequest: Update"
End Sub

Public Sub Create()
'***************************************
' Module/Form Name   : ReportRequest
'
' Procedure Name     : Create
'
' Purpose            :
'
' Date Created       : 20/02/2005
'
' Author             :
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
   On Error GoTo ErrorHandler:
'
'******** Code Starts Here *************
'
    Dim strSQL As String
    
    strSQL = "INSERT INTO REPORTREQUEST (Name, Parameter1, Parameter2, Parameter3) VALUES (" & _
             "'" & Replace(mvarName, "'", "''") & "', " & _
             "'" & Replace(mvarParameter1, "'", "''") & "'," & _
             "'" & Replace(mvarParameter2, "'", "''") & "'," & _
             "'" & Replace(mvarParameter3, "'", "''") & "')"
    
    gdbADO.Execute strSQL
  
    CheckPoint
'
'********* Code Ends Here **************
'
   On Error GoTo 0
   Exit Sub
'
ErrorHandler:
    Err.Raise Err.Number, "ReportRequest.Create"
End Sub

