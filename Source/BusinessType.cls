VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BusinessType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Private mvarName As String
Private mvarInitialReturnPeriod As Integer
Private mvarSL1ReturnPeriod As Integer
Private mvarSL2ReturnPeriod As Integer
Private mvarPhone1ReturnPeriod As Integer
Private mvarPhone2ReturnPeriod As Integer
Private mvarLossFeeReturnPeriod As Integer
'
'   ADO Recordset to hold a checkpointed copy of the data.
'
Private mrstCheckPoint As ADODB.Recordset

Public Property Let Name(vData As String)
10        mvarName = vData
End Property

Public Property Get Name() As String
10        Name = mvarName
End Property

Public Property Let InitialReturnPeriod(vData As Integer)
10        mvarInitialReturnPeriod = vData
End Property

Public Property Get InitialReturnPeriod() As Integer
10        InitialReturnPeriod = mvarInitialReturnPeriod
End Property

Public Property Let SL1ReturnPeriod(vData As Integer)
10        mvarSL1ReturnPeriod = vData
End Property

Public Property Get SL1ReturnPeriod() As Integer
10        SL1ReturnPeriod = mvarSL1ReturnPeriod
End Property

Public Property Let SL2ReturnPeriod(vData As Integer)
10        mvarSL2ReturnPeriod = vData
End Property

Public Property Get SL2ReturnPeriod() As Integer
10        SL2ReturnPeriod = mvarSL2ReturnPeriod
End Property

Public Property Let Phone1ReturnPeriod(vData As Integer)
10        mvarPhone1ReturnPeriod = vData
End Property

Public Property Get Phone1ReturnPeriod() As Integer
10        Phone1ReturnPeriod = mvarPhone1ReturnPeriod
End Property

Public Property Let Phone2ReturnPeriod(vData As Integer)
10        mvarPhone2ReturnPeriod = vData
End Property

Public Property Get Phone2ReturnPeriod() As Integer
10        Phone2ReturnPeriod = mvarPhone2ReturnPeriod
End Property

Public Property Let LossFeeReturnPeriod(vData As Integer)
10        mvarLossFeeReturnPeriod = vData
End Property

Public Property Get LossFeeReturnPeriod() As Integer
10        LossFeeReturnPeriod = mvarLossFeeReturnPeriod
End Property

Public Sub CheckPoint()
10        With mrstCheckPoint
20            !Business_Type = mvarName
30            !InitialReturnPeriod = IIf(mvarInitialReturnPeriod = 0, 1, mvarInitialReturnPeriod)
40            !SL1ReturnPeriod = IIf(mvarSL1ReturnPeriod = 0, 1, mvarSL1ReturnPeriod)
50            !SL2ReturnPeriod = IIf(mvarSL2ReturnPeriod = 0, 1, mvarSL2ReturnPeriod)
60            !Phone1ReturnPeriod = IIf(mvarPhone1ReturnPeriod = 0, 1, mvarPhone1ReturnPeriod)
70            !Phone2ReturnPeriod = IIf(mvarPhone2ReturnPeriod = 0, 1, mvarPhone2ReturnPeriod)
80            !LossFee = IIf(mvarLossFeeReturnPeriod = 0, 1, mvarLossFeeReturnPeriod)
90        End With
End Sub

Private Sub Class_Initialize()
          Dim rstSchema As ADODB.Recordset
          Dim fldBT As ADODB.Field
    
10        Set mrstCheckPoint = New ADODB.Recordset
20        Set rstSchema = gdbADO.OpenSchema(adSchemaColumns, Array(Empty, Empty, "BUSINESS_TYPE", Empty))
30        While Not rstSchema.EOF
40            With mrstCheckPoint
50                If rstSchema!DATA_TYPE = adChar Or rstSchema!DATA_TYPE = adWChar Then
60                    .Fields.Append rstSchema!COLUMN_NAME, rstSchema!DATA_TYPE, -1 ' rstSchema!CHARACTER_MAXIMUM_LENGTH
70                Else
80                    .Fields.Append rstSchema!COLUMN_NAME, rstSchema!DATA_TYPE
90                End If
100               rstSchema.MoveNext
110           End With
120       Wend
130       With mrstCheckPoint
140           .Open
150           .AddNew
160           !Business_Type = ""
170           !InitialReturnPeriod = 0
180           !SL1ReturnPeriod = 0
190           !SL2ReturnPeriod = 0
200           !Phone1ReturnPeriod = 0
210           !Phone2ReturnPeriod = 0
220           !LossFee = 0
230           .Update
240       End With
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
60                    Err.Raise vbObjectError + 1, "BusinessType: CheckRecordSame", "BusinessType has been updated since first obtaining it."
70                End If
80            Next fldCheckpoint
90        End With
100       CheckRecordSame = True
End Function

Private Function Refresh(rstRecord As ADODB.Recordset)
10        With rstRecord
20            mvarName = !Business_Type
30            mvarInitialReturnPeriod = IIf(NVL(!InitialReturnPeriod, NVLNumeric) = 0, 30, NVL(!InitialReturnPeriod, NVLNumeric))
40            mvarSL1ReturnPeriod = IIf(NVL(!SL1ReturnPeriod, NVLNumeric) = 0, 30, NVL(!SL1ReturnPeriod, NVLNumeric))
50            mvarSL2ReturnPeriod = IIf(NVL(!SL2ReturnPeriod, NVLNumeric) = 0, 30, NVL(!SL2ReturnPeriod, NVLNumeric))
60            mvarPhone1ReturnPeriod = IIf(NVL(!Phone1ReturnPeriod, NVLNumeric) = 0, 30, NVL(!Phone1ReturnPeriod, NVLNumeric))
70            mvarPhone2ReturnPeriod = IIf(NVL(!Phone2ReturnPeriod, NVLNumeric) = 0, 30, NVL(!Phone2ReturnPeriod, NVLNumeric))
80            mvarLossFeeReturnPeriod = IIf(NVL(!LossFee, NVLNumeric) = 0, 30, NVL(!LossFee, NVLNumeric))
90        End With
100       CheckPoint
End Function

Public Sub Read(Name As String)
    
          Dim rstBusinessType As ADODB.Recordset
          Dim strSQL As String
    
10        On Error GoTo ErrorProc
    
20        mvarName = Name
    
30        strSQL = "SELECT * " _
                 & "FROM BUSINESS_TYPE " _
                 & "WHERE BUSINESS_TYPE = '" & Replace(mvarName, "'", "''") & "'"
    
40        Set rstBusinessType = New ADODB.Recordset
50        rstBusinessType.Open strSQL, gdbADO, , , adCmdText
60        Refresh rstBusinessType
70        rstBusinessType.Close
80        Set rstBusinessType = Nothing
  
90        Exit Sub

ErrorProc:
100       Err.Raise Err.Number, "BusinessType: Read", Err.Description
End Sub

Public Sub Update()
          Dim rstBusinessType As ADODB.Recordset
          Dim strSQL As String
    
10        On Error GoTo ErrorProc
    
20        strSQL = "SELECT * FROM BUSINESS_TYPE " _
                 & "WHERE BUSINESS_TYPE = '" & Replace(mvarName, "'", "''") & "'"
30        Set rstBusinessType = New ADODB.Recordset
40        rstBusinessType.Open strSQL, gdbADO, adOpenKeyset, adLockPessimistic, adCmdText
  
50        If rstBusinessType.RecordCount <> 1 Then
60            Err.Raise vbObjectError + 2, , "Business Type does not Exist"
70        End If
80        With rstBusinessType
90            CheckRecordSame rstBusinessType
100           !InitialReturnPeriod = mvarInitialReturnPeriod
110           !SL1ReturnPeriod = mvarSL1ReturnPeriod
120           !SL2ReturnPeriod = mvarSL2ReturnPeriod
130           !Phone1ReturnPeriod = mvarPhone1ReturnPeriod
140           !Phone2ReturnPeriod = mvarPhone2ReturnPeriod
150           !LossFee = mvarLossFeeReturnPeriod
160           .Update
170       End With
180       CheckPoint
    
190       Exit Sub

ErrorProc:
200       Err.Raise Err.Number, "BusinessType: Update"
End Sub

Public Sub Delete()
          Dim rstBusinessType As ADODB.Recordset
          Dim strSQL As String
          Dim lngRecordsAffected As Long
    
10        On Error GoTo ErrorProc
    
20        strSQL = "DELETE * FROM BUSINESS_TYPE " _
                 & "WHERE BUSINESS_TYPE = '" & Replace(mvarName, "'", "''") & "'"
30        On Error Resume Next
40        gdbADO.Execute strSQL, lngRecordsAffected, adCmdText
50        If Err.Number = -2147467259 Then
60            On Error GoTo 0
70            Err.Raise vbObjectError + 4, "BusinessType: Delete", "This Business Type is used by one or more customers and cannot be deleted."
80        Else
90            If Err.Number = 0 Then
100               On Error GoTo ErrorProc
110           Else
120               GoTo ErrorProc
130           End If
140       End If
150       If lngRecordsAffected <> 1 Then
160           Err.Raise vbObjectError + 3, "BusinessType: Delete", "Business Type does not Exist"
170       End If
    
180       Exit Sub

ErrorProc:
190       Err.Raise Err.Number, "BusinessType: Delete"
End Sub

Public Sub create(Name As String, _
                  InitialPeriod As Integer, _
                  SL1Period As Integer, _
                  SL2Period As Integer, _
                  Phone1Period As Integer, _
                  Phone2Period As Integer, _
                  LossFeePeriod As Integer)

          Dim strSQL As String
    
10        On Error GoTo ErrorProc
    
20        strSQL = "INSERT INTO BUSINESS_TYPE (" _
                 & "Business_Type, InitialReturnPeriod, " _
                 & "SL1ReturnPeriod, SL2ReturnPeriod, " _
                 & "Phone1ReturnPeriod, Phone2ReturnPeriod, " _
                 & "LossFee) VALUES (" _
                 & "'" & Replace(Name, "'", "''") & "'," _
                 & CStr(InitialPeriod) & "," _
                 & CStr(SL1Period) & "," _
                 & CStr(SL2Period) & "," _
                 & CStr(Phone1Period) & "," _
                 & CStr(Phone2Period) & "," _
                 & CStr(LossFeePeriod) & ")"
     
30        gdbADO.Execute strSQL, , adCmdText
    
40        mvarName = Name
50        mvarInitialReturnPeriod = InitialPeriod
60        mvarSL1ReturnPeriod = SL1Period
70        mvarSL2ReturnPeriod = SL2Period
80        mvarPhone1ReturnPeriod = Phone1Period
90        mvarPhone2ReturnPeriod = Phone2Period
100       mvarLossFeeReturnPeriod = LossFeePeriod
110       Exit Sub

ErrorProc:
120       Err.Raise Err.Number, "BusinessType: Create"
End Sub

