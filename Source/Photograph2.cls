VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Photograph2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mvarPhotographKey   As Long
Private mvarBatchNo         As Long
Private mvarPhotographNo    As Long
Private mvarStatus          As String
Private mvarScanned         As Boolean
Private mvarDeleted         As Boolean

Public Property Let PhotographKey(ByVal vData As Long)
    mvarPhotographKey = vData
End Property

Public Property Get PhotographKey() As Long
    PhotographKey = mvarPhotographKey
End Property

Public Property Let BatchNo(ByVal vData As Long)
    mvarBatchNo = vData
End Property

Public Property Get BatchNo() As Long
    BatchNo = mvarBatchNo
End Property

Public Property Let PhotographNo(ByVal vData As Long)
    mvarPhotographNo = vData
End Property

Public Property Get PhotographNo() As Long
    PhotographNo = mvarPhotographNo
End Property

Public Property Let Status(ByVal vData As String)
    mvarStatus = vData
End Property

Public Property Get Status() As String
    Status = mvarStatus
End Property

Public Property Let scanned(ByVal vData As Boolean)
    mvarScanned = vData
End Property

Public Property Get scanned() As Boolean
    scanned = mvarScanned
End Property

Public Property Let Deleted(ByVal vData As Boolean)
    mvarDeleted = vData
End Property

Public Property Get Deleted() As Boolean
    Deleted = mvarDeleted
End Property

Public Sub Create()
'***************************************
' Module/Form Name   : Photograph2
'
' Procedure Name     : Create
'
' Purpose            :
'
' Date Created       : 24/06/2004
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo Create_Error
'
'******** Code Starts Here *************
'
''    Dim rstPhotograph   As ADODB.Recordset
    Dim rstPhotograph   As DAO.Recordset

''    Set rstPhotograph = New ADODB.Recordset
''    rstPhotograph.Open "Photograph", gdbADO, adOpenKeyset, adLockPessimistic
    Set rstPhotograph = db.OpenRecordset("Photograph", dbOpenTable)
    With rstPhotograph
        .AddNew
        !batch_no = mvarBatchNo
        !photograph_no = mvarPhotographNo
        !Status = mvarStatus
        .Update
        .Bookmark = .LastModified
        mvarPhotographKey = !photograph_key
    End With

    rstPhotograph.Close
    Set rstPhotograph = Nothing
'
'********* Code Ends Here **************
'
   Exit Sub
'
Create_Error:
    ErrorRaise "Photograph2.Create"
End Sub

Public Sub Read(ByVal plngBatchNo As Long, _
                ByVal plngPhotographNo As Long)
'***************************************
' Module/Form Name   : Photograph2
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
'                    : plngPhotographNo - Long
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
    Dim strSQL              As String
''    Dim rstPhotograph       As ADODB.Recordset
    Dim rstPhotograph       As DAO.Recordset
    
    strSQL = "SELECT * "
    strSQL = strSQL & "FROM PHOTOGRAPH "
    strSQL = strSQL & "WHERE BATCH_NO = " & CStr(plngBatchNo)
    strSQL = strSQL & " AND PHOTOGRAPH_NO = " & CStr(plngPhotographNo)
    
''    Set rstPhotograph = New ADODB.Recordset
''    rstPhotograph.Open strSQL, gdbADO, adOpenDynamic, adLockOptimistic, adCmdText
    Set rstPhotograph = db.OpenRecordset(strSQL, dbOpenSnapshot)
    With rstPhotograph
        If Not .EOF Then
            mvarBatchNo = NVL(!batch_no, NVLNumeric)
            mvarPhotographNo = NVL(!photograph_no, NVLNumeric)
            mvarPhotographKey = NVL(!photograph_key, NVLNumeric)
            mvarStatus = NVL(!Status, NVLString)
            mvarScanned = NVL(!scanned, NVLNumeric)
            mvarDeleted = NVL(!Deleted, NVLNumeric)
        End If
    End With
    rstPhotograph.Close
    Set rstPhotograph = Nothing
    
'
'********* Code Ends Here **************
'
   Exit Sub
'
Read_Error:
    ErrorRaise "Photograph2.Read"
End Sub


Public Sub Update()
'***************************************
' Module/Form Name   : Photograph2
'
' Procedure Name     : Update
'
' Purpose            :
'
' Date Created       : 24/06/2004
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo Update_Error
'
'******** Code Starts Here *************
'
    Dim strSQL              As String
   
    strSQL = "UPDATE PHOTOGRAPH SET "
    strSQL = strSQL & "STATUS = '" & mvarStatus & "', "
    strSQL = strSQL & "SCANNED = " & mvarScanned & ", "
    strSQL = strSQL & "DELETED = " & mvarDeleted & " "
    strSQL = strSQL & "WHERE BATCH_NO = " & CStr(mvarBatchNo)
    strSQL = strSQL & " AND PHOTOGRAPH_NO = " & CStr(mvarPhotographNo)
    
''    gdbADO.Execute strSQL
    db.Execute strSQL
'
'********* Code Ends Here **************
'
   Exit Sub
'
Update_Error:
    ErrorRaise "Photograph2.Update"
End Sub

Private Sub Class_Initialize()
    mvarPhotographKey = 0
    mvarBatchNo = 0
    mvarPhotographNo = 0
    mvarStatus = ""
    mvarScanned = False
    mvarDeleted = False
End Sub
