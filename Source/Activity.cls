VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Activity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mvarId                  As String
Private mvarActivityType        As String
Private mvarCustomerNo          As Integer
Private mvarJobNo               As Long
Private mvarSearchNo            As Long
Private mvarDescription         As String
Private mvarStatus              As Integer
Private mvarStartDate           As Date
Private mvarEndDate             As Date
Private mvarUserField1          As String
Private mvarUserField2          As String
'
Private Type udtActivity
    ID                  As String
    ActivityType        As String
    CustomerNo          As Integer
    JobNo               As Long
    SearchNo            As Long
    Description         As String
    Status              As Integer
    StartDate           As Date
    EndDate             As Date
    UserField1          As String
    UserField2          As String
End Type
Private CheckPointActivity As udtActivity

Public Property Let ID(vData As String)
    mvarId = vData
End Property

Public Property Get ID() As String
    ID = mvarId
End Property

Public Property Let ActivityType(vData As String)
    mvarActivityType = vData
End Property

Public Property Get ActivityType() As String
    ActivityType = mvarActivityType
End Property

Public Property Let CustomerNo(vData As Integer)
    mvarCustomerNo = vData
End Property

Public Property Get CustomerNo() As Integer
    CustomerNo = mvarCustomerNo
End Property

Public Property Let JobNo(vData As Long)
    mvarJobNo = vData
End Property

Public Property Get JobNo() As Long
    JobNo = mvarJobNo
End Property

Public Property Let SearchNo(vData As Long)
    mvarSearchNo = vData
End Property

Public Property Get SearchNo() As Long
    SearchNo = mvarSearchNo
End Property

Public Property Let Description(vData As String)
    mvarDescription = vData
End Property

Public Property Get Description() As String
    Description = mvarDescription
End Property

Public Property Let Status(vData As Integer)
    mvarStatus = vData
End Property

Public Property Get Status() As Integer
    Status = mvarStatus
End Property

Public Property Let StartDate(vData As Date)
    mvarStartDate = vData
End Property

Public Property Get StartDate() As Date
    StartDate = mvarStartDate
End Property

Public Property Let EndDate(vData As Date)
    mvarEndDate = vData
End Property

Public Property Get EndDate() As Date
    EndDate = mvarEndDate
End Property

Public Property Let UserField1(vData As String)
    mvarUserField1 = vData
End Property

Public Property Get UserField1() As String
    UserField1 = mvarUserField1
End Property

Public Property Let UserField2(vData As String)
    mvarUserField2 = vData
End Property

Public Property Get UserField2() As String
    UserField2 = mvarUserField2
End Property

Public Sub CheckPoint()
'***************************************
' Module/Form Name   : Activity
'
' Procedure Name     : CheckPoint
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
On Error GoTo Checkpoint_Error
'
'******** Code Starts Here *************
'
    With CheckPointActivity
        .ID = mvarId
        .ActivityType = mvarActivityType
        .CustomerNo = mvarCustomerNo
        .JobNo = mvarJobNo
        .SearchNo = mvarSearchNo
        .Description = mvarDescription
        .Status = mvarStatus
        .StartDate = mvarStartDate
        .EndDate = mvarEndDate
        .UserField1 = mvarUserField1
        .UserField2 = mvarUserField2
    End With
'
'********* Code Ends Here **************
'
   Exit Sub
'
Checkpoint_Error:
    ErrorRaise "Activity.CheckPoint"
End Sub

Private Sub Class_Initialize()
    With CheckPointActivity
        .ID = 0
        .ActivityType = ""
        .CustomerNo = 0
        .JobNo = 0
        .SearchNo = 0
        .Description = ""
        .Status = 0
        .StartDate = 0
        .EndDate = 0
        .UserField1 = ""
        .UserField2 = ""
    End With
End Sub

''Private Function CheckRecordSame(rstRecord As ADODB.Recordset) As Boolean
Private Function CheckRecordSame(ByRef rstRecord As DAO.Recordset) As Boolean
    CheckRecordSame = True
    '
    '   Check if record has changed by comparing it with the data
    '   at last refresh.
    '
    With rstRecord
        If NVL(!ActivityType, NVLString) <> CheckPointActivity.ActivityType Or _
           NVL(!CustomerNo, NVLNumeric) <> CheckPointActivity.CustomerNo Or _
           NVL(!JobNo, NVLNumeric) <> CheckPointActivity.JobNo Or _
           NVL(!SearchNo, NVLNumeric) <> CheckPointActivity.SearchNo Or _
           NVL(!Description, NVLString) <> CheckPointActivity.Description Or _
           NVL(!Status, NVLNumeric) <> CheckPointActivity.Status Or _
           NVL(!StartDate, NVLNumeric) <> CheckPointActivity.StartDate Or _
           NVL(!EndDate, NVLNumeric) <> CheckPointActivity.EndDate Or _
           NVL(!UserField1, NVLString) <> CheckPointActivity.UserField1 Or _
           NVL(!UserField2, NVLString) <> CheckPointActivity.UserField2 Then
            Refresh rstRecord
            CheckPoint
            Err.Raise vbObjectError + 5, "Location: CheckRecordSame", "Location has been updated by someone else since first obtaining it."
        End If
    End With
    CheckRecordSame = False
End Function

''Private Function Refresh(rstRecord As ADODB.Recordset)
Private Function Refresh(rstRecord As DAO.Recordset)
    With rstRecord
        mvarId = !ID
        mvarActivityType = !ActivityType
        mvarCustomerNo = !CustomerNo
        mvarJobNo = !JobNo
        mvarSearchNo = !SearchNo
        mvarDescription = !Description
        mvarStatus = !Status
        mvarStartDate = !StartDate
        mvarEndDate = !EndDate
        mvarUserField1 = !UserField1
        mvarUserField2 = !UserField2
    End With
    CheckPoint
End Function

Public Sub Read(Optional ByVal plngID As Long = 0)
'***************************************
' Module/Form Name   : Activity
'
' Procedure Name     : Read
'
' Purpose            :
'
' Date Created       : 09/07/2004
'
' Author             : GARETH SAUNDERS
'
' Parameters         : plngID - Long
'                    : = -
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
    
''    Dim rstActivity As ADODB.Recordset
    Dim rstActivity As DAO.Recordset
    Dim strSQL          As String
    
    strSQL = "SELECT * "
    strSQL = strSQL & "FROM ACTIVITIES "
    strSQL = strSQL & "WHERE ID = " & CStr(plngID)
    
''    Set rstActivity = New ADODB.Recordset
''    rstActivity.Open strSQL, gdbADO, , , adCmdText
    Set rstActivity = db.OpenRecordset(strSQL, dbOpenSnapshot)
    Refresh rstActivity
    rstActivity.Close
    Set rstActivity = Nothing
  
'
'********* Code Ends Here **************
'
   Exit Sub
'
Read_Error:
    ErrorRaise "Activity.Read"
End Sub

Public Sub update()
'***************************************
' Module/Form Name   : Activity
'
' Procedure Name     : update
'
' Purpose            :
'
' Date Created       : 09/07/2004
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo update_Error
'
'******** Code Starts Here *************
'
''    Dim rstActivity     As ADODB.Recordset
    Dim rstActivity     As DAO.Recordset
    Dim strSQL          As String
    
    strSQL = "SELECT * FROM ACTIVITIES "
    strSQL = strSQL & "WHERE ID = " & CStr(mvarId)
    
''    Set rstActivity = New ADODB.Recordset
''    rstActivity.Open strSQL, gdbADO, adOpenKeyset, adLockPessimistic, adCmdText
    Set rstActivity = db.OpenRecordset(strSQL, dbOpenDynaset)
    
    If rstActivity.RecordCount <> 1 Then
        Err.Raise vbObjectError + 2, , "Activity does not Exist"
    End If
    With rstActivity
        CheckRecordSame rstActivity
        .Edit
        !ActivityType = mvarActivityType
        !CustomerNo = mvarCustomerNo
        !JobNo = mvarJobNo
        !SearchNo = mvarSearchNo
        !Description = mvarDescription
        !Status = mvarStatus
        !StartDate = mvarStartDate
        !EndDate = mvarEndDate
        !UserField1 = mvarUserField1
        !UserField2 = mvarUserField2
        .update
    End With
    CheckPoint
'
'********* Code Ends Here **************
'
   Exit Sub
'
update_Error:
    ErrorRaise "Activity.update"
End Sub

Public Sub Create()
'***************************************
' Module/Form Name   : Activity
'
' Procedure Name     : create
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

    Dim strSQL          As String
''    Dim rstActivity     As ADODB.Recordset
    Dim rstActivity     As DAO.Recordset
    '
    '   This technique is required to retrieve the Job Number afterwards.
    '
''    Set rstActivity = New ADODB.Recordset
''    rstActivity.Open "Activities", gdbADO, adOpenKeyset, adLockPessimistic
    Set rstActivity = db.OpenRecordset("Activities", dbOpenTable)
    
    With rstActivity
        .AddNew
        !ActivityType = mvarActivityType
        !CustomerNo = mvarCustomerNo
        !JobNo = mvarJobNo
        !SearchNo = IIf(mvarSearchNo = 0, Null, mvarSearchNo)
        !Description = mvarDescription
        !Status = mvarStatus
        !StartDate = mvarStartDate
        !EndDate = mvarEndDate
        !UserField1 = mvarUserField1
        !UserField2 = mvarUserField2
        !Machine = goSystemConfig.ComputerName
        !User = goSystemConfig.UserName
        .update
    End With
    '
    '   Set up the internal values.
    '
    mvarId = rstActivity!ID
    rstActivity.Close
    Set rstActivity = Nothing
    
'
'********* Code Ends Here **************
'
   Exit Sub
'
Create_Error:
    ErrorRaise "Activity.create"
End Sub

