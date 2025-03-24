VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Chaser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mvarId As Long
Private mvarDeliveryNoteNo As Long
Private mvarContact As String
Private mvarAction As String
Private mvarUser As String
Private mvarChaserDate As Date
Private mvarNextAction As String
Private mvarComment As String
Private mvarDatePrinted As Date
Private mvarDateCreated As Date
Private mvarDateAmended As Date
'
'   ADO Recordset used for data access.
'
Private mrstChaser As ADODB.Recordset
'
'   ADO Recordset to hold a checkpointed copy of the data.
'
Private mrstCheckPoint As ADODB.Recordset

Public Property Let ID(vData As Long)
10        mvarId = vData
End Property

Public Property Get ID() As Long
10        ID = mvarId
End Property

Public Property Let DeliveryNoteNo(vData As Long)
10        mvarDeliveryNoteNo = vData
End Property

Public Property Get DeliveryNoteNo() As Long
10        DeliveryNoteNo = mvarDeliveryNoteNo
End Property

Public Property Let Contact(vData As String)
10        mvarContact = vData
End Property

Public Property Get Contact() As String
10        Contact = mvarContact
End Property

Public Property Let Action(vData As String)
10        mvarAction = vData
End Property

Public Property Get Action() As String
10        Action = mvarAction
End Property

Public Property Let User(vData As String)
10        mvarUser = vData
End Property

Public Property Get User() As String
10        User = mvarUser
End Property

Public Property Let ChaserDate(vData As Date)
10        mvarChaserDate = vData
End Property

Public Property Get ChaserDate() As Date
10        ChaserDate = mvarChaserDate
End Property

Public Property Let NextAction(vData As String)
10        mvarNextAction = vData
End Property

Public Property Get NextAction() As String
10        NextAction = mvarNextAction
End Property

Public Property Let Comment(vData As String)
10        mvarComment = vData
End Property

Public Property Get Comment() As String
10        Comment = mvarComment
End Property

Public Property Let DatePrinted(vData As Date)
10        mvarDatePrinted = vData
End Property

Public Property Get DatePrinted() As Date
10        DatePrinted = mvarDatePrinted
End Property

Public Property Let DateCreated(vData As Date)
10        mvarDateCreated = vData
End Property

Public Property Get DateCreated() As Date
10        DateCreated = mvarDateCreated
End Property

Public Property Let DateAmended(vData As Date)
10        mvarDateAmended = vData
End Property

Public Property Get DateAmended() As Date
10        DateAmended = mvarDateAmended
End Property

Public Sub CheckPoint()
10        With mrstCheckPoint
20            !ID = mvarId
30            !DeliveryNoteNo = mvarDeliveryNoteNo
40            !Contact = mvarContact
50            !Action = mvarAction
60            !User = mvarUser
70            !ReturnByDate = mvarChaserDate
80            !NextAction = mvarNextAction
90            !Comment = mvarComment
100           !DatePrinted = mvarDatePrinted
110           !DateCreated = mvarDateCreated
120           !DateAmended = mvarDateAmended
130       End With
End Sub

Private Sub Class_Initialize()
      '    Dim rstSchema As ADODB.Recordset
          Dim fldBT As ADODB.Field
    
10        Set mrstCheckPoint = New ADODB.Recordset
      '    Set rstSchema = gdbADO.OpenSchema(adSchemaColumns, Array(Empty, Empty, "BUSINESS_TYPE", Empty))
      '    While Not rstSchema.EOF
20        With mrstCheckPoint
30            .Fields.Append "Id", adInteger
40            .Fields.Append "DeliveryNoteNo", adInteger
50            .Fields.Append "Contact", adVarWChar, 50
60            .Fields.Append "Action", adVarWChar, 13
70            .Fields.Append "User", adVarWChar, 20
80            .Fields.Append "ReturnByDate", adDate
90            .Fields.Append "NextAction", adVarWChar, 13
100           .Fields.Append "Comment", adVarWChar, 100
110           .Fields.Append "DatePrinted", adDate
120           .Fields.Append "DateCreated", adDate
130           .Fields.Append "DateAmended", adDate
140       End With
150       With mrstCheckPoint
160           .Open
170           .AddNew
180           !ID = 0
190           !DeliveryNoteNo = 0
200           !Contact = ""
210           !Action = ""
220           !User = ""
230           !ReturnByDate = 0
240           !NextAction = ""
250           !Comment = ""
260           !DatePrinted = 0
270           !DateCreated = 0
280           !DateAmended = 0
290           .update
300       End With
310       mvarDatePrinted = 0
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
60                    Err.Raise vbObjectError + 1, "Chaser: CheckRecordSame", "Chaser has been updated since first obtaining it."
70                End If
80            Next fldCheckpoint
90        End With
100       CheckRecordSame = True
End Function

Private Function Refresh(rstRecord As ADODB.Recordset)
10        With rstRecord
20            mvarId = !ID
30            mvarDeliveryNoteNo = !DeliveryNoteNo
40            mvarContact = !Contact
50            mvarAction = !Action
60            mvarUser = !User
70            mvarChaserDate = !ReturnByDate
80            mvarNextAction = !NextAction
90            mvarComment = !Comment
100           mvarDatePrinted = !DatePrinted
110           mvarDateCreated = !DateCreated
120           mvarDateAmended = !DateAmended
130       End With
140       CheckPoint
End Function

Public Sub Read(ID As Long)
    
          Dim strSQL As String
    
10        On Error GoTo ErrorProc
    
20        strSQL = "SELECT * " _
                 & "FROM CHASER " _
                 & "WHERE ID = " & CStr(ID)
    
30        Set mrstChaser = New ADODB.Recordset
40        mrstChaser.Open strSQL, gdbADO, , , adCmdText
50        Refresh mrstChaser
60        mrstChaser.Close
70        Set mrstChaser = Nothing
  
80        Exit Sub

ErrorProc:
90        Err.Raise Err.Number, "Chaser: Read", Err.Description
End Sub

Public Sub update()
          Dim strSQL As String
    
10        On Error GoTo ErrorProc
    
20        strSQL = "SELECT * FROM CHASER " _
                 & "WHERE ID = " & CStr(mvarId)
30        Set mrstChaser = New ADODB.Recordset
40        mrstChaser.Open strSQL, gdbADO, adOpenKeyset, adLockPessimistic, adCmdText
  
50        If mrstChaser.RecordCount <> 1 Then
60            Err.Raise vbObjectError + 2, , "Chaser does not Exist"
70        End If
80        With mrstChaser
90            CheckRecordSame mrstChaser
100           !DeliveryNoteNo = mvarDeliveryNoteNo
110           !Contact = mvarContact
120           !Action = mvarAction
130           !User = mvarUser
140           !ReturnByDate = mvarChaserDate
150           !NextAction = mvarNextAction
160           !Comment = mvarComment
170           !DatePrinted = mvarDatePrinted
180           mvarDateAmended = CDate(Now)
190           !DateAmended = mvarDateAmended
200           .update
210           .Close
  
220       End With
230       CheckPoint
240       Set mrstChaser = Nothing
250       Exit Sub

ErrorProc:
260       Err.Raise Err.Number, "Chaser: Update"
End Sub

Public Sub Delete()
          Dim strSQL As String
          Dim lngRecordsAffected As Long
    
10        On Error GoTo ErrorProc
    
20        strSQL = "DELETE * FROM CHASER " _
                 & "WHERE ID = " & CStr(mvarId)
30        gdbADO.Execute strSQL, lngRecordsAffected, adCmdText
40        If lngRecordsAffected <> 1 Then
50            Err.Raise vbObjectError + 3, "Chaser: Delete", "Chaser does not Exist"
60        End If
    
70        Exit Sub

ErrorProc:
80        Err.Raise Err.Number, "Chaser: Delete"
End Sub

Public Sub create(ByVal DeliveryNoteNo As Long, _
                  ByVal Contact As String, _
                  ByVal Action As String, _
                  ByVal User As String, _
                  ByVal ChaserDate As Date, _
                  ByVal NextAction As String, _
                  ByVal Comment As String)

    Dim strSQL As String
    
    On Error GoTo ErrorProc
    
    strSQL = "SELECT TOP 1 * FROM CHASER WHERE DELIVERYNOTENO = " & CStr(DeliveryNoteNo)
    Set mrstChaser = New ADODB.Recordset
    mrstChaser.Open strSQL, gdbADO, adOpenDynamic, adLockOptimistic, adCmdText
'    strSQL = "INSERT INTO CHASER (" _
'           & "DeliveryNoteNo, Contact, " _
'           & "Action, User, " _
'           & "ChaserDate, NextAction, " _
'           & "Coment) VALUES (" _
'           & CStr(DeliveryNoteNo) & ",'" _
'           & Contact & "','" _
'           & Action & "','" _
'           & User & "',#" _
'           & Format(ChaserDate, "MM/DD/YYYY") & "#,'" _
'           & NextAction & "','" _
'           & Comment & "')"
'    gdbADO.Execute strSQL, , adCmdText
    
    mvarId = mrstChaser!ID
    mvarDeliveryNoteNo = DeliveryNoteNo
    mvarContact = Contact
    mvarAction = Action
    mvarUser = User
    mvarChaserDate = ChaserDate
    If mvarAction = "Invoiced" Then
        mvarNextAction = ""
    Else
        mvarNextAction = NextAction
    End If
    mvarComment = Comment
    
    
    With mrstChaser
        .AddNew
        !DeliveryNoteNo = mvarDeliveryNoteNo
        !Contact = mvarContact
        !Action = mvarAction
        !User = mvarUser
        !ReturnByDate = mvarChaserDate
        !NextAction = mvarNextAction
        !Comment = mvarComment
        !DatePrinted = mvarDatePrinted
        !DateCreated = CDate(Now)
        !DateAmended = CDate(Now)
        .update
    End With
    '
    mrstChaser.Close
    Set mrstChaser = Nothing
    
    Exit Sub

ErrorProc:
    ErrorRaise "Chaser: Create"
End Sub

Public Sub CreateDAO(DeliveryNoteNo As Long, _
                     Contact As String, _
                     Action As String, _
                     User As String, _
                     ChaserDate As Date, _
                     NextAction As String, _
                     Comment As String)

          Dim strSQL As String
          Dim rstChaserDAO As DAO.Recordset
    
10        On Error GoTo ErrorProc
    
20        Set rstChaserDAO = db.OpenRecordset("Chaser", dbOpenTable)
30        With rstChaserDAO
40            .AddNew
50            !DeliveryNoteNo = DeliveryNoteNo
60            !Contact = Contact
70            !Action = Action
80            !User = User
90            !ReturnByDate = ChaserDate
100           !NextAction = NextAction
110           !Comment = Comment
120           !DatePrinted = mvarDatePrinted
130           !DateCreated = CDate(Now)
140           !DateAmended = CDate(Now)
150           .update
160           .Bookmark = .LastModified
170       End With
          '
180       mvarId = rstChaserDAO!ID
190       mvarDeliveryNoteNo = DeliveryNoteNo
200       mvarContact = Contact
210       mvarAction = Action
220       mvarUser = User
230       mvarChaserDate = ChaserDate
240       mvarNextAction = NextAction
250       mvarComment = Comment
260       rstChaserDAO.Close
270       Set rstChaserDAO = Nothing
280       Exit Sub

ErrorProc:
290       ErrorRaise "Chaser: Create"
End Sub


