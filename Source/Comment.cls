VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Comment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"Comment"
Option Explicit
Public Event CommentDeleted()
Public Key As String

Private mvarCustomerNo As Integer 'local copy
Private mvarAuthor As String 'local copy
Private mvarComment As String 'local copy
Private mvarDateWritten As Date 'local copy
Private mvarDeleted As Boolean 'local copy

Private Type udtComment
    CustomerNo As Integer
    ID As Integer
    Author As String
    Comment    As String
    DateWritten As Date
    Deleted As Boolean
End Type

Private CheckPointComment As udtComment

Public Property Let Deleted(ByVal vData As Boolean)
10        mvarDeleted = vData
End Property

Public Property Get Deleted() As Boolean
10        Deleted = mvarDeleted
End Property

Public Property Let DateWritten(ByVal vData As Date)
10        mvarDateWritten = vData
End Property

Public Property Get DateWritten() As Date
10        DateWritten = mvarDateWritten
End Property

Public Property Let Comment(ByVal vData As String)
10        mvarComment = vData
End Property

Public Property Get Comment() As String
10        Comment = mvarComment
End Property

Public Property Let Author(ByVal vData As String)
10        mvarAuthor = vData
End Property

Public Property Get Author() As String
10        Author = mvarAuthor
End Property

Public Property Let CustomerNo(ByVal vData As Integer)
10        mvarCustomerNo = vData
End Property

Public Property Get CustomerNo() As Integer
10        CustomerNo = mvarCustomerNo
End Property

Public Sub update()
          Dim strSQL As String
          Dim recComment As ADODB.Recordset
    
10        On Error GoTo ErrorProc
          '
          '   Get current copy of data and lock it.
          '
20        strSQL = "SELECT * " & _
                   "FROM COMMENT " & _
                   "WHERE ID = " & CInt(Key)
    
30        Set recComment = New ADODB.Recordset
40        recComment.Open strSQL, gdbADO, adOpenDynamic, adLockOptimistic, adCmdText
          '
          '   Update Data.
          '
50        With recComment
60            CheckRecordSame recComment
70            !Author = mvarAuthor
80            !Comment = mvarComment
90            !DateWritten = mvarDateWritten
100           !Deleted = mvarDeleted
110           .update
120       End With

130       Set recComment = Nothing
140       CheckPoint
150       Exit Sub
ErrorProc:
160       Set recComment = Nothing
170       Err.Raise Err.Number, "Comment: Update"
End Sub

Public Sub CheckPoint()
10        With CheckPointComment
20            .Author = mvarAuthor
30            .CustomerNo = mvarCustomerNo
40            .Comment = mvarComment
50            .DateWritten = mvarDateWritten
60            .Deleted = mvarDeleted
70        End With
End Sub

Public Sub Delete()
    
          Dim recComment As ADODB.Recordset
          Dim strSQL As String
          Dim lngRecordsAffected As Long
    
10        On Error GoTo ErrorProc
          '
          '   Read comment and lock before deleting.
          '
20        strSQL = "SELECT * " _
              & "FROM COMMENT " _
              & "WHERE ID = " & Key
    
30        Set recComment = New ADODB.Recordset
40        recComment.Open strSQL, gdbADO, adOpenDynamic, adLockOptimistic, adCmdText
    
50        CheckRecordSame recComment
60        strSQL = "DELETE * FROM COMMENT " _
                 & "WHERE ID = " & Key
70        gdbADO.Execute strSQL, lngRecordsAffected, adCmdText
80        If lngRecordsAffected <> 1 Then
90            Err.Raise vbObjectError + 3, "Comment: Delete", "Comment does not Exist"
100       End If
    
110       Set recComment = Nothing
120       Exit Sub
ErrorProc:
130       Set recComment = Nothing
140       Err.Raise Err.Number, "Comment: Delete"
End Sub

Private Function CheckRecordSame(recRecord As ADODB.Recordset) As Boolean
10        CheckRecordSame = True
          '
          '   Check it record has changed by comparing it with the data
          '   at last refresh.
          '
20        With recRecord
30            If !Author <> CheckPointComment.Author Or _
                 !Comment <> CheckPointComment.Comment Or _
                 !DateWritten <> CheckPointComment.DateWritten Or _
                 !Deleted <> CheckPointComment.Deleted Then
40                Refresh recRecord
50                CheckPoint
60                Err.Raise vbObjectError + 5, "Comment: CheckRecordSame", "Comment has been updated by someone else since first obtaining it."
70            End If
80        End With
90        CheckRecordSame = False
End Function

Private Function Refresh(recRecord As ADODB.Recordset)
10        mvarAuthor = recRecord!Author
20        mvarCustomerNo = recRecord!CustomerNo
30        mvarComment = recRecord!Comment
40        mvarDateWritten = recRecord!DateWritten
50        mvarDeleted = recRecord!Deleted
End Function

