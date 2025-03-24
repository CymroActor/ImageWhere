VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Comments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"Comment"
Attribute VB_Ext_KEY = "Member0" ,"Comment"
Option Explicit

'local variable to hold collection
Private mCol As Collection
Private mvarCustomerNo As Integer

Public Sub Refresh()
    Dim strSQL As String
    Dim recComments As ADODB.Recordset
    Dim objComment As Comment
    
    On Error GoTo ErrorProc
    '
    '   Clear all existing Comments.
    '
    While Not mCol.Count = 0
        mCol.Remove 1
    Wend
    '
    '   Refresh Comments.
    '
    strSQL = "SELECT * " _
           & "FROM COMMENT " _
           & "WHERE CUSTOMERNO = " & mvarCustomerNo & " " _
           & "ORDER BY DATEWRITTEN DESC"
    Set recComments = New ADODB.Recordset
    recComments.Open strSQL, gdbADO, adOpenStatic, , adCmdText
    With recComments
        While Not recComments.EOF
            Set objComment = Add(CStr(!ID), _
                                 !CustomerNo, _
                                 !DateWritten, _
                                 !Author, _
                                 !Comment, _
                                 !Deleted)
            objComment.CheckPoint
            .MoveNext
        Wend
    End With
    
    recComments.Close
    Set recComments = Nothing
    Exit Sub
ErrorProc:
    Err.Raise Err.Number, "Comments: Refresh"
End Sub
Private Function Add(sKey As String, _
                    CustomerNo As Integer, _
                    DateWritten As Date, _
                    Author As String, _
                    Comment As String, _
                    Deleted As Boolean) As Comment
        
    'create a new object
    Dim objNewMember As Comment
    Set objNewMember = New Comment

    'set the properties passed into the method
    objNewMember.Author = Author
    objNewMember.CustomerNo = CustomerNo
    objNewMember.Comment = Comment
    objNewMember.DateWritten = DateWritten
    objNewMember.Deleted = Deleted
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        objNewMember.Key = sKey
        mCol.Add objNewMember, sKey
    End If
    
    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As Comment
Attribute Item.VB_UserMemId = 0
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

Public Property Get CustomerNo() As Integer
    CustomerNo = mvarCustomerNo
End Property

Public Property Let CustomerNo(ByVal vNewValue As Integer)
    mvarCustomerNo = vNewValue
    Refresh
End Property

Public Function create(Author As String, _
                       Comment As String, _
                       DateWritten As Date, _
                       Deleted As Boolean) As Integer

    Dim strSQL As String
    Dim recComment As ADODB.Recordset
    
    On Error GoTo ErrorProc
    
    strSQL = "SELECT * FROM COMMENT WHERE ID = 0"
    Set recComment = New ADODB.Recordset
    recComment.Open strSQL, gdbADO, adOpenDynamic, adLockOptimistic, adCmdText
    
    With recComment
        .AddNew
            !CustomerNo = mvarCustomerNo
            !Author = Author
            !Comment = Comment
            !DateWritten = DateWritten
            !Deleted = Deleted
        .update
    End With
  
    create = recComment!ID
    
    Refresh
    Set recComment = Nothing
    Exit Function

ErrorProc:
    Err.Raise Err.Number, "Comments: Create"
    Set recComment = Nothing
End Function
