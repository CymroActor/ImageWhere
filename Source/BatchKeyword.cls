VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BatchKeyword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mvarBatchNo                 As Long
Private mvarKeyword                 As String
Private mvarKeywordType             As String
Private mvarDirty                   As enmPending
    
Public Property Get BatchNo() As Long
    BatchNo = mvarBatchNo
End Property

Public Property Let BatchNo(ByVal vData As Long)
    mvarBatchNo = vData
End Property

Public Property Get Keyword() As String
    Keyword = mvarKeyword
End Property

Public Property Let Keyword(ByVal vData As String)
    mvarKeyword = vData
    mvarDirty = pdgUpdate
End Property

Public Property Get KeywordType() As String
    KeywordType = mvarKeywordType
End Property

Public Property Let KeywordType(ByVal vData As String)
    If vData <> "A" And vData <> "M" And vData <> "S" Then
        Err.Raise vbObjectError + 1, , "KeywordType must be 'A', 'M' or 'S'"
    End If
    mvarKeywordType = vData
End Property

Public Property Get Dirty() As enmPending
    Dirty = mvarDirty
End Property

Public Property Let Dirty(ByVal vData As enmPending)
    mvarDirty = vData
End Property

Public Sub Create()
    Dim strSQL              As String
    Dim rstBatchKeyword     As DAO.Recordset
    
'***************************************
' Module/Form Name   : BatchKeyword
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

    strSQL = "INSERT INTO BATCHKEYWORDS (BATCHNO, KEYWORD, KEYWORDTYPE) VALUES ("
    strSQL = strSQL & "" & CStr(mvarBatchNo) & ""
    strSQL = strSQL & ",'" & mvarKeyword & "'"
    strSQL = strSQL & ",'" & mvarKeywordType & "'"
    strSQL = strSQL & ")"

    db.Execute strSQL, dbFailOnError
'
'********* Code Ends Here **************
'
   On Error GoTo 0
   Exit Sub
'
ErrorHandler:
    ErrorRaise "BatchKeyword.Create"
End Sub

Public Sub Delete()
    Dim strSQL              As String
    Dim rstBatchKeyword     As DAO.Recordset
    
'***************************************
' Module/Form Name   : BatchKeyword
'
' Procedure Name     : Delete
'
' Purpose            :
'
' Date Created       : 18/07/2005
'
' Author             : Gareth Saunders
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                      18/07/2005 Gareth Saunders
'
'***************************************
'
    On Error GoTo ErrorHandler:
'
'******** Code Starts Here *************
'

    strSQL = "DELETE FROM BATCHKEYWORDS "
    strSQL = strSQL & "WHERE BATCHNO = " & CStr(mvarBatchNo) & " AND "
    strSQL = strSQL & "KEYWORD = '" & mvarKeyword & "'"

    db.Execute strSQL, dbFailOnError
'
'********* Code Ends Here **************
'
    On Error GoTo 0
    Exit Sub
'
ErrorHandler:
    ErrorRaise "BatchKeyword.Delete"
End Sub

