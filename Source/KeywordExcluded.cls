VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "KeywordExcluded"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mvarKeyword                 As String
Private mvarDirty                   As enmPending

Public Property Let Keyword(ByVal vData As String)
    mvarKeyword = vData
    mvarDirty = pdgUpdate
End Property

Public Property Get Keyword() As String
    Keyword = mvarKeyword
End Property

Public Property Get Dirty() As enmPending
    Dirty = mvarDirty
End Property

Public Property Let Dirty(ByVal vData As enmPending)
    mvarDirty = vData
End Property

Public Sub Create()
    Dim strSQL                  As String
    Dim rstKeywordExcluded      As DAO.Recordset

    strSQL = "INSERT INTO KeywordsExcluded (KEYWORD) VALUES ("
    strSQL = strSQL & "'" & mvarKeyword & "'"
    strSQL = strSQL & ")"

    db.Execute strSQL, dbFailOnError
End Sub

Public Sub Delete()
    Dim strSQL                  As String
    Dim rstKeywordExcluded      As DAO.Recordset
    
    strSQL = "DELETE FROM KeywordsExcluded "
    strSQL = strSQL & "WHERE KEYWORD = '" & mvarKeyword & "'"

    db.Execute strSQL, dbFailOnError
End Sub


