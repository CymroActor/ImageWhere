VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "KeywordSynonym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mvarID                      As Long
Private mvarKeyword                 As String
Private mvarSynonym                 As String
Private mvarPendingSynonym          As String
Private mvarDirty                   As enmPending

Public Property Let ID(ByVal vData As Long)
    mvarID = vData
End Property

Public Property Get ID() As Long
    ID = mvarID
End Property

Public Property Let Keyword(ByVal vData As String)
    mvarKeyword = vData
    mvarDirty = pdgUpdate
End Property

Public Property Get Keyword() As String
    Keyword = mvarKeyword
End Property

Public Property Let Synonym(ByVal vData As String)
    mvarSynonym = vData
    mvarDirty = pdgUpdate
End Property

Public Property Get Synonym() As String
    Synonym = mvarSynonym
End Property

Public Property Let PendingSynonym(ByVal vData As String)
    mvarPendingSynonym = vData
    mvarDirty = pdgUpdate
End Property

Public Property Get PendingSynonym() As String
    PendingSynonym = mvarPendingSynonym
End Property

Public Property Get Dirty() As enmPending
    Dirty = mvarDirty
End Property

Public Property Let Dirty(ByVal vData As enmPending)
    mvarDirty = vData
End Property

Public Sub Create()
    Dim strSQL                  As String
    Dim rstKeywordSynonym      As DAO.Recordset

'***************************************
' Module/Form Name   : KeywordSynonym
'
' Procedure Name     : Create
'
' Purpose            :
'
' Date Created       : 19/07/2005
'
' Author             : Gareth Saunders
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                      19/07/2005 Gareth Saunders
'
'***************************************
'
    On Error GoTo ErrorHandler:
'
'******** Code Starts Here *************
'
    strSQL = "INSERT INTO KeywordSynonyms (KEYWORD, SYNONYM) VALUES ("
    strSQL = strSQL & "'" & mvarKeyword & "'"
    strSQL = strSQL & ",'" & mvarSynonym & "'"
    strSQL = strSQL & ")"

    db.Execute strSQL, dbFailOnError
'
'********* Code Ends Here **************
'
    On Error GoTo 0
    Exit Sub
'
ErrorHandler:
    ErrorRaise "KeywordSynonym.Create"
End Sub

Public Sub Update()
    Dim strSQL              As String
    Dim rstKeywordSynonym   As DAO.Recordset
    
'***************************************
' Module/Form Name   : KeywordSynonym
'
' Procedure Name     : Update
'
' Purpose            :
'
' Date Created       : 19/07/2005
'
' Author             : Gareth Saunders
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                      19/07/2005 Gareth Saunders
'
'***************************************
'
    On Error GoTo ErrorHandler:
'
'******** Code Starts Here *************
'
    strSQL = "UPDATE KeywordSynonyms SET "
    strSQL = strSQL & "SYNONYM = '" & mvarPendingSynonym & "' "
    strSQL = strSQL & "WHERE KEYWORD = '" & mvarKeyword & "' "
    strSQL = strSQL & "AND SYNONYM = '" & mvarSynonym & "'"

    db.Execute strSQL, dbFailOnError
'
'********* Code Ends Here **************
'
    On Error GoTo 0
    Exit Sub
'
ErrorHandler:
    ErrorRaise "KeywordSynonym.Update"
End Sub

Public Sub Delete()
    Dim strSQL                  As String
    Dim rstKeywordSynonymns     As DAO.Recordset
    
'***************************************
' Module/Form Name   : KeywordSynonym
'
' Procedure Name     : Delete
'
' Purpose            :
'
' Date Created       : 19/07/2005
'
' Author             : Gareth Saunders
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                      19/07/2005 Gareth Saunders
'
'***************************************
'
    On Error GoTo ErrorHandler:
'
'******** Code Starts Here *************
'
    strSQL = "DELETE FROM KeywordSynonyms "
    strSQL = strSQL & "WHERE KEYWORD = '" & mvarKeyword & "' AND "
    strSQL = strSQL & "SYNONYM = '" & mvarSynonym & "'"

    db.Execute strSQL, dbFailOnError
'
'********* Code Ends Here **************
'
    On Error GoTo 0
    Exit Sub
'
ErrorHandler:
    ErrorRaise "KeywordSynonym.Delete"
End Sub

