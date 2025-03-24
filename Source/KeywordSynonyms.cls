VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "KeywordSynonyms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mCol As Collection

Public Sub Refresh()
    Dim strSQL As String
    Dim recKeywordSynonyms As DAO.Recordset
    Dim objKeywordSynonym  As KeywordSynonym
    '
    '   Clear all existing Keyword Synonyms.
    '
    Set mCol = Nothing
    Set mCol = New Collection
    '
    '   Refresh Keyword Synonyms Collection.
    '
    strSQL = "SELECT * FROM KEYWORDSYNONYMS ORDER BY KEYWORD"
    Set recKeywordSynonyms = db.OpenRecordset(strSQL, dbOpenSnapshot, dbForwardOnly)
    With recKeywordSynonyms
        While Not .EOF
            Set objKeywordSynonym = Add(!Keyword, _
                                        !Synonym)
            objKeywordSynonym.Dirty = pdgClean
            .MoveNext
        Wend
    End With

    recKeywordSynonyms.Close
    Set recKeywordSynonyms = Nothing
End Sub

Public Function Add(ByVal Keyword As String, _
                    ByVal Synonym As String) As KeywordSynonym
    'create a new object
    Dim objNewMember As KeywordSynonym
'***************************************
' Module/Form Name   : KeywordSynonyms
'
' Procedure Name     : Add
'
' Purpose            :
'
' Date Created       : 25/07/2005
'
' Author             : Gareth Saunders
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                      25/07/2005 Gareth Saunders
'
'***************************************
'
    On Error GoTo ErrorHandler:
'
'******** Code Starts Here *************
'
    Set objNewMember = New KeywordSynonym
    
    'set the properties passed into the method
    objNewMember.Keyword = Keyword
    objNewMember.Synonym = Synonym
    objNewMember.Dirty = pdgCreate
    mCol.Add objNewMember, Keyword & "/" & Synonym
    Set Add = objNewMember
    Set objNewMember = Nothing
'
'********* Code Ends Here **************
'
    On Error GoTo 0
    Exit Function
'
ErrorHandler:
    ErrorRaise "KeywordSynonyms.Add"
End Function

Public Property Get Item(vntIndexKey As Variant) As KeywordSynonym
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
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

Public Sub Update()
    
    Dim oKeywordSynonym As KeywordSynonym
    Dim intIndex As Integer
    
    For intIndex = 1 To mCol.Count
        Set oKeywordSynonym = mCol.Item(intIndex)
        Select Case oKeywordSynonym.Dirty
        Case Is = pdgCreate
            oKeywordSynonym.Create
        Case Is = pdgUpdate
            oKeywordSynonym.Update
        Case Is = pdgDelete
            oKeywordSynonym.Delete
        End Select
    Next intIndex
    '
    '   Now refresh to pick up any other updates made by anyone else.
    '
    Refresh
End Sub

Public Sub RemovePending(vntIndexKey As Variant)
    mCol.Item(vntIndexKey).Dirty = pdgDelete
End Sub
