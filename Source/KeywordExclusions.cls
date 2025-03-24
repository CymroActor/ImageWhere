VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "KeywordExclusions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mCol As Collection

Public Sub Refresh()
    Dim strSQL As String
    Dim recKeywordExclusions As DAO.Recordset
    Dim objKeywordExcluded   As KeywordExcluded
'***************************************
' Module/Form Name   : KeywordExclusions
'
' Procedure Name     : Refresh
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
    '
    '   Clear all existing Keyword Exclusions.
    '
    Set mCol = Nothing
    Set mCol = New Collection
    '
    '   Refresh Keyword Synonyms Collection.
    '
    strSQL = "SELECT * FROM KeywordsExcluded ORDER BY KEYWORD"
    Set recKeywordExclusions = db.OpenRecordset(strSQL, dbOpenSnapshot, dbForwardOnly)
    With recKeywordExclusions
        While Not .EOF
            Set objKeywordExcluded = Add(!Keyword)
            objKeywordExcluded.Dirty = pdgClean
            .MoveNext
        Wend
    End With

    recKeywordExclusions.Close
    Set recKeywordExclusions = Nothing
'
'********* Code Ends Here **************
'
    On Error GoTo 0
    Exit Sub
'
ErrorHandler:
    ErrorRaise "KeywordExclusions.Refresh"
End Sub

Public Function Add(ByVal Keyword As String) As KeywordExcluded
    'create a new object
    Dim objNewMember As KeywordExcluded
'***************************************
' Module/Form Name   : KeywordExclusions
'
' Procedure Name     : Add
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
    Set objNewMember = New KeywordExcluded
    
    'set the properties passed into the method
    objNewMember.Keyword = Keyword
    mCol.Add objNewMember, Keyword
    objNewMember.Dirty = pdgCreate
    Set Add = objNewMember
    Set objNewMember = Nothing
'
'********* Code Ends Here **************
'
    On Error GoTo 0
    Exit Function
'
ErrorHandler:
    ErrorRaise "KeywordExclusions.Add"
End Function

Public Property Get Item(vntIndexKey As Variant) As KeywordExcluded
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

Public Sub RemovePending(vntIndexKey As Variant)
    mCol.Item(vntIndexKey).Dirty = pdgDelete
End Sub

Public Sub update()
    
    Dim oKeywordEclusion As KeywordExcluded
    Dim intIndex As Integer
    
'***************************************
' Module/Form Name   : KeywordExclusions
'
' Procedure Name     : update
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
    For intIndex = 1 To mCol.Count
        Set oKeywordEclusion = mCol.Item(intIndex)
        Select Case oKeywordEclusion.Dirty
        Case Is = pdgCreate
            oKeywordEclusion.Create
        Case Is = pdgDelete
            oKeywordEclusion.Delete
        End Select
    Next intIndex
    '
    '   Now refresh to pick up any other updates made by anyone else.
    '
    Refresh
'
'********* Code Ends Here **************
'
    On Error GoTo 0
    Exit Sub
'
ErrorHandler:
    DisplayError , "KeywordExclusions.update", vbExclamation
    'To pass the error back to the calling proc., use the following line
'    ErrorRaise "KeywordExclusions.update"
End Sub

Public Sub Validate(ByVal pstrKeywordExcluded As String)
    
    Dim oKeywordExcluded    As KeywordExcluded
    Dim intIndex            As Integer
    Dim blnKeywordExists    As Boolean
    
'***************************************
' Module/Form Name   : KeywordExclusions
'
' Procedure Name     : Validate
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
    If InStr(pstrKeywordExcluded, ",") <> 0 Or _
       InStr(pstrKeywordExcluded, " ") <> 0 Or _
       InStr(pstrKeywordExcluded, ";") <> 0 Or _
       InStr(pstrKeywordExcluded, "(") <> 0 Or _
       InStr(pstrKeywordExcluded, ")") <> 0 Or _
       InStr(pstrKeywordExcluded, ".") <> 0 Or _
       InStr(pstrKeywordExcluded, "/") <> 0 Or _
       InStr(pstrKeywordExcluded, "\") <> 0 Or _
       InStr(pstrKeywordExcluded, "&") <> 0 Or _
       InStr(pstrKeywordExcluded, ":") <> 0 Then
        Err.Raise vbObjectError + 1, , "Invalid characters in Excluded Keyword"
    End If
       
    blnKeywordExists = False
    For intIndex = 1 To mCol.Count
        Set oKeywordExcluded = mCol.Item(intIndex)
        If LCase(oKeywordExcluded.Keyword) = LCase(pstrKeywordExcluded) Then
            blnKeywordExists = True
            Exit For
        End If
    Next intIndex

    If blnKeywordExists Then
        Err.Raise vbObjectError + 2, , "Excluded Keyword already exists"
    End If
'
'********* Code Ends Here **************
'
    On Error GoTo 0
    Exit Sub
'
ErrorHandler:
    ErrorRaise "KeywordExclusions.Validate"
End Sub

