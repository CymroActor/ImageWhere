VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BatchKeywords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mCol            As Collection
Private mvarBatchNo     As Long

Public Property Get BatchNo() As Long
    BatchNo = mvarBatchNo
End Property

Public Property Let BatchNo(ByVal vData As Long)
    mvarBatchNo = vData
End Property

Public Sub Refresh()
'***************************************
' Module/Form Name   : BatchKeywords
'
' Procedure Name     : Refresh
'
' Purpose            :
'
' Date Created       : 07/12/2005 20:52
'
' Author             : Gareth
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
    Dim strSQL              As String
    Dim recBatchKeywords    As DAO.Recordset
    Dim objBatchKeyword     As BatchKeyword
    '
    '   Clear all existing BatchKeywords.
    '
    On Error GoTo ErrorProc

    Set mCol = Nothing
    Set mCol = New Collection
    '
    '   Refresh BatchKeywords Collection.
    '
    strSQL = "SELECT * FROM BatchKeywords WHERE BatchNo = " & mvarBatchNo & " ORDER BY KEYWORD"
    Set recBatchKeywords = db.OpenRecordset(strSQL, dbOpenSnapshot, dbForwardOnly)
    With recBatchKeywords
        While Not .EOF
            On Error Resume Next
            Set objBatchKeyword = Add(!BatchNo, _
                                      UCase(!Keyword), _
                                      !KeywordType)
            If Err.Number <> 0 And _
               Err.Number <> 457 Then        '   Duplicate keyword doesn't really matter.
                ErrorSave
                On Error GoTo ErrorProc
                ErrorRestore
            Else
                On Error GoTo ErrorProc
            End If
            '
            objBatchKeyword.Dirty = pdgClean
            .MoveNext
        Wend
    End With

    recBatchKeywords.Close
    Set recBatchKeywords = Nothing

    On Error GoTo 0
    Exit Sub
'
'********* Code Ends Here **************
'
ErrorProc:
    ErrorRaise "BatchKeywords.Refresh"
End Sub

Public Function Add(ByVal BatchNo As Long, _
                    ByVal Keyword As String, _
                    ByVal KeywordType As String) As BatchKeyword
    'create a new object
    Dim objNewMember As BatchKeyword
    Set objNewMember = New BatchKeyword
    
    'set the properties passed into the method
    objNewMember.BatchNo = BatchNo
    objNewMember.Keyword = Keyword
    objNewMember.KeywordType = KeywordType
    objNewMember.Dirty = pdgCreate
    On Error Resume Next
    mCol.Add objNewMember, Keyword
    '
    '   If the object is already there but is pending delete, simply set it to be updated.
    '
    If Err.Number = 457 Then
        Set objNewMember = mCol.Item(Keyword)
        objNewMember.Dirty = pdgUpdate
        objNewMember.KeywordType = KeywordType
    ElseIf Err.Number <> 0 Then
        ErrorSave
        On Error GoTo 0
        ErrorRestore
    End If
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As BatchKeyword
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub

Public Sub RemovePending(vntIndexKey As Variant)
    mCol.Item(vntIndexKey).Dirty = pdgDelete
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

Public Sub update()
    
    Dim oKeyword As BatchKeyword
    Dim intIndex As Integer
    
'***************************************
' Module/Form Name   : BatchKeywords
'
' Procedure Name     : Update
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
    For intIndex = 1 To mCol.Count
        Set oKeyword = mCol.Item(intIndex)
        Select Case oKeyword.Dirty
        Case Is = pdgCreate
            oKeyword.Create
        Case Is = pdgDelete
            oKeyword.Delete
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
    ErrorRaise "BatchKeywords.Update"
End Sub

Public Sub Validate(ByVal pstrKeyword As String, ByVal pAction As UpdateMode)
    
    Dim oKeyword            As BatchKeyword
    Dim intIndex            As Integer
    Dim blnKeywordExists    As Boolean
    
'***************************************
' Module/Form Name   : BatchKeywords
'
' Procedure Name     : Validate
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

    If InStr(pstrKeyword, ",") <> 0 Or _
       InStr(pstrKeyword, " ") <> 0 Or _
       InStr(pstrKeyword, ";") <> 0 Or _
       InStr(pstrKeyword, "(") <> 0 Or _
       InStr(pstrKeyword, ")") <> 0 Or _
       InStr(pstrKeyword, ".") <> 0 Or _
       InStr(pstrKeyword, "/") <> 0 Or _
       InStr(pstrKeyword, "\") <> 0 Or _
       InStr(pstrKeyword, "&") <> 0 Or _
       InStr(pstrKeyword, """") <> 0 Or _
       InStr(pstrKeyword, ":") <> 0 Then
        Err.Raise vbObjectError + 1, , "Invalid characters in Keyword"
    End If
       
    blnKeywordExists = False
    For intIndex = 1 To mCol.Count
        Set oKeyword = mCol.Item(intIndex)
        If LCase(oKeyword.Keyword) = LCase(pstrKeyword) And oKeyword.Dirty <> pdgDelete Then
            blnKeywordExists = True
            Exit For
        End If
    Next intIndex

    If blnKeywordExists Then
        Err.Raise vbObjectError + 2, , "Keyword already exists"
    End If
'
'********* Code Ends Here **************
'
    On Error GoTo 0
    Exit Sub
'
ErrorHandler:
    ErrorRaise "BatchKeywords.Validate"
End Sub

Public Sub DeleteAll(ByVal lngFrom As Long, ByVal lngTo As Long)
'***************************************
' Module/Form Name   : BatchKeywords
'
' Procedure Name     : DeleteAll
'
' Purpose            :
'
' Date Created       : 08/10/2005 20:11
'
' Author             : Gareth
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
    On Error GoTo ErrorProc

    Dim strSQL As String
    
    strSQL = "DELETE * FROM BATCHKEYWORDS "
    strSQL = strSQL & "WHERE KEYWORDTYPE <> 'M' "
    strSQL = strSQL & "AND BATCHNO >= " & CStr(lngFrom) & " "
    strSQL = strSQL & "AND BATCHNO <= " & CStr(lngTo) & " "
    
''    gdbADO.Execute strSQL, dbFailOnError
    db.Execute strSQL, dbFailOnError

    On Error GoTo 0
    Exit Sub
'
'********* Code Ends Here **************
'
ErrorProc:
    ErrorRaise "BatchKeywords.DeleteAll"
End Sub

