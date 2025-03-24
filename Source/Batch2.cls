VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Batch2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mvarBatchNo                 As Long
Private mvarParentBatch             As Long
Private mvarDescription             As String
Private mvarFormat                  As String
Private mvarCopyRight               As String
Private mvarDateTaken               As Date
Private mvarCredit                  As String
Private mvarTiffLocation            As String
Private mvarExclusive               As Boolean
Private mvarKeyWords                As BatchKeywords
Private mvarKeywordExclusions       As KeywordExclusions
Private mvarKeywordSynonyms         As KeywordSynonyms
Private mvarDigitalImages           As DigitalImages
    
Public Property Get BatchNo() As Long
    BatchNo = mvarBatchNo
End Property

Public Property Let BatchNo(ByVal vData As Long)
    mvarBatchNo = vData
End Property

Public Property Get ParentBatch() As Long
    ParentBatch = mvarParentBatch
End Property

Public Property Let ParentBatch(ByVal vData As Long)
    mvarParentBatch = vData
End Property

Public Property Get Description() As String
    Description = mvarDescription
End Property

Public Property Let Description(ByVal vData As String)
    mvarDescription = vData
End Property

Public Property Get Format() As String
    Format = mvarFormat
End Property

Public Property Let Format(ByVal vData As String)
    mvarFormat = vData
End Property

Public Property Get CopyRight() As String
    CopyRight = mvarCopyRight
End Property

Public Property Let CopyRight(ByVal vData As String)
    mvarCopyRight = vData
End Property

Public Property Get DateTaken() As Date
    DateTaken = mvarDateTaken
End Property

Public Property Let DateTaken(ByVal vData As Date)
    mvarDateTaken = vData
End Property

Public Property Get Credit() As String
    Credit = mvarCredit
End Property

Public Property Let Credit(ByVal vData As String)
    mvarCredit = vData
End Property

Public Property Get TiffLocation() As String
    TiffLocation = mvarTiffLocation
End Property

Public Property Let TiffLocation(ByVal vData As String)
    mvarTiffLocation = vData
End Property

Public Property Get Exclusive() As Boolean
    Exclusive = mvarExclusive
End Property

Public Property Let Exclusive(ByVal vData As Boolean)
    mvarExclusive = vData
End Property

Public Property Get Keywords() As BatchKeywords
    If mvarKeyWords Is Nothing Then
        Set mvarKeyWords = New BatchKeywords
        mvarKeyWords.BatchNo = mvarBatchNo
        mvarKeyWords.Refresh
    End If

    Set Keywords = mvarKeyWords
End Property

Public Property Set Keywords(vData As BatchKeywords)
    Set mvarKeyWords = vData
End Property

Public Property Get DigitalImages() As DigitalImages
    If mvarDigitalImages Is Nothing Then
        Set mvarDigitalImages = New DigitalImages
        mvarDigitalImages.BatchNo = mvarBatchNo
        mvarDigitalImages.Refresh
    End If

    Set DigitalImages = mvarDigitalImages
End Property

Public Property Set DigitalImages(vData As DigitalImages)
    Set mvarDigitalImages = vData
End Property

Public Sub DeleteAutomaticKeywords()
'***************************************
' Module/Form Name   : Batch2
'
' Procedure Name     : DeleteAutomaticKeywords
'
' Purpose            :
'
' Date Created       : 14/07/2005
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo DeleteAutomaticKeywords_Error
'
'******** Code Starts Here *************
'
    Dim strSQL As String
    
    strSQL = "DELETE " _
           & "FROM BATCHKEYWORDS " _
           & "WHERE BatchNo = " & mvarBatchNo & " AND " _
           & "KeywordType IN ('A','S')"
    
    db.Execute strSQL, dbFailOnError
'
'********* Code Ends Here **************
'
   Exit Sub
'
DeleteAutomaticKeywords_Error:
    ErrorRaise "Batch2.DeleteAutomaticKeywords"
End Sub

Public Sub CreateAutomaticKeywords(Optional ByRef oKeywordExclusions As KeywordExclusions = Nothing, _
                                   Optional ByRef oKeywordSynonyms As KeywordSynonyms = Nothing, _
                                   Optional ByRef intFileNo As Integer = 0)
'***************************************
' Module/Form Name   : Batch2
'
' Procedure Name     : CreateAutomaticKeywords
'
' Purpose            :
'
' Date Created       : 14/07/2005
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo CreateAutomaticKeywords_Error
'
'******** Code Starts Here *************
'
    Dim strSQL              As String
    Dim i                   As Integer
    Dim strKeywords()       As String
    Dim tdfWork             As TableDef
    Dim fldTemp             As Field
    Dim intBound            As Integer
    Dim colNewKeywords      As Collection
    Dim oKeywordSynonym     As KeywordSynonym
    Dim oBatchKeyword       As BatchKeyword
    Dim oNewBatchKeyword    As BatchKeyword
    Dim strKeyword          As String
    Dim blnAddKeyword       As Boolean
    '
    '   Store the Keyword Exclusions and Synonyms.
    '
    If oKeywordExclusions Is Nothing Then
        Set mvarKeywordExclusions = New KeywordExclusions
        mvarKeywordExclusions.Refresh
    Else
        Set mvarKeywordExclusions = oKeywordExclusions
    End If
    '
    If oKeywordSynonyms Is Nothing Then
        Set mvarKeywordSynonyms = New KeywordSynonyms
        mvarKeywordSynonyms.Refresh
    Else
        Set mvarKeywordSynonyms = oKeywordSynonyms
    End If
    '
    '   Create the New Keywords collection
    '
    Set colNewKeywords = Nothing
    Set colNewKeywords = New Collection
    '
    '   Add in the Manual Keywords.
    '
    If Not mvarKeyWords Is Nothing Then
        mvarKeyWords.Refresh
    End If
    '
    For Each oBatchKeyword In Keywords
        If oBatchKeyword.KeywordType = "M" Then
            Set oNewBatchKeyword = New BatchKeyword
            oNewBatchKeyword.Keyword = oBatchKeyword.Keyword
            oNewBatchKeyword.KeywordType = oBatchKeyword.KeywordType
            On Error Resume Next
            colNewKeywords.Add oNewBatchKeyword, oBatchKeyword.Keyword
            If Err.Number <> 457 And Err.Number <> 0 Then
                ErrorSave
                On Error GoTo CreateAutomaticKeywords_Error
                ErrorRestore
            End If
            '
            Set oNewBatchKeyword = Nothing
        End If
    Next oBatchKeyword
    '
    '   Create an array with each Automatic Keyword in generated from the description.
    '
    strKeywords() = KeywordSPlit(mvarDescription)
    '
    '   Insert Automatic Keywords.
    '
    On Error Resume Next
    intBound = UBound(strKeywords, 2)
    If Err.Number <> 9 Then      '   Empty Array.
        On Error GoTo CreateAutomaticKeywords_Error
        For i = 0 To UBound(strKeywords, 2)
            Set oNewBatchKeyword = New BatchKeyword
            oNewBatchKeyword.Keyword = strKeywords(0, i)
            oNewBatchKeyword.KeywordType = "A"
            On Error Resume Next
            colNewKeywords.Add oNewBatchKeyword, strKeywords(0, i)
            If Err.Number <> 457 And Err.Number <> 0 Then
                ErrorSave
                On Error GoTo CreateAutomaticKeywords_Error
                ErrorRestore
            End If
            '
            Set oNewBatchKeyword = Nothing
        Next i
    End If
    On Error GoTo CreateAutomaticKeywords_Error
    '
    '   For each Keyword Synonym, check to see if it exists in the new keyword collection.
    '   If so, then copy the synonym.
    '
    For Each oKeywordSynonym In mvarKeywordSynonyms
        blnAddKeyword = False
''        If InStr(UCase(mvarDescription), UCase(oKeywordSynonym.Keyword)) <> 0 Then
''            blnAddKeyword = True
''        Else
            For Each oBatchKeyword In colNewKeywords
                If UCase(oBatchKeyword.Keyword) = UCase(oKeywordSynonym.Keyword) Then
                    blnAddKeyword = True
                    Exit For
                End If
            Next oBatchKeyword
''        End If
        '
        If blnAddKeyword Then
            Set oNewBatchKeyword = New BatchKeyword
            oNewBatchKeyword.Keyword = oKeywordSynonym.Synonym
            oNewBatchKeyword.KeywordType = "S"
            On Error Resume Next
            colNewKeywords.Add oNewBatchKeyword, oKeywordSynonym.Synonym
            If Err.Number <> 457 And Err.Number <> 0 Then
                ErrorSave
                On Error GoTo CreateAutomaticKeywords_Error
                ErrorRestore
            End If
            '
            Set oNewBatchKeyword = Nothing
        End If
    Next oKeywordSynonym
    '
    '   Now apply the new keywords.
    '
    For Each oBatchKeyword In colNewKeywords
        '
        '   Don't write away manual keywords.
        '
        If oBatchKeyword.KeywordType <> "M" Then
            '
            '   Is this to be written to the Database or a file?
            '
            If intFileNo = 0 Then
                strSQL = "INSERT INTO BATCHKEYWORDS (BATCHNO, KEYWORD, KEYWORDTYPE) VALUES("
                strSQL = strSQL & "" & mvarBatchNo & ""
                strSQL = strSQL & ",'" & UCase(JetSQLFixup(oBatchKeyword.Keyword)) & "'"
                strSQL = strSQL & ",'" & oBatchKeyword.KeywordType & "'"
                strSQL = strSQL & ")"
                db.Execute strSQL, dbFailOnError
            Else
                Write #intFileNo, CStr(mvarBatchNo), UCase(oBatchKeyword.Keyword), oBatchKeyword.KeywordType
            End If
        End If
    Next oBatchKeyword
    On Error GoTo CreateAutomaticKeywords_Error
'
'********* Code Ends Here **************
'
   Exit Sub
'
CreateAutomaticKeywords_Error:
    ErrorRaise "Batch2.CreateAutomaticKeywords"
End Sub

Private Function KeywordSPlit(parDescription As String) As String()
'***************************************
' Module/Form Name   : Batch2
'
' Procedure Name     : KeywordSPlit
'
' Purpose            :
'
' Date Created       : 21/07/2005
'
' Author             : Gareth Saunders
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                      21/07/2005 Gareth Saunders
'
'***************************************
'
    On Error GoTo ErrorHandler:
'
'******** Code Starts Here *************
'
    Dim i, j                As Integer
    Dim strKeyword          As String
    Dim strKeyWordArray()   As String
    Dim strDescriptionByte  As String
    Dim intBound            As Integer
    Dim intChar             As Integer
    '
    strKeyword = ""
    j = 0
    For i = 1 To Len(parDescription)
        strDescriptionByte = Mid(parDescription, i, 1)
        '
        '   Only allow alphabetic characters.
        '   Also allow hyphens and apostrophes. (Requested by Jane on 08/12/2005)
        '
        intChar = Asc(strDescriptionByte)
        If Not ((intChar >= 65 And intChar <= 90) Or _
               (intChar >= 97 And intChar <= 122) Or _
                intChar = 39 Or intChar = 45) Then
            If strKeyword <> "" Then
                AddAutomaticKeyword strKeyword, strKeyWordArray(), j
            End If
        Else
            strKeyword = strKeyword & strDescriptionByte
        End If
    Next i
    '
    AddAutomaticKeyword strKeyword, strKeyWordArray, j
    
    KeywordSPlit = strKeyWordArray
'
'********* Code Ends Here **************
'
    On Error GoTo 0
    Exit Function
'
ErrorHandler:
    ErrorRaise "Batch2.KeywordSPlit"
End Function

Private Sub AddAutomaticKeyword(ByRef pstrKeyword As String, _
                                ByRef pstrKeyWordArray() As String, _
                                ByRef pintkeywordposition As Integer)
    
    Dim oKeywordExclusion   As KeywordExcluded
    
'***************************************
' Module/Form Name   : Batch2
'
' Procedure Name     : AddAutomaticKeyword
'
' Purpose            :
'
' Date Created       : 02/01/2006 18:31
'
' Author             : Gareth
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
    On Error GoTo ErrorProc
    
    Dim strKeywordToApply       As String

    If pstrKeyword <> "" Then
        '
        '   Ignore purely numeric keywords such as 1 or 57.
        '
        If Not IsNumeric(pstrKeyword) Then
            '
            '   Check to see if it is excluded and ensure it is not a single
            '   character hyphen or apostrophe.
            '
            strKeywordToApply = pstrKeyword
            
            If Trim(strKeywordToApply) = "'" Or Trim(strKeywordToApply) = "-" Then
                strKeywordToApply = ""
            End If
            
            If Right(strKeywordToApply, 1) = "'" Or Right(strKeywordToApply, 1) = "-" Then
                strKeywordToApply = Left(strKeywordToApply, Len(strKeywordToApply) - 1)
            End If
            
            If Left(strKeywordToApply, 1) = "'" Or Left(strKeywordToApply, 1) = "-" Then
                strKeywordToApply = Right(strKeywordToApply, Len(strKeywordToApply) - 1)
            End If
            
            If strKeywordToApply <> "" Then
                On Error Resume Next
                Set oKeywordExclusion = mvarKeywordExclusions.Item(strKeywordToApply)
                If Err.Number = 5 Then
                    ReDim Preserve pstrKeyWordArray(2, pintkeywordposition)
                    pstrKeyWordArray(0, pintkeywordposition) = strKeywordToApply
                    pstrKeyWordArray(1, pintkeywordposition) = "A"
                    pintkeywordposition = pintkeywordposition + 1
                End If
                On Error GoTo ErrorProc:
            End If
        End If
        pstrKeyword = ""
    End If

    On Error GoTo 0
    Exit Sub
'
'********* Code Ends Here **************
'
ErrorProc:
    ErrorRaise "Batch2.AddAutomaticKeyword"
End Sub

Public Sub Read(parBatchNo As Long)
    
    Dim strSQL      As String
    Dim recBatch    As DAO.Recordset
    
'***************************************
' Module/Form Name   : Batch2
'
' Procedure Name     : Read
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
    strSQL = "SELECT * " _
           & "FROM BATCH " _
           & "WHERE BATCH_NO = " & CStr(parBatchNo)
    
    Set recBatch = db.OpenRecordset(strSQL, dbOpenSnapshot)
    With recBatch
        If Not .EOF Then
            mvarBatchNo = !batch_no
            mvarParentBatch = !parent_batch
            mvarDescription = !Description & ""
            mvarFormat = !Format & ""
            mvarCopyRight = !CopyRight & ""
            If IsNull(!date_taken) Then
                mvarDateTaken = 0
            Else
                mvarDateTaken = !date_taken
            End If
            mvarCredit = !Credit & ""
            mvarTiffLocation = !TiffLocation & ""
            mvarExclusive = !Exclusive
        End If
    End With
    Set recBatch = Nothing
'
'********* Code Ends Here **************
'
   On Error GoTo 0
   Exit Sub
'
ErrorHandler:
    ErrorRaise "Batch2.Read"
End Sub

Public Sub update()
    
    Dim strSQL      As String
    
'***************************************
' Module/Form Name   : Batch2
'
' Procedure Name     : Update
'
' Purpose            :
'
' Date Created       : 11/08/2005
'
' Author             : Gareth Saunders
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                      11/08/2005 Gareth Saunders
'
'***************************************
'
    On Error GoTo ErrorHandler:
'
'******** Code Starts Here *************
'
    strSQL = "UPDATE BATCH SET "
    strSQL = strSQL & "EXCLUSIVE = " & CStr(mvarExclusive) & " "
    strSQL = strSQL & "WHERE BATCH_NO = " & CStr(mvarBatchNo)
    
    db.Execute strSQL, dbFailOnError
'
'********* Code Ends Here **************
'
    On Error GoTo 0
    Exit Sub
'
ErrorHandler:
    ErrorRaise "Batch2.Update"
End Sub

Private Sub Class_Terminate()
    Set mvarKeywordExclusions = Nothing
    Set mvarKeywordSynonyms = Nothing
End Sub
