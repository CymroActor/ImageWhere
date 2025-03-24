VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Batches"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const SW_HIDE = 0

Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Private mCol As Collection

Public Sub Refresh(Optional ByVal plngFrom As Long = 0, _
                   Optional ByVal plngTo As Long = 0)
'***************************************
' Module/Form Name   : Batches
'
' Procedure Name     : Refresh
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
On Error GoTo Refresh_Error
'
'******** Code Starts Here *************
'
    Dim strSQL              As String
    Dim recBatch            As DAO.Recordset
    Dim objBatch            As Batch2
    Dim strWhereClause      As String
    '
    '   Clear all existing Batches.
    '
    Set mCol = Nothing
    Set mCol = New Collection
    '
    '   Refresh Batches Collection.
    '
    strWhereClause = ""
    If plngFrom <> 0 Then
        strWhereClause = "Batch_No >= " & CStr(plngFrom)
    End If
    If plngTo <> 0 Then
        strWhereClause = strWhereClause & IIf(strWhereClause = "", "", " AND ") & "Batch_No <= " & CStr(plngTo)
    End If
    If strWhereClause <> "" Then
        strWhereClause = " WHERE " & strWhereClause
    End If
    '
    strSQL = "SELECT * FROM Batch" & strWhereClause
    Set recBatch = db.OpenRecordset(strSQL, dbOpenSnapshot, dbForwardOnly)
    With recBatch
        While Not .EOF
            Set objBatch = Add(!batch_no)
            objBatch.ParentBatch = !parent_batch
            objBatch.Description = !Description & ""
            objBatch.Format = !Format & ""
            objBatch.CopyRight = !CopyRight & ""
            If IsNull(!date_taken) Then
                objBatch.DateTaken = 0
            Else
                objBatch.DateTaken = !date_taken
            End If
            objBatch.Credit = !Credit & ""
            objBatch.TiffLocation = !TiffLocation & ""
            objBatch.Exclusive = !Exclusive
            .MoveNext
        Wend
    End With

    recBatch.Close
    Set recBatch = Nothing
'
'********* Code Ends Here **************
'
   Exit Sub
'
Refresh_Error:
    ErrorRaise "Batches.Refresh"
End Sub

Public Function Add(ByVal BatchNo As String) As Batch2
    'create a new object
    Dim objNewMember As Batch2
    Set objNewMember = New Batch2
    
    'set the properties passed into the method
    objNewMember.BatchNo = BatchNo
    mCol.Add objNewMember, CStr(BatchNo)
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As Batch2
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

Public Sub ImportTextFile(pstrFile As String)
'***************************************
' Module/Form Name   : Batches
'
' Procedure Name     : ImportTextFile
'
' Purpose            :
'
' Date Created       : 08/10/2005 20:01
'
' Author             : Gareth
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
    Dim strSQL As String
    Dim oFSO    As Scripting.FileSystemObject
    
    On Error GoTo ErrorProc
    
    Set oFSO = New Scripting.FileSystemObject
    strSQL = "INSERT INTO BATCHKEYWORDS "
    strSQL = strSQL & "SELECT * FROM [BatchKeywords.txt] IN """ & oFSO.GetParentFolderName(pstrFile) & """ ""TEXT;"""
    Set oFSO = Nothing
    db.Execute strSQL

    On Error GoTo 0
    Exit Sub
'
'********* Code Ends Here **************
'
ErrorProc:
    ErrorRaise "Batches.ImportTextFile"
End Sub

Public Sub ImportTextFileWithAccess(pstrFile As String)
    Dim oAccessDB As Object
    
    '
    '   Close the database.
    '
'***************************************
' Module/Form Name   : Batches
'
' Procedure Name     : ImportTextFileWithAccess
'
' Purpose            :
'
' Date Created       : 08/01/2006 21:28
'
' Author             : Gareth
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
    On Error GoTo ErrorProc

    db.Close
    gdbADO.Close
    '
    Set oAccessDB = GetObject(glo_dbname, "Access.Application")
    oAccessDB.DoCmd.Maximize
    oAccessDB.DoCmd.RunCommand 10                           'acCmdAppMaximize
    oAccessDB.DoCmd.TransferText 0, , "BatchKeywords", pstrFile, True
    Set oAccessDB = Nothing
    '
    '   Re-open the database.
    '
    Set db = Nothing
    Set db = ws.OpenDatabase(glo_dbname)
    gdbADO.Open gADOConnection

    On Error GoTo 0
    Exit Sub
'
'********* Code Ends Here **************
'
ErrorProc:
    ErrorRaise "Batches.ImportTextFileWithAccess"
End Sub

