VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Activities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private mCol                    As Collection
Private mvarActivityType        As String
Private mvarCustomerNo          As Integer
Private mvarJobNo               As Long
Private mvarSearchNo            As Long

Public Function Add(Key As String, Optional sKey As String) As Activity
    'create a new object
    Dim objNewMember As Activity
    Set objNewMember = New Activity

    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As Activity
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
    mvarActivityType = ""
    mvarCustomerNo = 0
    mvarJobNo = 0
    mvarSearchNo = 0
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

Public Property Let ActivityType(vData As String)
    mvarActivityType = vData
End Property

Public Property Get ActivityType() As String
    ActivityType = mvarActivityType
End Property

Public Property Let CustomerNo(vData As Integer)
    mvarCustomerNo = vData
End Property

Public Property Let JobNo(vData As Long)
    mvarJobNo = vData
End Property

Public Property Let SearchNo(vData As Long)
    mvarSearchNo = vData
End Property

Public Sub Refresh()
'***************************************
' Module/Form Name   : Activities
'
' Procedure Name     : Refresh
'
' Purpose            :
'
' Date Created       : 24/06/2004
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
    
    Dim strSQL          As String
''    Dim rstActivities   As ADODB.Recordset
    Dim rstActivities   As DAO.Recordset
    Dim oActivity       As Activity
    Dim strKey          As String
    Dim strWhereClause  As String
    '
    '   Clear all existing Jobs.
    '
    Set mCol = Nothing
    Set mCol = New Collection
    '
    '   Build the Where Clause.
    '
    strWhereClause = ""
    If mvarActivityType <> "" Then
        strWhereClause = "AND ACTIVITYTYPE = '" & mvarActivityType & "' "
    End If
    '
    If mvarCustomerNo <> 0 Then
        strWhereClause = strWhereClause & "AND CUSTOMERNO = " & mvarCustomerNo & " "
    End If
    '
    If mvarJobNo <> 0 Then
        strWhereClause = strWhereClause & "AND JOBNO = " & mvarJobNo & " "
    End If
    '
    If mvarSearchNo <> 0 Then
        strWhereClause = strWhereClause & "AND SEARCHNO = " & mvarSearchNo & " "
    End If
    '
    '   Refresh Activities.
    '
    strSQL = "SELECT * "
    strSQL = strSQL & "FROM ACTIVITIES "
    strSQL = strSQL & "WHERE STATUS IN ('0','1') "
    strSQL = strSQL & strWhereClause
    '
''    Set rstActivities = New ADODB.Recordset
''    rstActivities.Open strSQL, gdbADO, adOpenForwardOnly, , adCmdText
    Set rstActivities = db.OpenRecordset(strSQL, dbOpenSnapshot)
    With rstActivities
        While Not .EOF
            strKey = NVL(!ID, NVLNumeric)
            Set oActivity = Add(strKey, strKey)
            oActivity.ID = NVL(!ID, NVLNumeric)
            oActivity.ActivityType = NVL(!ActivityType, NVLString)
            oActivity.CustomerNo = NVL(!CustomerNo, NVLNumeric)
            oActivity.JobNo = NVL(!JobNo, NVLNumeric)
            oActivity.SearchNo = NVL(!SearchNo, NVLNumeric)
            oActivity.Description = Trim(NVL(!Description, NVLString))
            oActivity.Status = NVL(!Status, NVLNumeric)
            oActivity.StartDate = NVL(!StartDate, NVLNumeric)
            oActivity.EndDate = NVL(!EndDate, NVLNumeric)
            oActivity.UserField1 = NVL(!UserField1, NVLString)
            oActivity.UserField2 = NVL(!UserField2, NVLString)
            oActivity.CheckPoint
            .MoveNext
        Wend
    End With
    rstActivities.Close
    Set rstActivities = Nothing
'
'********* Code Ends Here **************
'
   Exit Sub
'
Refresh_Error:
    ErrorRaise "Activities.Refresh"
End Sub

