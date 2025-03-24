VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Chasers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private mCol                        As Collection
Private mvarDeliveryNoteNo          As Long
Private mvarCurrentId               As Long
Private mvarLatestChaserDate        As Date
Private mvarLatestContact           As String
Private mvarOriginalReturnByDate    As Date
Private mvarLatestNextAction        As String
Private mvarLastID                  As Long

Public Function Add(ID As Long, Optional sKey As String) As Chaser
    'create a new object
    Dim objNewMember As Chaser
    Set objNewMember = New Chaser

    objNewMember.ID = ID
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As Chaser
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

Public Sub Refresh()
    Dim strSQL As String
''    Dim rstChasers As ADODB.Recordset
    Dim rstChasers As DAO.Recordset
    Dim oChaser As Chaser
    Dim blnCurrentIdFound As Boolean
    Dim strPreviousId As Long
    Dim dteLatestChaserDate
    
    On Error GoTo ErrorProc
    '
    '   Clear all existing Business Types.
    '
    While Not mCol.Count = 0
        mCol.Remove 1
    Wend
    '
    '   Refresh Business Types.
    '
    strSQL = "SELECT * " _
           & "FROM CHASER " _
           & "WHERE DELIVERYNOTENO = " & CStr(mvarDeliveryNoteNo) & " " _
           & "ORDER BY ID "
    Set rstChasers = db.OpenRecordset(strSQL, dbOpenSnapshot)
    With rstChasers
        strPreviousId = 0
        blnCurrentIdFound = False
        mvarOriginalReturnByDate = 0
        mvarLatestChaserDate = 0
        mvarLatestContact = ""
        mvarLatestNextAction = ""
        While Not .EOF
            '
            '   Determine the orginal return by date.
            '
            If !Action = "None" And !NextAction = "None" Then
                mvarOriginalReturnByDate = !ReturnByDate
            Else
                Set oChaser = Add(!ID, !ID)
                With oChaser
                    .DeliveryNoteNo = NVL(rstChasers!DeliveryNoteNo, NVLNumeric)
                    .Contact = NVL(rstChasers!Contact)
                    .Action = NVL(rstChasers!Action, NVLString)
                    .User = Format(rstChasers!User)
                    .ChaserDate = NVL(rstChasers!ReturnByDate, NVLNumeric)
                    .NextAction = NVL(rstChasers!NextAction, NVLString)
                    .Comment = NVL(rstChasers!Comment, NVLString)
                    .DateAmended = NVL(rstChasers!DateAmended, NVLNumeric)
                    .DateCreated = NVL(rstChasers!DateCreated, NVLNumeric)
                    .DatePrinted = NVL(rstChasers!DatePrinted, NVLNumeric)
                    oChaser.CheckPoint
                End With
            End If
            '
            '   Determine the current ID.
            '
            If Not blnCurrentIdFound Then
                If !ID > mvarCurrentId Then
                    If strPreviousId = 0 Then
                        mvarCurrentId = !ID
                        blnCurrentIdFound = True
                    Else
                        mvarCurrentId = strPreviousId
                        blnCurrentIdFound = True
                    End If
                ElseIf !ID = mvarCurrentId Then
                    blnCurrentIdFound = True
                End If
            End If
            strPreviousId = !ID
            .MoveNext
            
            If Not oChaser Is Nothing Then
                If oChaser.Action <> "Invoiced" Then
                    mvarLastID = oChaser.ID
                    mvarLatestChaserDate = oChaser.ChaserDate
                    mvarLatestNextAction = oChaser.NextAction
                Else
                    mvarLatestNextAction = oChaser.Action
                End If
                mvarLatestContact = oChaser.Contact
            End If
        Wend
    End With
    If Not blnCurrentIdFound Then
        mvarCurrentId = strPreviousId
    End If
    
    rstChasers.Close
    Set rstChasers = Nothing
    Exit Sub
ErrorProc:
    Err.Raise Err.Number, "Chasers: Refresh"
End Sub

Public Property Get LastID() As Long
    LastID = mvarLastID
End Property

Public Property Let LastID(vData As Long)
    mvarLastID = vData
End Property

Public Property Get CurrentId() As Long
    CurrentId = mvarCurrentId
End Property

Public Property Let CurrentId(vData As Long)
    mvarCurrentId = vData
End Property

Public Property Get DeliveryNoteNo() As Long
    DeliveryNoteNo = mvarDeliveryNoteNo
End Property

Public Property Let DeliveryNoteNo(ByVal vNewValue As Long)
    mvarDeliveryNoteNo = vNewValue
    Refresh
End Property

Public Property Get LatestChaserDate() As Date
    LatestChaserDate = mvarLatestChaserDate
End Property

Public Property Let LatestChaserDate(vData As Date)
    mvarLatestChaserDate = vData
End Property

Public Property Get LatestContact() As String
    LatestContact = mvarLatestContact
End Property

Public Property Let LatestContact(vData As String)
    mvarLatestContact = vData
End Property

Public Property Get OriginalReturnByDate() As Date
    OriginalReturnByDate = mvarOriginalReturnByDate
End Property

Public Property Let OriginalReturnByDate(vData As Date)
    mvarOriginalReturnByDate = vData
End Property

Public Property Get LatestNextAction() As String
    LatestNextAction = mvarLatestNextAction
End Property

Public Property Let LatestNextAction(vData As String)
    mvarLatestNextAction = vData
End Property

