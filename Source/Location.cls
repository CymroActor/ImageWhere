VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Location"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Key As String

Private mvarLocationType    As String
Private mvarID              As String
Private mvarMachine         As String
Private mvarDescription     As String
Private mvarPrefix          As String
Private mvarSuffix          As String

Private Type udtLocation
    LocationType    As String
    ID              As String
    Machine         As String
    Description     As String
    Prefix          As String
    Suffix          As String
End Type

Private CheckPointLocation As udtLocation

Public Property Let LocationType(ByVal vData As String)
    mvarLocationType = vData
End Property

Public Property Get LocationType() As String
    LocationType = mvarLocationType
End Property

Public Property Let ID(ByVal vData As String)
    mvarID = vData
End Property

Public Property Get ID() As String
    ID = mvarID
End Property

Public Property Let Machine(ByVal vData As String)
    mvarMachine = vData
End Property

Public Property Get Machine() As String
    Machine = mvarMachine
End Property

Public Property Let Description(ByVal vData As String)
    mvarDescription = vData
End Property

Public Property Get Description() As String
    Description = mvarDescription
End Property

Public Property Let Prefix(ByVal vData As String)
    mvarPrefix = vData
End Property

Public Property Get Prefix() As String
    Prefix = mvarPrefix
End Property

Public Property Let Suffix(ByVal vData As String)
    mvarSuffix = vData
End Property

Public Property Get Suffix() As String
    Suffix = mvarSuffix
End Property

Public Sub update()
    Dim strSQL As String
    Dim recLocation As ADODB.Recordset
    
    On Error GoTo ErrorProc
    '
    '   Get current copy of data and lock it.
    '
    strSQL = "SELECT * "
    strSQL = strSQL & "FROM LOCATIONS "
    strSQL = strSQL & "WHERE LOCATIONTYPE = '" & mvarLocationType & "' "
    strSQL = strSQL & "AND ID = '" & mvarID & "' "
    
    Set recLocation = New ADODB.Recordset
    recLocation.Open strSQL, gdbADO, adOpenDynamic, adLockOptimistic, adCmdText
    '
    '   Update Data.
    '
    With recLocation
        CheckRecordSame recLocation
        !Description = mvarDescription
        !Prefix = mvarPrefix
        !Suffix = mvarSuffix
        .update
    End With

    Set recLocation = Nothing
    CheckPoint
    Exit Sub
ErrorProc:
    Set recLocation = Nothing
    Err.Raise Err.Number, "Comment: Update"
End Sub

Public Sub CheckPoint()
    With CheckPointLocation
        .LocationType = mvarLocationType
        .ID = mvarID
        .Machine = mvarMachine
        .Description = mvarDescription
        .Prefix = mvarPrefix
        .Suffix = mvarSuffix
    End With
End Sub

Public Sub Delete()
    
    Dim recLocation As ADODB.Recordset
    Dim strSQL As String
    Dim lngRecordsAffected As Long
    
    On Error GoTo ErrorProc
    '
    '   Read comment and lock before deleting.
    '
    strSQL = "SELECT * "
    strSQL = strSQL & "FROM LOCATIONS "
    strSQL = strSQL & "WHERE LOCATIONTYPE = '" & mvarLocationType & "' "
    strSQL = strSQL & "AND ID = '" & mvarID & "' "
    
    Set recLocation = New ADODB.Recordset
    recLocation.Open strSQL, gdbADO, adOpenDynamic, adLockOptimistic, adCmdText
    
    CheckRecordSame recLocation
    strSQL = "DELETE * FROM LOCATIONS "
    strSQL = strSQL & "WHERE LOCATIONTYPE = '" & mvarLocationType & "' "
    strSQL = strSQL & "AND ID = '" & mvarID & "' "
    
    gdbADO.Execute strSQL, lngRecordsAffected, adCmdText
    If lngRecordsAffected <> 1 Then
        Err.Raise vbObjectError + 3, "Location: Delete", "Location does not Exist"
    End If
    
    Set recLocation = Nothing
    Exit Sub
ErrorProc:
    Set recLocation = Nothing
    Err.Raise Err.Number, "Location: Delete"
End Sub

Private Function CheckRecordSame(recRecord As ADODB.Recordset) As Boolean
'***************************************
' Module/Form Name   : Location
'
' Procedure Name     : CheckRecordSame
'
' Purpose            :
'
' Date Created       : 24/06/2004
'
' Author             : GARETH SAUNDERS
'
' Parameters         : recRecord - ADODB.Recordset
'
' Returns            : Boolean
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo CheckRecordSame_Error
'
'******** Code Starts Here *************
'
    CheckRecordSame = True
    '
    '   Check it record has changed by comparing it with the data
    '   at last refresh.
    '
    With recRecord
        If NVL(!LocationType, NVLString) <> CheckPointLocation.LocationType Or _
           NVL(!ID, NVLString) <> CheckPointLocation.ID Or _
           NVL(!Machine, NVLString) <> CheckPointLocation.Machine Or _
           NVL(!Description, NVLString) <> CheckPointLocation.Description Or _
           NVL(!Prefix, NVLString) <> CheckPointLocation.Prefix Or _
           NVL(!Suffix, NVLString) <> CheckPointLocation.Suffix Then
            Refresh recRecord
            CheckPoint
            Err.Raise vbObjectError + 5, "Location: CheckRecordSame", "Location has been updated by someone else since first obtaining it."
        End If
    End With
    CheckRecordSame = False
'
'********* Code Ends Here **************
'
   Exit Function
'
CheckRecordSame_Error:
    ErrorRaise "Location.CheckRecordSame"
End Function

Private Function Refresh(recRecord As ADODB.Recordset)
    mvarLocationType = recRecord!LocationType
    mvarID = recRecord!ID
    mvarMachine = recRecord!Machine
    mvarDescription = recRecord!Description
    mvarPrefix = recRecord!Prefix
    mvarSuffix = recRecord!Suffix
End Function



