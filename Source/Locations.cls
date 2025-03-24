VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Locations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private mCol As Collection

Public Function Add(Key As String, Optional sKey As String) As Location
    'create a new object
    Dim objNewMember As Location
    Set objNewMember = New Location

    objNewMember.Key = Key
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As Location
Attribute Item.VB_UserMemId = 0
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
    Dim strSQL          As String
    Dim rstLocations    As ADODB.Recordset
    Dim oLocation       As Location
    Dim strKey          As String
    '
    '   Clear all existing Jobs.
    '
    Set mCol = Nothing
    Set mCol = New Collection
    '
    '   Refresh Jobs.
    '
    strSQL = "SELECT * "
    strSQL = strSQL & "FROM LOCATIONS"

    Set rstLocations = New ADODB.Recordset
    rstLocations.Open strSQL, gdbADO, adOpenForwardOnly, , adCmdText
    With rstLocations
        While Not .EOF
            strKey = NVL(!LocationType, NVLString) & NVL(!ID, NVLString) & NVL(!Machine, NVLString)
            Set oLocation = Add(strKey, strKey)
            oLocation.LocationType = Trim(NVL(!LocationType, NVLString))
            oLocation.ID = NVL(!ID, NVLString)
            oLocation.Description = NVL(!Description, NVLString)
            oLocation.Prefix = NVL(!Prefix, NVLString)
            oLocation.Suffix = NVL(!Suffix, NVLString)
            oLocation.Checkpoint
            .MoveNext
        Wend
    End With
    rstLocations.Close
    Set rstLocations = Nothing
    '
    '********* Code Ends Here **************
    '
    Exit Sub
    '
Refresh_Error:
    ErrorRaise "Locations.Refresh"
End Sub



