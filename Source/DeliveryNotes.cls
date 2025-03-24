VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DeliveryNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mCol As Collection

Public Function Add(ByVal lngDNoteNo As Long, _
                    ByVal blnSelected As Boolean) As DeliveryNote

    'create a new object
    Dim objNewMember As DeliveryNote
'***************************************
' Module/Form Name   : DeliveryNotes
'
' Procedure Name     : Add
'
' Purpose            :
'
' Date Created       : 21/05/2006 19:02
'
' Author             : Gareth
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
    On Error GoTo ErrorProc

    Set objNewMember = New DeliveryNote
    
    'set the properties passed into the method
    objNewMember.DNoteNo = lngDNoteNo
    objNewMember.Selected = blnSelected
    mCol.Add objNewMember, CStr(lngDNoteNo)
    Set Add = objNewMember
    Set objNewMember = Nothing

    On Error GoTo 0
    Exit Function
'
'********* Code Ends Here **************
'
ErrorProc:
    ErrorRaise "DeliveryNotes.Add"
End Function

Public Property Get Item(vntIndexKey As Variant) As DeliveryNote
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

Public Function LatestChaserDate(Optional blnSelectedOnly As Boolean = True) As Date
    
    Dim oDeliveryNote       As DeliveryNote
    Dim oChasers            As Chasers
    Dim dteLatestChaserDate As Date
    
'***************************************
' Module/Form Name   : DeliveryNotes
'
' Procedure Name     : LatestChaserDate
'
' Purpose            :
'
' Date Created       : 25/05/2006 23:24
'
' Author             : Gareth
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
    On Error GoTo ErrorProc

    dteLatestChaserDate = 0
    For Each oDeliveryNote In mCol
        If oDeliveryNote.Selected Or blnSelectedOnly = False Then
            Set oChasers = New Chasers
            oChasers.DeliveryNoteNo = oDeliveryNote.DNoteNo
            If oChasers.LatestChaserDate > dteLatestChaserDate Then
                dteLatestChaserDate = oChasers.LatestChaserDate
            End If
        End If
    Next oDeliveryNote
    
    LatestChaserDate = dteLatestChaserDate

    On Error GoTo 0
    Exit Function
'
'********* Code Ends Here **************
'
ErrorProc:
    ErrorRaise "DeliveryNotes.LatestChaserDate"
End Function

Public Function LatestNextAction(Optional blnSelectedOnly As Boolean = True) As String
    
    Dim oDeliveryNote       As DeliveryNote
    Dim oChasers            As Chasers
    Dim strLatestNextAction As String
    Dim lngLatestID         As Long
        
'***************************************
' Module/Form Name   : DeliveryNotes
'
' Procedure Name     : LatestNextAction
'
' Purpose            :
'
' Date Created       : 26/05/2006 00:01
'
' Author             : Gareth
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
    On Error GoTo ErrorProc

    lngLatestID = 0
    For Each oDeliveryNote In mCol
        If oDeliveryNote.Selected Or blnSelectedOnly = False Then
            Set oChasers = New Chasers
            oChasers.DeliveryNoteNo = oDeliveryNote.DNoteNo
            If oChasers.LastID > lngLatestID Then
                lngLatestID = oChasers.LastID
                strLatestNextAction = oChasers.LatestNextAction
            End If
        End If
    Next oDeliveryNote
    
    LatestNextAction = strLatestNextAction

    On Error GoTo 0
    Exit Function
'
'********* Code Ends Here **************
'
ErrorProc:
    ErrorRaise "DeliveryNotes.LatestNextAction"
End Function

