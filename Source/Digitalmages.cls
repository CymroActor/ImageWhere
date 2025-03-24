VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DigitalImages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"DigitalImage"
Attribute VB_Ext_KEY = "Member0" ,"DigitalImage"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Option Explicit

'local variable to hold collection
Private mCol                As Collection
Private mvarBatchNo         As Long

Public Property Get BatchNo() As Long
    BatchNo = mvarBatchNo
End Property

Public Property Let BatchNo(vData As Long)
    mvarBatchNo = vData
End Property

Public Function Add(BatchNo As Long, _
                    Description As String, _
                    Selected As Boolean, _
                    DigitalImageIndex As Integer, _
                    Optional sKey As String) As DigitalImage
    'create a new object
    Dim objNewMember As DigitalImage
    Set objNewMember = New DigitalImage


    'set the properties passed into the method
    objNewMember.BatchNo = BatchNo
    objNewMember.Description = Description
    objNewMember.Selected = Selected
    objNewMember.DigitalImageIndex = DigitalImageIndex
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As DigitalImage
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
    Dim strSQL As String
    Dim recDigitalImages As DAO.Recordset
    Dim objDigitalImage  As DigitalImage
    '
    '   Clear all existing Digital Images.
    '
'***************************************
' Module/Form Name   : DigitalImages
'
' Procedure Name     : Refresh
'
' Purpose            :
'
' Date Created       : 13/06/2006 00:09
'
' Author             : Gareth
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
    On Error GoTo ErrorProc

    Set mCol = Nothing
    Set mCol = New Collection
    '
    '   Refresh Digital Images Collection.
    '
    If mvarBatchNo <> 0 Then
        strSQL = "SELECT batch.batch_no, Search_result.delivery_note_no, Customer.customer_no, "
        strSQL = strSQL & "Customer.customer_name, Customer.address_line_1, Search_result.DigitalStatus "
        strSQL = strSQL & "FROM Customer INNER JOIN (batch INNER JOIN Search_result ON batch.batch_no = Search_result.batch_no) ON Customer.customer_no = Search_result.customer_no "
        strSQL = strSQL & "WHERE batch.batch_no = " & CStr(mvarBatchNo) & " And search_result.photograph_no = 0"

        Set recDigitalImages = db.OpenRecordset(strSQL, dbOpenSnapshot, dbForwardOnly)
        With recDigitalImages
            While Not recDigitalImages.EOF
      
                Set objDigitalImage = Add(!batch_no, "", False, 0)
                objDigitalImage.DeliveryNoteNo = NVL(!delivery_note_no, NVLNumeric)
                objDigitalImage.CustomerNo = NVL(!customer_no, NVLNumeric)
                objDigitalImage.CustomerName = !customer_name & ""
                objDigitalImage.AddressLine1 = !address_line_1 & ""
                objDigitalImage.DigitalStatus = !DigitalStatus & ""
                .MoveNext
            Wend
        End With

        recDigitalImages.Close
        Set recDigitalImages = Nothing
    End If

    On Error GoTo 0
    Exit Sub
'
'********* Code Ends Here **************
'
ErrorProc:
    ErrorRaise "DigitalImages.Refresh"
End Sub


