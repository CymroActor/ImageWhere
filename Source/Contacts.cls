VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Contacts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"Contact"
Attribute VB_Ext_KEY = "Member0" ,"Contact"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Private mCol As Collection
Private mintCustomerNo As Integer

Public Sub Refresh()
    '***************************************
    ' Module/Form Name   : Contacts
    '
    ' Procedure Name     : Refresh
    '
    ' Purpose            :
    '
    ' Date Created       : 09/12/2001
    '
    ' Author             : GARETH
    '
    ' Amendment History  : Date       Author    Description
    '                    : --------------------------------
    '                    : 09/12/2001 GARETH
    '
    '***************************************
    '
On Error GoTo Refresh_Error
    '
    '******** Code Starts Here *************
    '
    Dim strSQL As String
    Dim recContacts As DAO.Recordset
    Dim objContact As Contact

    '
    '   Clear all existing Contacts.
    '
Set mCol = Nothing
Set mCol = New Collection
    '
    '   Refresh Contacts Collection.
    '
strSQL = "SELECT * FROM CONTACTS WHERE CUSTOMER_NO = " & mintCustomerNo & " ORDER BY NAME"
    'Set recContacts = New ADODB.Recordset
    'recContacts.Open strSQL, gdbADO, adOpenStatic, , adCmdText
Set recContacts = db.OpenRecordset(strSQL, dbOpenSnapshot, dbForwardOnly)
With recContacts
  While Not recContacts.EOF


      Set objContact = Add(!ID, _
                           !MailingList, _
                           !customer_no, _
                           NVL(!Name, NVLString), _
                           NVL(!Position, NVLString), _
                           NVL(!Comments, NVLString), _
                           NVL(!Phone, NVLString), _
                           NVL(!Fax, NVLString), _
                           NVL(!Email, NVLString), _
                           NVL(!WebUser, NVLNumeric), _
                           CStr(!ID))
      objContact.DateAmended = NVL(!DateAmended, NVLNumeric)
      objContact.PreferredPhotoType = NVL(!PreferredPhotoType, NVLString)
      objContact.CheckPoint
      .MoveNext
  Wend
End With

recContacts.Close
Set recContacts = Nothing
    '
    '********* Code Ends Here **************
    '
Exit Sub
    '
Refresh_Error:
ErrorRaise "Contacts.Refresh"
End Sub
Public Function Add(ID, _
                    Mail As Boolean, _
                    CustomerNo As Integer, _
                    Name As String, _
                    Position As String, _
                    Comments As String, _
                    Phone As String, _
                    Fax As String, _
                    Email As String, _
                    WebUser As Boolean, _
                    Optional sKey As String) As Contact
    '***************************************
    ' Module/Form Name   : Contacts
    '
    ' Procedure Name     : Add
    '
    ' Purpose            :
    '
    ' Date Created       : 09/12/2001
    '
    ' Author             : GARETH
    '
    ' Parameters         : mail - Boolean
    '                    : CustomerNo - Integer
    '                    : Name - String
    '                    : Position - String
    '                    : Comments - String
    '                    : Phone - String
    '                    : Fax - String
    '                    : Email - String
    '                    : sKey - String
    '
    ' Returns            : Contact
    '
    ' Amendment History  : Date       Author    Description
    '                    : --------------------------------
    '                    : 09/12/2001 GARETH
    '
    '***************************************
    '
On Error GoTo Add_Error
    '
    '******** Code Starts Here *************
    '

    'create a new object
    Dim objNewMember As Contact
    Set objNewMember = New Contact
    
    'set the properties passed into the method
    objNewMember.ID = ID
    objNewMember.Mail = Mail
    objNewMember.CustomerNo = CustomerNo
    objNewMember.Name = Name
    objNewMember.Position = Position
    objNewMember.Comments = Comments
    objNewMember.Phone = Phone
    objNewMember.Fax = Fax
    objNewMember.Email = Email
    objNewMember.WebUser = WebUser
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If
        '
        '   Take a checkpoint.
        '
        'objNewMember.CheckPoint
        'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing
    
        '
        '********* Code Ends Here **************
        '
    Exit Function
    '
Add_Error:
ErrorRaise "Contacts.Add"
End Function

Public Property Get Item(vntIndexKey As Variant) As Contact
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

Public Property Get CustomerNo() As Integer
    CustomerNo = mintCustomerNo
End Property

Public Property Let CustomerNo(ByVal vNewValue As Integer)
  
    mintCustomerNo = vNewValue
    Refresh

End Property


