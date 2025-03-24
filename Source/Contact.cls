VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Contact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"Comments"
Option Explicit

Private mvarId                      As Long
Private mvarMail                    As Boolean
Private mvarCustomerNo              As Integer
Private mvarName                    As String
Private mvarPosition                As String
Private mvarComments                As String
Private mvarPhone                   As String
Private mvarFax                     As String
Private mvarEmail                   As String
Private mvarDateAmended             As Date
Private mvarWebUser                 As Boolean
Private mvarPreferredPhotoType      As String

'Private mrstCheckPoint As ADODB.Recordset
Private mrecContact As DAO.Recordset
Private Type udtCheckPoint
    ID                  As Long
    CustomerNo          As Integer
    Name                As String
    Position            As String
    Comments            As String
    Phone               As String
    Fax                 As String
    Email               As String
    DateAmended         As Date
    Mail                As Boolean
    WebUser             As Boolean
    PreferredPhotoType  As String
End Type

Private ContactCheckPoint As udtCheckPoint

Public Sub Delete()
    
    Dim SQL As String
    
    On Error GoTo ErrorProc
    '
    '   Read contact and lock before deleting.
    '
    SQL = "SELECT * " _
        & "FROM CONTACTS " _
        & "WHERE ID = " & mvarId
    
    'Set recContact = New ADODB.Recordset
    'recContact.Open SQL, gdbADO, adOpenDynamic, adLockPessimistic, adCmdText
    Set mrecContact = db.OpenRecordset(SQL, dbOpenDynaset, dbPessimistic)
    mrecContact.Edit
    CheckRecordSame
    mrecContact.Delete
    MsgBox "Contact Deleted", vbInformation
    Exit Sub
ErrorProc:
    Err.Raise Err.Number, "Contact: Delete"
End Sub

Public Property Get DateAmended() As Date
    DateAmended = mvarDateAmended
End Property

Public Property Let DateAmended(ByVal vData As Date)
    mvarDateAmended = vData
End Property

Public Property Let Email(ByVal vData As String)
    mvarEmail = vData
End Property

Public Property Get Email() As String
    Email = mvarEmail
End Property

Public Property Let Fax(ByVal vData As String)
    mvarFax = vData
End Property

Public Property Get Fax() As String
    Fax = mvarFax
End Property

Public Property Let Phone(ByVal vData As String)
    mvarPhone = vData
End Property

Public Property Get Phone() As String
    Phone = mvarPhone
End Property

Public Property Let Comments(ByVal vData As String)
    mvarComments = vData
End Property

Public Property Get Comments() As String
    Comments = mvarComments
End Property

Public Property Let Position(ByVal vData As String)
    mvarPosition = vData
End Property

Public Property Get Position() As String
    Position = mvarPosition
End Property

Public Property Let PreferredPhotoType(ByVal vData As String)
    mvarPreferredPhotoType = vData
End Property

Public Property Get PreferredPhotoType() As String
    If mvarPreferredPhotoType = "" Then
        PreferredPhotoType = "T"
    Else
        PreferredPhotoType = mvarPreferredPhotoType
    End If
End Property

Public Sub Add(CustomerNo As Integer, Name As String)
'***************************************
' Module/Form Name   : Contact
'
' Procedure Name     : Add
'
' Purpose            :
'
' Date Created       : 05/07/2004
'
' Author             : GARETH SAUNDERS
'
' Parameters         : CustomerNo - Integer
'                    : Name - String
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo Add_Error
'
'******** Code Starts Here *************
'
    Dim strSQL          As String
    Dim rstContact      As DAO.Recordset
    
    mvarCustomerNo = CustomerNo
    mvarName = Name
    mvarDateAmended = CDate(Now)
    Validate
''    strSQL = "INSERT INTO CONTACTS " _
''           & "(CUSTOMER_NO, NAME, POSITION, COMMENTS, PHONE, FAX, EMAIL, DATEAMENDED, MAILINGLIST, WEBUSER, PREFERREDPHOTOTYPE) " _
''           & "VALUES (" & CStr(CustomerNo) & ",'" & Replace(Name, "'", "''") _
''           & "','" & Replace(mvarPosition, "'", "''") _
''           & "','" & Replace(mvarComments, "'", "''") _
''           & "','" & Replace(mvarPhone, "'", "''") _
''           & "','" & Replace(mvarFax, "'", "''") _
''           & "','" & Replace(mvarEmail, "'", "''") _
''           & "',#" & Format(mvarDateAmended, "mm/dd/yyyy") & "#" _
''           & "," & CStr(mvarMail) _
''           & "," & CStr(mvarWebUser) _
''           & ",'" & Replace(mvarPreferredPhotoType, "'", "''") & "')"
''
''    db.Execute strSQL, dbFailOnError
    

    Set rstContact = db.OpenRecordset("Contacts", dbOpenTable)
''    rstContact.Open "Contacts", gdbADO, adOpenKeyset, adLockPessimistic

    With rstContact
        .AddNew
        !customer_no = CStr(CustomerNo)
        !Name = Replace(Name, "'", "''")
        !Position = Replace(mvarPosition, "'", "''")
        !Comments = Replace(mvarComments, "'", "''")
        !Phone = Replace(mvarPhone, "'", "''")
        !Fax = Replace(mvarFax, "'", "''")
        !Email = Replace(mvarEmail, "'", "''")
        !DateAmended = mvarDateAmended
        !MAILINGLIST = mvarMail
        !WebUser = mvarWebUser
        !PreferredPhotoType = mvarPreferredPhotoType
        .Update
        .Bookmark = .LastModified
        mvarId = !ID
    End With
    '
    rstContact.Close
    Set rstContact = Nothing
'
'********* Code Ends Here **************
'
   Exit Sub
'
Add_Error:
    ErrorRaise "Contact.Add"
End Sub

Public Property Get Name() As String
    Name = mvarName
End Property

Public Property Let Name(vData As String)
    mvarName = vData
End Property

Public Property Get CustomerNo() As Integer
    CustomerNo = mvarCustomerNo
End Property

Public Property Let CustomerNo(vData As Integer)
    mvarCustomerNo = vData
End Property

Public Sub Update()
'***************************************
' Module/Form Name   : Contact
'
' Procedure Name     : update
'
' Purpose            :
'
' Date Created       : 04/04/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 04/04/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo Update_Error
'
'******** Code Starts Here *************
'
    Dim strSQL As String

''    strSQL = "SELECT * FROM CONTACTS " _
''           & "WHERE CUSTOMER_NO = " & CStr(mvarCustomerNo) & " AND " _
''           & "NAME = '" & Replace(mvarName, "'", "''") & "'"
    strSQL = "SELECT * FROM CONTACTS " _
           & "WHERE ID = " & CStr(mvarId)

    'Set recContact = New ADODB.Recordset
    Set mrecContact = db.OpenRecordset(strSQL, dbOpenDynaset, dbSeeChanges)

    If mrecContact.EOF Then
        Err.Raise 1, , "Contact does not Exist"
    End If
    With mrecContact
        mvarDateAmended = CDate(Now)
        CheckRecordSame
        Validate
        .Edit
        !Name = mvarName
        !Position = mvarPosition
        !Comments = mvarComments
        !Phone = mvarPhone
        !Fax = mvarFax
        !Email = mvarEmail
        !DateAmended = mvarDateAmended
        !MAILINGLIST = mvarMail
        !WebUser = mvarWebUser
        !PreferredPhotoType = mvarPreferredPhotoType
        .Update
        CheckPoint
    End With

    mrecContact.Close
    '
    '********* Code Ends Here **************
    '
    Exit Sub
    '
Update_Error:
    ErrorRaise "Contact.update"
End Sub

Public Sub Read(ByVal ID As Long)
    
    'Dim recContact As dao.Recordset
    Dim strSQL As String
    
    On Error GoTo ErrorProc
    
''    mvarName = Name
''    mvarCustomerNo = CustomerNo
    
''    strSQL = "SELECT * " _
''           & "FROM CONTACTS " _
''           & "WHERE CUSTOMER_NO = " & CStr(mvarCustomerNo) & " AND " _
''           & "NAME = '" & Replace(Name, "'", "''") & "'"
    
    strSQL = "SELECT * " _
           & "FROM CONTACTS " _
           & "WHERE ID = " & CStr(ID)
    
'    Set mrecContact = New ADODB.Recordset
'    mrecContact.Open strSQL, gdbADO, adOpenStatic, , adCmdText
    Set mrecContact = db.OpenRecordset(strSQL, dbOpenSnapshot)
    If mrecContact.EOF Then
        Err.Raise vbObjectError + 5, , "Contact '" & Name & "' does not exist"
    Else
        Refresh
    End If
    CheckPoint
    mrecContact.Close
    Set mrecContact = Nothing
  
    Exit Sub

ErrorProc:
    ErrorRaise "Contact: Read"
End Sub

Public Property Let Mail(ByVal vData As Boolean)
    mvarMail = vData
End Property

Public Property Get Mail() As Boolean
    Mail = mvarMail
End Property

Public Property Let WebUser(ByVal vData As Boolean)
    mvarWebUser = vData
End Property

Public Property Get WebUser() As Boolean
    WebUser = mvarWebUser
End Property

Public Property Let ID(ByVal vData As Long)
    mvarId = vData
End Property

Public Property Get ID() As Long
    ID = mvarId
End Property

'Private Function CheckRecordSame(recRecord As ADODB.Recordset) As Boolean
'    Dim fldCheckpoint As ADODB.Field
'
'    CheckRecordSame = False
'    '
'    '   Check if record has changed by comparing it with the data
'    '   at last refresh.
'    '
'    With recRecord
'        For Each fldCheckpoint In .Fields
'            If fldCheckpoint.Value <> mrstCheckPoint.Fields(fldCheckpoint.Name).Value Then
'                Refresh recRecord
'                Err.Raise vbObjectError + 1, "Contact: CheckRecordSame", "Contact has been updated since first obtaining it."
'            End If
'        Next fldCheckpoint
'    End With
'    CheckRecordSame = True
'End Function

Private Function CheckRecordSame() As Boolean
    '***************************************
    ' Module/Form Name   : Contact
    '
    ' Procedure Name     : CheckRecordSame
    '
    ' Purpose            :
    '
    ' Date Created       : 04/04/2002
    '
    ' Author             : GARETH SAUNDERS
    '
    ' Returns            : Boolean
    '
    ' Amendment History  : Date       Author    Description
    '                    : --------------------------------
    '                    : 04/04/2002 GARETH SAUNDERS
    '
    '***************************************
    '
On Error GoTo CheckRecordSame_Error
    '
    '******** Code Starts Here *************
    '

CheckRecordSame = False
    '
    '   Check if record has changed by comparing it with the data
    '   at last refresh.
    '
With ContactCheckPoint
  If mrecContact!customer_no <> .CustomerNo Or _
     mrecContact!Name <> .Name Or _
     mrecContact!Position <> .Position Or _
     mrecContact!Comments <> .Comments Or _
     mrecContact!Phone <> .Phone Or _
     mrecContact!Fax <> .Fax Or _
     mrecContact!Email <> .Email Or _
     mrecContact!DateAmended <> .DateAmended Or _
     mrecContact!MAILINGLIST <> .Mail Or _
     mrecContact!WebUser <> .WebUser Then
      Refresh
      Err.Raise vbObjectError + 1, "Contact: CheckRecordSame", "Contact has been updated since first obtaining it."
  End If
End With
CheckRecordSame = True
    '
    '********* Code Ends Here **************
    '
Exit Function
    '
CheckRecordSame_Error:
    ErrorRaise "Contact.CheckRecordSame"
End Function

'Private Function Refresh(recRecord As ADODB.Recordset)
'    mvarPosition = Format(recRecord!Position)
'    mvarComments = Format(recRecord!Comments)
'    mvarPhone = Format(recRecord!Phone)
'    mvarFax = Format(recRecord!Fax)
'    mvarEmail = Format(recRecord!Email)
'    mvarMail = Format(recRecord!MailingList)
'End Function

Private Function Refresh()
'Private Function Refresh(rstContact As dao.Recordset)
    With mrecContact
        mvarId = NVL(!ID, NVLNumeric)
        mvarCustomerNo = NVL(!customer_no, NVLNumeric)
        mvarName = NVL(!Name, NVLString)
        mvarPosition = NVL(!Position)
        mvarComments = NVL(!Comments)
        mvarPhone = NVL(!Phone)
        mvarFax = NVL(!Fax)
        mvarEmail = NVL(!Email)
        mvarDateAmended = NVL(!DateAmended, NVLNumeric)
        mvarMail = NVL(!MAILINGLIST)
        mvarWebUser = NVL(!WebUser)
        mvarPreferredPhotoType = NVL(!PreferredPhotoType, NVLString)
    End With
End Function

Public Sub CheckPoint()
    On Error GoTo ErrorProc
    With ContactCheckPoint
        
        .CustomerNo = mvarCustomerNo
        .Name = mvarName
        .Position = mvarPosition
        .Comments = mvarComments
        .Phone = mvarPhone
        .Fax = mvarFax
        .Email = mvarEmail
        .DateAmended = mvarDateAmended
        .Mail = mvarMail
        .WebUser = mvarWebUser
        .PreferredPhotoType = mvarPreferredPhotoType
    End With
    Exit Sub
ErrorProc:
    ErrorRaise "Contact: Checkpoint"
End Sub

'Private Sub Class_Initialize()
'    Dim rstSchema As ADODB.Recordset
'    Dim fldContact As ADODB.Field
'
'    Set mrstCheckPoint = New ADODB.Recordset
'    Set rstSchema = gdbADO.OpenSchema(adSchemaColumns, Array(Empty, Empty, "CONTACTS", Empty))
'    While Not rstSchema.EOF
'        With mrstCheckPoint
'            If rstSchema!DATA_TYPE = adChar Or rstSchema!DATA_TYPE = adWChar Then
'                .Fields.Append rstSchema!COLUMN_NAME, rstSchema!DATA_TYPE, -1 ' rstSchema!CHARACTER_MAXIMUM_LENGTH
'            Else
'                .Fields.Append rstSchema!COLUMN_NAME, rstSchema!DATA_TYPE
'            End If
'            rstSchema.MoveNext
'        End With
'    Wend
'    With mrstCheckPoint
'        .Open
'        .AddNew
'        For Each fldContact In mrstCheckPoint.Fields
'            If fldContact.Type = adChar Or fldContact.Type = adWChar Then
'                fldContact.Value = ""
'            Else
'                fldContact.Value = 0
'            End If
'        Next fldContact
'        .update
'    End With
'End Sub

Private Sub Validate()
    '***************************************
    ' Module/Form Name   : Contact
    '
    ' Procedure Name     : Validate
    '
    ' Purpose            :
    '
    ' Date Created       : 04/04/2002
    '
    ' Author             : GARETH SAUNDERS
    '
    ' Amendment History  : Date       Author    Description
    '                    : --------------------------------
    '                    : 04/04/2002 GARETH SAUNDERS
    '
    '***************************************
    '
    On Error GoTo Validate_Error
    '
    '******** Code Starts Here *************
    '
    '
    '   A Web User must have an email address.
    '
    If mvarWebUser Then
        If Trim(mvarEmail) = "" Then
            Err.Raise vbObjectError + 3, , "A Web User must have an Email Address."
        End If
        '
    End If
    '
    '   Validate the email address as best we can.
    '
    If Trim(mvarEmail) <> "" Then
        If InStr(mvarEmail, "@") = 0 Or _
           Left(mvarEmail, 1) = "@" Or _
           Right(mvarEmail, 1) = "@" Then
            Err.Raise vbObjectError + 4, , "Please enter a valid Email Address with a '@' in the middle."
        End If
    End If
    '
    '********* Code Ends Here **************
    '
    Exit Sub
    '
Validate_Error:
    ErrorRaise "Contact.Validate"
End Sub
