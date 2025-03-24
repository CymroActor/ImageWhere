VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Customer2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Comments"
Attribute VB_Ext_KEY = "Member1" ,"Contacts"
Option Explicit

Private mvarCustomerNo As Integer 'local copy
Private mvarCustomerName As String 'local copy
Private mvarAddress1 As String 'local copy
Private mvarAddress2 As String 'local copy
Private mvarAddress3 As String 'local copy
Private mvarCountry As String 'local copy
Private mvarTelephone As String 'local copy
Private mvarPostCode As String 'local copy
Private mvarFax As String 'local copy
Private mvarInformation As String 'local copy
Private mvarCountryCode As String 'local copy
Private mvarVATNo As String 'local copy
Private mvarPreferredDelivery As String 'local copy
Private mvarOutstandingPhotos As Integer 'local copy
Private mvarPendingPhotos As Integer 'local copy
Private mvarBusinessType As String 'local copy
Private mvarCountyOrState As String 'local copy
Private mvarSageURN As String 'local copy
Private mvarTermsAndConditionsApproved As Boolean 'local copy
Private mvarOnHold As Boolean 'local copy
Private mvarContacts As Contacts
Private mvarComments As Comments
Private mvarServiceFeeRequired As Boolean
Private mvarServiceFee As Currency
Private mvarDateAmended As Date
'
'   ADO Recordset to hold a checkpointed copy of the data.
'
Private mrstCheckPoint As ADODB.Recordset


Public Property Get Comments() As Comments
    If mvarComments Is Nothing Then
        Set mvarComments = New Comments
        mvarComments.CustomerNo = mvarCustomerNo
    End If

    Set Comments = mvarComments
End Property


Public Property Set Comments(vData As Comments)
    Set mvarComments = vData
End Property

Public Property Get Contacts() As Contacts
    If mvarContacts Is Nothing Then
        Set mvarContacts = New Contacts
        mvarContacts.CustomerNo = mvarCustomerNo
    End If

    Set Contacts = mvarContacts
End Property

Public Property Set Contacts(vData As Contacts)
    Set mvarContacts = vData
End Property

Private Sub Class_Initialize()
    Dim rstSchema As ADODB.Recordset
    Dim fldCust As ADODB.Field
    
    Set mrstCheckPoint = New ADODB.Recordset
    Set rstSchema = gdbADO.OpenSchema(adSchemaColumns, Array(Empty, Empty, "CUSTOMER", Empty))
    While Not rstSchema.EOF
        With mrstCheckPoint
            If rstSchema!DATA_TYPE = adChar Or rstSchema!DATA_TYPE = adWChar Then
                .Fields.Append rstSchema!COLUMN_NAME, rstSchema!DATA_TYPE, -1 ' rstSchema!CHARACTER_MAXIMUM_LENGTH
            Else
                .Fields.Append rstSchema!COLUMN_NAME, rstSchema!DATA_TYPE
            End If
            rstSchema.MoveNext
        End With
    Wend
    With mrstCheckPoint
        .Open
        .AddNew
        For Each fldCust In mrstCheckPoint.Fields
            If fldCust.Type = adChar Or fldCust.Type = adWChar Then
                fldCust.Value = ""
            Else
                fldCust.Value = 0
            End If
        Next fldCust
        .update
    End With
End Sub

Private Sub Class_Terminate()
    Set mvarComments = Nothing
    Set mvarContacts = Nothing
End Sub

Public Sub Create(Optional Force As Boolean)
'***************************************
' Module/Form Name   : Customer2
'
' Procedure Name     : create
'
' Purpose            :
'
' Date Created       : 19/05/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : Force - Boolean
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo create_Error
'
'******** Code Starts Here *************
'
    Dim recCustomer As DAO.Recordset
    Dim recCustomerNo    As DAO.Recordset
    Dim strSQL As String
    Dim strError As String
    Dim strAddress As String
    
    begin_trans
    '
    '   Determine whether there is an existing customer with the same name.
    '
        If Not Force Then
            strSQL = "select customer_name, address_line_1 from customer " _
                & "where customer_name = '" & mvarCustomerName & "';"
            Set recCustomer = db.OpenRecordset(strSQL, dbOpenSnapshot)
            If Not recCustomer.EOF Then
                recCustomer.MoveFirst
                While Not recCustomer.EOF
                    If LCase(Trim(recCustomer!address_line_1)) = LCase(Trim(mvarAddress1)) Then
                        '
                        '   Customer Already Exists.
                        '
                        roll_back
                        strError = "There is already a customer with this name and" & vbCr & "first line of address."
                        Err.Raise vbObjectError + 1, "Customer: Create", strError
                        Exit Sub
                    End If
                    recCustomer.MoveNext
                Wend
                '
                '   Customer Name already exists at a different address.
                '
                recCustomer.MoveFirst
                strAddress = recCustomer!address_line_1
                roll_back
                strError = "Customer '" & mvarCustomerName & "' already exists" & vbCr _
                         & "First line of address: " & strAddress
                Err.Raise vbObjectError + 2, "Customer: Create", strError
                Exit Sub
            End If
        End If
        '
        '   Lock the Customer Number table to stop anyone from updating it while it is being
        '   used.
        '
        Set recCustomerNo = db.OpenRecordset("Customerno", dbOpenTable, dbDenyWrite)
  
        recCustomerNo.MoveFirst
    
        mvarDateAmended = CDate(Now)
        strSQL = "INSERT INTO CUSTOMER "
        strSQL = strSQL & "(CUSTOMER_NO, CUSTOMER_NAME, ADDRESS_LINE_1, ADDRESS_LINE_2, ADDRESS_LINE_3, "
        strSQL = strSQL & "COUNTRY, POST_CODE, TELEPHONE_NO, FAX_NO, INFORMATION, COUNTRY_CODE, VAT_NO, "
        strSQL = strSQL & "PREFERRED_DELIVERY_METHOD, NO_OUTSTANDING_PHOTOS, NO_PENDING_PHOTOS, "
        strSQL = strSQL & "BUSINESS_TYPE, COUNTY_OR_STATE, SAGEURN, SERVICEFEEREQUIRED, SERVICEFEE, DATEAMENDED) "
        strSQL = strSQL & "VALUES (" & CStr(recCustomerNo!next_no) & ",'" & Replace(mvarCustomerName, "'", "''")
        strSQL = strSQL & "','" & JetSQLFixup(mvarAddress1)
        strSQL = strSQL & "','" & JetSQLFixup(mvarAddress2)
        strSQL = strSQL & "','" & JetSQLFixup(mvarAddress3)
        strSQL = strSQL & "','" & JetSQLFixup(mvarCountry)
        strSQL = strSQL & "','" & JetSQLFixup(mvarPostCode)
        strSQL = strSQL & "','" & JetSQLFixup(mvarTelephone)
        strSQL = strSQL & "','" & JetSQLFixup(mvarFax)
        strSQL = strSQL & "','" & JetSQLFixup(mvarInformation)
        strSQL = strSQL & "','" & JetSQLFixup(mvarCountryCode)
        strSQL = strSQL & "','" & JetSQLFixup(mvarVATNo)
        strSQL = strSQL & "','" & JetSQLFixup(mvarPreferredDelivery)
        strSQL = strSQL & "'," & CStr(mvarOutstandingPhotos)
        strSQL = strSQL & "," & CStr(mvarPendingPhotos)
        strSQL = strSQL & ",'" & JetSQLFixup(mvarBusinessType)
        strSQL = strSQL & "','" & JetSQLFixup(mvarCountyOrState)
        strSQL = strSQL & "'," & IIf(mvarSageURN = "", "Null", "'" & JetSQLFixup(mvarSageURN) & "'")
        strSQL = strSQL & "," & CStr(mvarServiceFeeRequired)
        strSQL = strSQL & "," & CStr(mvarServiceFee)
        strSQL = strSQL & ",#" & Format(mvarDateAmended, "MM/DD/YYYY") & "#"
        strSQL = strSQL & ")"
  
        On Error Resume Next
        db.Execute strSQL, dbSeeChanges + dbFailOnError
        If Err.Number = 3022 Then
            Err.Raise 3 + vbObjectError, , "Customer already exists"
        ElseIf Err.Number = 0 Then
            On Error GoTo create_Error
        Else
            ErrorSave
            On Error GoTo create_Error
            ErrorRestore
        End If
        '
        mvarCustomerNo = recCustomerNo!next_no
        '
        '   Update the next customer number to be used.
        '
        With recCustomerNo
            .Edit
            !next_no = !next_no + 1
            .update
        End With
        recCustomerNo.Close
        Set recCustomerNo = Nothing
    commit_trans
    
    Exit Sub
'
'********* Code Ends Here **************
'
   Exit Sub
'
create_Error:
    ErrorRaise "Customer2.create"
End Sub

Public Sub Read(CustomerNo As Integer)
'***************************************
' Module/Form Name   : Customer2
'
' Procedure Name     : Read
'
' Purpose            :
'
' Date Created       : 19/05/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : CustomerNo - Integer
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo Read_Error
'
'******** Code Starts Here *************
'
    
''    Dim recCustomer As ADODB.Recordset
    Dim recCustomer As DAO.Recordset
    Dim strSQL As String
    
    mvarCustomerNo = CustomerNo
    
    strSQL = "SELECT * " _
           & "FROM CUSTOMER " _
           & "WHERE CUSTOMER_NO = " & CStr(mvarCustomerNo)
    
''    Set recCustomer = New ADODB.Recordset
''    recCustomer.Open strSQL, gdbADO, adOpenStatic, , adCmdText

    Set recCustomer = db.OpenRecordset(strSQL, dbOpenSnapshot)
    If recCustomer.EOF Then
        Err.Raise vbObjectError + 8, , "Customer Doesn't Exist"
    Else
        FillProperties recCustomer
    End If
    
    recCustomer.Close
    Set recCustomer = Nothing
'
'********* Code Ends Here **************
'
   Exit Sub
'
Read_Error:
    ErrorRaise "Customer2.Read"
End Sub

Public Sub Locate(Name As String)
    '***************************************
    ' Module/Form Name   : Customer2
    '
    ' Procedure Name     : Locate
    '
    ' Purpose            :
    '
    ' Date Created       : 09/12/2001
    '
    ' Author             : GARETH
    '
    ' Parameters         : Name - String
    '
    ' Amendment History  : Date       Author    Description
    '                    : --------------------------------
    '                    : 09/12/2001 GARETH
    '
    '***************************************
    '
On Error GoTo Locate_Error
    '
    '******** Code Starts Here *************
    '

''    Dim rstCustomer As ADODB.Recordset
    Dim rstCustomer As DAO.Recordset
    Dim strSQL As String
    Dim fCustList As frm_cust_list

''    strSQL = "SELECT * " _
''         & "FROM CUSTOMER " _
''         & "WHERE CUSTOMER_NAME LIKE '%" & JetSQLFixup(Name) & "%'"
    strSQL = "SELECT * " _
         & "FROM CUSTOMER " _
         & "WHERE CUSTOMER_NAME LIKE '*" & JetSQLFixup(Name) & "*'"

''    Set rstCustomer = New ADODB.Recordset
''    rstCustomer.Open strSQL, gdbADO, adOpenStatic, , adCmdText
    Set rstCustomer = db.OpenRecordset(strSQL, dbOpenSnapshot)
    
    With rstCustomer
        If Not .EOF Then
            .MoveLast
        End If
        If .RecordCount = 0 Then
            Err.Raise vbObjectError + 6, "Customer2: Locate", "No customer found"
        ElseIf .RecordCount > 1 Then
            Set fCustList = New frm_cust_list
            fCustList.ProcessCustomer rstCustomer
            If fCustList.Cancel = True Then
                Err.Raise vbObjectError + 7, "Customer2: Locate", "Customer Locate Cancelled"
            End If
            rstCustomer.MoveFirst
''            rstCustomer.Find "Customer_No = " & fCustList.CustomerNo
            rstCustomer.FindFirst "Customer_No = " & fCustList.CustomerNo
            Set fCustList = Nothing
        End If
    End With
    
    FillProperties rstCustomer
    
    CheckPoint
    rstCustomer.Close
    Set rstCustomer = Nothing
    '
    '********* Code Ends Here **************
    '
Exit Sub
    '
Locate_Error:
    ErrorRaise "Customer2.Locate"
End Sub

Public Sub update()
'***************************************
' Module/Form Name   : Customer2
'
' Procedure Name     : update
'
' Purpose            :
'
' Date Created       : 21/07/2004
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo update_Error
'
'******** Code Starts Here *************
'
''    Dim rstCustomer As ADODB.Recordset
    Dim rstCustomer As DAO.Recordset
    Dim strSQL As String
    
    strSQL = "SELECT * FROM CUSTOMER " _
           & "WHERE CUSTOMER_NO = " & mvarCustomerNo
''    Set rstCustomer = New ADODB.Recordset
''    rstCustomer.Open strSQL, gdbADO, adOpenKeyset, adLockPessimistic, adCmdText
    Set rstCustomer = db.OpenRecordset(strSQL, dbOpenDynaset)
    If rstCustomer.RecordCount <> 1 Then
        Err.Raise vbObjectError + 2, , "Customer does not Exist"
    End If
    mvarDateAmended = CDate(Now)
    With rstCustomer
        CheckRecordSame rstCustomer
        .Edit
        !customer_name = mvarCustomerName
        !address_line_1 = mvarAddress1
        !address_line_2 = mvarAddress2
        !address_line_3 = mvarAddress3
        !Country = mvarCountry
        !Telephone_no = mvarTelephone
        !Post_code = mvarPostCode
        !Fax_no = mvarFax
        !Information = mvarInformation
        !VAT_no = mvarVATNo
        !Country_code = mvarCountryCode
        !preferred_delivery_method = mvarPreferredDelivery
        !Business_Type = mvarBusinessType
        !County_Or_State = mvarCountyOrState
        If mvarSageURN = "" Then
            !SageURN = Null
        Else
            !SageURN = mvarSageURN
        End If
        !ServiceFeeRequired = mvarServiceFeeRequired
        !ServiceFee = mvarServiceFee
        !DateAmended = mvarDateAmended
        .update
    End With
    FillSageProperties
    CheckPoint
    
    rstCustomer.Close
    Set rstCustomer = Nothing
'
'********* Code Ends Here **************
'
   Exit Sub
'
update_Error:
    ErrorRaise "Customer2.update"
End Sub

Public Sub Delete()
'***************************************
' Module/Form Name   : Customer2
'
' Procedure Name     : Delete
'
' Purpose            :
'
' Date Created       : 21/07/2004
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo Delete_Error
'
'******** Code Starts Here *************
'
    
''    Dim rstCustomer As ADODB.Recordset
    Dim rstCustomer As DAO.Recordset
    Dim strSQL As String
    Dim customer_name As String
    Dim address_line_1 As String
   
    strSQL = "SELECT * " _
        & "FROM customer " _
        & "WHERE customer_no = " & Str(mvarCustomerNo) & ";"
    
''    Set rstCustomer = New ADODB.Recordset
''    rstCustomer.Open strSQL, gdbADO, adOpenKeyset, adLockPessimistic, adCmdText
    Set rstCustomer = db.OpenRecordset(strSQL, dbOpenDynaset)
    
    If rstCustomer.EOF Then
        MsgBox "Customer not found", vbExclamation, "Delete Customer"
        rstCustomer.Close
        Exit Sub
    End If
    
    With rstCustomer
        If MsgBox("Are you certain you wish to Delete Customer '" & rstCustomer!customer_name + "'" + vbCr _
            & "First line of address: " & rstCustomer!address_line_1, vbYesNo + vbQuestion + vbDefaultButton2, "Delete Customer") = vbNo Then
            rstCustomer.Close
            Exit Sub
        End If
        customer_name = !customer_name
        address_line_1 = !address_line_1
        On Error Resume Next
        .Delete
        If Err.Number = 3200 Then
            MsgBox "Customer has related information e.g. Jobs, Delivery Notes, Contacts, Comments etc." & vbCrLf & "It cannot therefore be deleted", vbInformation, "Delete Customer"
            On Error GoTo Delete_Error
            Exit Sub
        ElseIf Err.Number = 0 Then
            On Error GoTo Delete_Error
        Else
            ErrorSave
            On Error GoTo Delete_Error
            ErrorRestore
        End If
    End With

    MsgBox "Customer '" & customer_name & "'" + vbCr + "First line of address: " & address_line_1 & vbCr _
         & "has been deleted.", vbInformation, "Delete Customer"
    
    Set rstCustomer = Nothing
'
'********* Code Ends Here **************
'
   Exit Sub
'
Delete_Error:
    ErrorRaise "Customer2.Delete"
End Sub

Public Property Let CountyOrState(ByVal vData As String)
    mvarCountyOrState = vData
End Property

Public Property Get CountyOrState() As String
    CountyOrState = mvarCountyOrState
End Property

Public Property Let BusinessType(ByVal vData As String)
    mvarBusinessType = vData
End Property

Public Property Get BusinessType() As String
    BusinessType = mvarBusinessType
End Property

Public Property Let PendingPhotos(ByVal vData As Integer)
    mvarPendingPhotos = vData
End Property

Public Property Get PendingPhotos() As Integer
    PendingPhotos = mvarPendingPhotos
End Property

Public Property Let OutstandingPhotos(ByVal vData As Integer)
    mvarOutstandingPhotos = vData
End Property

Public Property Get OutstandingPhotos() As Integer
    OutstandingPhotos = mvarOutstandingPhotos
End Property

Public Property Let PreferredDelivery(ByVal vData As String)
    mvarPreferredDelivery = vData
End Property

Public Property Get PreferredDelivery() As String
    PreferredDelivery = mvarPreferredDelivery
End Property

Public Property Let VATNo(ByVal vData As String)
    mvarVATNo = vData
End Property

Public Property Get VATNo() As String
    VATNo = mvarVATNo
End Property

Public Property Let CountryCode(ByVal vData As String)
    mvarCountryCode = vData
End Property

Public Property Get CountryCode() As String
    CountryCode = mvarCountryCode
End Property

Public Property Let Information(ByVal vData As String)
    mvarInformation = vData
End Property

Public Property Get Information() As String
    Information = mvarInformation
End Property

Public Property Let Fax(ByVal vData As String)
    mvarFax = vData
End Property

Public Property Get Fax() As String
    Fax = mvarFax
End Property

Public Property Let PostCode(ByVal vData As String)
    mvarPostCode = vData
End Property

Public Property Get PostCode() As String
    PostCode = mvarPostCode
End Property

Public Property Let Telephone(ByVal vData As String)
    mvarTelephone = vData
End Property

Public Property Get Telephone() As String
    Telephone = mvarTelephone
End Property

Public Property Let Country(ByVal vData As String)
    mvarCountry = vData
End Property

Public Property Get Country() As String
    Country = mvarCountry
End Property

Public Property Let Address3(ByVal vData As String)
    mvarAddress3 = vData
End Property

Public Property Get Address3() As String
    Address3 = mvarAddress3
End Property

Public Property Let Address2(ByVal vData As String)
    mvarAddress2 = vData
End Property

Public Property Get Address2() As String
    Address2 = mvarAddress2
End Property

Public Property Let Address1(ByVal vData As String)
    mvarAddress1 = vData
End Property

Public Property Get Address1() As String
    Address1 = mvarAddress1
End Property

Public Property Let CustomerName(ByVal vData As String)
    mvarCustomerName = vData
End Property

Public Property Get CustomerName() As String
    CustomerName = mvarCustomerName
End Property

Public Property Let CustomerNo(ByVal vData As Integer)
    mvarCustomerNo = vData
End Property

Public Property Get CustomerNo() As Integer
    CustomerNo = mvarCustomerNo
End Property

Public Property Get DateAmended() As Date
    DateAmended = mvarDateAmended
End Property

Public Property Let SageURN(ByVal vData As String)
    Dim strSQL As String
    Dim rstSage As DAO.Recordset
    
    mvarSageURN = vData
    If mvarSageURN = "" Then
        mvarTermsAndConditionsApproved = False
        mvarOnHold = False
        Exit Property
    End If
    
    strSQL = "SELECT ACCOUNT_REF, " _
           & "ACCOUNT_ON_HOLD, " _
           & "TERMS_AGREED " _
           & "FROM SALES_LEDGER " _
           & "WHERE ACCOUNT_REF = '" & vData & "'"
    Set rstSage = db.OpenRecordset(strSQL, dbOpenSnapshot)
    If rstSage.EOF Then
        Err.Raise vbObjectError + 3, "Customer: Let SageURN", "URN does not exist on SAGE"
    End If
    mvarTermsAndConditionsApproved = rstSage!Terms_agreed
    mvarOnHold = rstSage!Account_On_Hold
    '
    '   Store the URN on the customer.
    '
'    ws.BeginTrans
'        strSQL = "UPDATE CUSTOMER " _
'               & "SET SAGEURN = '" & vData & "' " _
'               & "WHERE CUSTOMER_NO = " & mvarCustomerNo
'        db.Execute strSQL, dbFailOnError
'        If db.RecordsAffected <> 1 Then
'            Err.Raise vbObjectError + 4, "Customer: Let SageURN", "More than 1 customer with Customer No = " & mvarCustomerNo
'        End If
'    ws.CommitTrans
End Property

Public Property Get SageURN() As String
    SageURN = mvarSageURN
End Property

Public Property Get ServiceFee() As Currency
    ServiceFee = mvarServiceFee
End Property

Public Property Let ServiceFee(vData As Currency)
    mvarServiceFee = vData
End Property

Public Property Get ServiceFeeRequired() As Currency
    ServiceFeeRequired = mvarServiceFeeRequired
End Property

Public Property Let ServiceFeeRequired(vData As Currency)
    mvarServiceFeeRequired = vData
End Property

Public Property Get TermsAndConditionsApproved() As Boolean
    TermsAndConditionsApproved = mvarTermsAndConditionsApproved
End Property

Public Property Get OnHold() As Boolean
    OnHold = mvarOnHold
End Property

''Private Sub FillProperties(rstCustomer As ADODB.Recordset)
Private Sub FillProperties(rstCustomer As DAO.Recordset)
    
    On Error GoTo ErrorProc
    
    With rstCustomer
        mvarCustomerNo = !customer_no
        mvarCustomerName = !customer_name
        mvarAddress1 = !address_line_1
        mvarAddress2 = !address_line_2
        mvarAddress3 = !address_line_3
        mvarCountry = !Country
        mvarTelephone = !Telephone_no
        mvarPostCode = !Post_code
        mvarFax = !Fax_no
        mvarInformation = !Information
        mvarCountryCode = !Country_code
        mvarVATNo = !VAT_no
        mvarPreferredDelivery = !preferred_delivery_method
        mvarBusinessType = !Business_Type
        mvarCountyOrState = !County_Or_State
        mvarSageURN = Format(!SageURN)
        mvarServiceFeeRequired = !ServiceFeeRequired
        mvarServiceFee = NVL(!ServiceFee, NVLNumeric)
        mvarDateAmended = NVL(!DateAmended, NVLNumeric)
    End With
    
    If goSystemConfig.SageLink Then
        FillSageProperties
    End If
    
    Exit Sub
ErrorProc:
    Err.Raise Err.Number, "Customer2: FillProperties"
End Sub

''Private Function CheckRecordSame(rstRecord As ADODB.Recordset) As Boolean
Private Function CheckRecordSame(rstRecord As DAO.Recordset) As Boolean
''    Dim fldCheckpoint As ADODB.Field
    Dim fldCheckpoint As DAO.Field
    
    CheckRecordSame = False
    '
    '   Check if record has changed by comparing it with the data
    '   at last refresh.
    '
    With rstRecord
        For Each fldCheckpoint In rstRecord.Fields
            If fldCheckpoint.Value <> mrstCheckPoint.Fields(fldCheckpoint.Name).Value Then
                Refresh rstRecord
                Err.Raise vbObjectError + 1, "Customer: CheckRecordSame", "Customer has been updated since first obtaining it."
            End If
        Next fldCheckpoint
    End With
    CheckRecordSame = True
End Function

''Private Function Refresh(rstRecord As ADODB.Recordset)
Private Function Refresh(rstRecord As DAO.Recordset)
    With rstRecord
        mvarCustomerNo = !customer_no
        mvarCustomerName = !customer_name
        mvarAddress1 = !address_line_1
        mvarAddress2 = !address_line_2
        mvarAddress3 = !address_line_3
        mvarCountry = !Country
        mvarTelephone = !Telephone_no
        mvarPostCode = !Post_code
        mvarFax = !Fax_no
        mvarInformation = !Information
        mvarCountryCode = !Country_code
        mvarVATNo = !VAT_no
        mvarPreferredDelivery = !preferred_delivery_method
        mvarBusinessType = !Business_Type
        mvarCountyOrState = !County_Or_State
        mvarSageURN = Format(!SageURN)
        mvarServiceFeeRequired = !ServiceFeeRequired
        mvarServiceFee = NVL(!ServiceFee, NVLNumeric)
        mvarDateAmended = !DateAmended
    End With
    CheckPoint
End Function

Private Sub FillSageProperties()
'***************************************
' Module/Form Name   : Customer2
'
' Procedure Name     : FillSageProperties
'
' Purpose            :
'
' Date Created       : 04/04/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo FillSageProperties_Error
'
'******** Code Starts Here *************
'
''    Dim rstSage As ADODB.Recordset
    Dim rstSage As DAO.Recordset
    Dim strSQL As String
    
    If Not goSystemConfig.SageLink Then Exit Sub
    '
    If Format(SageURN) <> "" Then
        strSQL = "SELECT ACCOUNT_REF, " _
               & "ACCOUNT_ON_HOLD, " _
               & "TERMS_AGREED " _
               & "FROM SALES_LEDGER " _
               & "WHERE ACCOUNT_REF = '" & mvarSageURN & "'"
  
''        Set rstSage = New ADODB.Recordset
''        rstSage.Open strSQL, gdbADO, adOpenStatic, , adCmdText
        Set rstSage = db.OpenRecordset(strSQL, dbOpenSnapshot)
        If Not rstSage.EOF Then
            rstSage.MoveLast
            If rstSage.RecordCount <> 1 Then
                Err.Raise vbObjectError + 5, "Customer: Read", "More than 1 Sage Account for ref: " & mvarSageURN
            Else
                rstSage.MoveFirst
            End If
            mvarTermsAndConditionsApproved = rstSage!Terms_agreed
            mvarOnHold = rstSage!Account_On_Hold
        Else
            mvarTermsAndConditionsApproved = False
            mvarOnHold = False
        End If
        rstSage.Close
        Set rstSage = Nothing
    Else
        mvarTermsAndConditionsApproved = False
        mvarOnHold = False
    End If

    Exit Sub
'
'********* Code Ends Here **************
'
   Exit Sub
'
FillSageProperties_Error:
    ErrorRaise "Customer2.FillSageProperties"
End Sub

Public Sub CheckPoint()
    With mrstCheckPoint
        !customer_no = mvarCustomerNo
        !customer_name = mvarCustomerName
        !address_line_1 = mvarAddress1
        !address_line_2 = mvarAddress2
        !address_line_3 = mvarAddress3
        !Country = mvarCountry
        !Telephone_no = mvarTelephone
        !Post_code = mvarPostCode
        !Fax_no = mvarFax
        !Information = mvarInformation
        !Country_code = mvarCountryCode
        !VAT_no = mvarVATNo
        !preferred_delivery_method = mvarPreferredDelivery
        !Business_Type = mvarBusinessType
        !County_Or_State = mvarCountyOrState
        !SageURN = mvarSageURN
        !ServiceFeeRequired = mvarServiceFeeRequired
        !ServiceFee = mvarServiceFee
        !DateAmended = mvarDateAmended
    End With
End Sub

