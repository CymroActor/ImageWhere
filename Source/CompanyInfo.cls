VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CompanyInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'local variable(s) to hold property value(s)
Private mvarSignatory As String 'local copy
Private mvarCompanyName As String
Private mvarAddress1 As String
Private mvarAddress2 As String
Private mvarAddress3 As String
Private mvarTown As String
Private mvarCounty As String
Private mvarCountry As String
Private mvarPostCode As String
Private mvarTelNo As String
Private mvarFaxNo As String
Private mvarVATNo As String
Private mvarEmail As String
Private mvarWebSite As String
Private mvarInfo1 As String
Private mvarInfo2 As String

Public Property Let Signatory(ByVal vData As String)
10        mvarSignatory = vData
End Property

Public Property Get Signatory() As String
10        Signatory = mvarSignatory
End Property

Public Property Get CompanyName() As String
10        CompanyName = Trim(mvarCompanyName)
End Property

Public Property Let CompanyName(ByVal vData As String)
10        vData = Trim(vData) & " "
20        mvarCompanyName = vData
End Property

Public Property Let Address1(ByVal vData As String)
10        vData = Trim(vData) & " "
20        mvarAddress1 = vData
End Property

Public Property Get Address1() As String
10        Address1 = Trim(mvarAddress1)
End Property

Public Property Let Address2(ByVal vData As String)
10        vData = Trim(vData) & " "
20        mvarAddress2 = vData
End Property

Public Property Get Address2() As String
10        Address2 = Trim(mvarAddress2)
End Property

Public Property Let Address3(ByVal vData As String)
10        vData = Trim(vData) & " "
20        mvarAddress3 = vData
End Property

Public Property Get Address3() As String
10        Address3 = Trim(mvarAddress3)
End Property

Public Property Let Town(ByVal vData As String)
10        vData = Trim(vData) & " "
20        mvarTown = vData
End Property

Public Property Get Town() As String
10        Town = Trim(mvarTown)
End Property

Public Property Let County(ByVal vData As String)
10        vData = Trim(vData) & " "
20        mvarCounty = vData
End Property

Public Property Get County() As String
10        County = Trim(mvarCounty)
End Property

Public Property Let Country(ByVal vData As String)
10        vData = Trim(vData) & " "
20        mvarCountry = vData
End Property

Public Property Get Country() As String
10        Country = Trim(mvarCountry)
End Property

Public Property Let PostCode(ByVal vData As String)
10        vData = Trim(vData) & " "
20        mvarPostCode = vData
End Property

Public Property Get PostCode() As String
10        PostCode = Trim(mvarPostCode)
End Property

Public Property Let TelNo(ByVal vData As String)
10        vData = Trim(vData) & " "
20        mvarTelNo = vData
End Property

Public Property Get TelNo() As String
10        TelNo = Trim(mvarTelNo)
End Property

Public Property Let FaxNo(ByVal vData As String)
10        vData = Trim(vData) & " "
20        mvarFaxNo = vData
End Property

Public Property Get FaxNo() As String
10        FaxNo = Trim(mvarFaxNo)
End Property

Public Property Let VATNo(ByVal vData As String)
10        vData = Trim(vData) & " "
20        mvarVATNo = vData
End Property

Public Property Get VATNo() As String
10        VATNo = Trim(mvarVATNo)
End Property

Public Property Let Email(ByVal vData As String)
10        vData = Trim(vData) & " "
20        mvarEmail = vData
End Property

Public Property Get Email() As String
10        Email = Trim(mvarEmail)
End Property

Public Property Let WebSite(ByVal vData As String)
10        vData = Trim(vData) & " "
20        mvarWebSite = vData
End Property

Public Property Get WebSite() As String
10        WebSite = Trim(mvarWebSite)
End Property

Public Property Let Info1(ByVal vData As String)
10        vData = Trim(vData) & " "
20        mvarInfo1 = vData
End Property

Public Property Get Info1() As String
10        Info1 = Trim(mvarInfo1)
End Property

Public Property Let Info2(ByVal vData As String)
10        vData = Trim(vData) & " "
20        mvarInfo2 = vData
End Property

Public Property Get Info2() As String
10        Info2 = Trim(mvarInfo2)
End Property
''
''Private Sub Class_Initialize()
''10        On Error GoTo ErrorProc
''20        Refresh
''
''30        Exit Sub
''ErrorProc:
''40        ErrorRaise "CompanyInfo: Initialise"
''End Sub

Public Sub update()
          '***************************************
          ' Module/Form Name   : CompanyInfo
          '
          ' Procedure Name     : update
          '
          ' Purpose            :
          '
          ' Date Created       : 08/06/2002
          '
          ' Author             : GARETH SAUNDERS
          '
          ' Amendment History  : Date       Author    Description
          '                    : --------------------------------
          '                    : 08/06/2002 GARETH SAUNDERS
          '
          '***************************************
          '
10    On Error GoTo update_Error
          '
          '******** Code Starts Here *************
          '
          Dim strSQL As String
          Dim rstCompanyInfo As DAO.Recordset

20    strSQL = "SELECT * FROM COMPANYINFO"
30    Set rstCompanyInfo = db.OpenRecordset(strSQL, dbOpenDynaset)
40    With rstCompanyInfo
50      .Edit
60      !CompanyName = Trim(mvarCompanyName)
70      !Signatory = mvarSignatory
80      !Address1 = mvarAddress1
90      !Address2 = mvarAddress2
100     !Address3 = mvarAddress3
110     !Town = mvarTown
120     !County = mvarCounty
130     !Country = mvarCountry
140     !PostCode = mvarPostCode
150     !Telephone = mvarTelNo
160     !Fax = mvarFaxNo
170     !VATNo = mvarVATNo
180     !Email = mvarEmail
190     !WebAddress = mvarWebSite
200     !Info1 = mvarInfo1
210     !Info2 = mvarInfo2
220     .update
230   End With
240   Set rstCompanyInfo = Nothing
          '
          '********* Code Ends Here **************
          '
250   Exit Sub
          '
update_Error:
260       ErrorRaise "CompanyInfo.update"
End Sub

Public Sub Refresh()
          '***************************************
          ' Module/Form Name   : CompanyInfo
          '
          ' Procedure Name     : Refresh
          '
          ' Purpose            :
          '
          ' Date Created       : 08/06/2002
          '
          ' Author             : GARETH SAUNDERS
          '
          ' Amendment History  : Date       Author    Description
          '                    : --------------------------------
          '                    : 08/06/2002 GARETH SAUNDERS
          '
          '***************************************
          '
10    On Error GoTo Refresh_Error
          '
          '******** Code Starts Here *************
          '
          Dim strSQL As String
          Dim rstCompanyInfo As DAO.Recordset

20    strSQL = "SELECT * FROM COMPANYINFO"
30    Set rstCompanyInfo = db.OpenRecordset(strSQL, dbOpenSnapshot)

40    With rstCompanyInfo
50      If .EOF Then
60          Err.Raise vbObjectError + 1, "CompanyInfo: Initialize", "Company Information Record Missing"
70      Else
80          .MoveLast
90          .MoveFirst
100         If .RecordCount > 1 Then
110             Err.Raise vbObjectError + 2, "CompanyInfo: Initialize", "More than one Company Information record found"
120         Else
130             .MoveFirst
140         End If
150     End If

160     mvarSignatory = Format(!Signatory)
170     mvarCompanyName = Format(!CompanyName)
180     mvarAddress1 = Format(!Address1)
190     mvarAddress2 = Format(!Address2)
200     mvarAddress3 = Format(!Address3)
210     mvarTown = Format(!Town)
220     mvarCounty = Format(!County)
230     mvarCountry = Format(!Country)
240     mvarPostCode = Format(!PostCode)
250     mvarTelNo = Format(!Telephone)
260     mvarFaxNo = Format(!Fax)
270     mvarVATNo = Format(!VATNo)
280     mvarEmail = Format(!Email)
290     mvarWebSite = Format(!WebAddress)
300     mvarInfo1 = Format(!Info1)
310     mvarInfo2 = Format(!Info2)
320   End With

330   rstCompanyInfo.Close
          '
          '********* Code Ends Here **************
          '
340   Exit Sub
          '
Refresh_Error:
350   ErrorRaise "CompanyInfo.Refresh"
End Sub
