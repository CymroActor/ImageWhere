VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SageAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'local variable to hold collection
Private mCol As Collection

Public Function Add(Key As String, _
                    Optional sKey As String) As SageAccount
          'create a new object
          Dim objNewMember As SageAccount
10        Set objNewMember = New SageAccount

          'set the properties passed into the method
20        objNewMember.Key = Key
30        If Len(sKey) = 0 Then
40            mCol.Add objNewMember
50        Else
60            mCol.Add objNewMember, sKey
70        End If


          'return the object created
80        Set Add = objNewMember
90        Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As SageAccount
10      Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
10        Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
10        mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
10        Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
10        Set mCol = New Collection
20        Refresh
End Sub

Private Sub Class_Terminate()
10        Set mCol = Nothing
End Sub

Public Sub Refresh()
          Dim strSQL As String
          Dim rstSageAccounts As DAO.Recordset
          Dim objSageAccount As SageAccount
    
10        On Error GoTo ErrorProc
          '
          '   Clear all existing Accounts.
          '
20        While Not mCol.Count = 0
30            mCol.Remove 1
40        Wend
          '
          '   Refresh Accounts.
          '
50        strSQL = "SELECT ACCOUNT_REF " _
                 & "FROM SALES_LEDGER " _
                 & "ORDER BY ACCOUNT_REF"
    
60        Set rstSageAccounts = db.OpenRecordset(strSQL, dbOpenSnapshot, dbForwardOnly)
70        With rstSageAccounts
80            While Not .EOF
90                Set objSageAccount = Add(!ACCOUNT_REF, _
                                           !ACCOUNT_REF)
100               .MoveNext
110           Wend
120       End With
    
130       rstSageAccounts.Close
140       Set rstSageAccounts = Nothing
    
150       Exit Sub
ErrorProc:
160       Err.Raise Err.Number, "Sage Accounts: Refresh"
End Sub


