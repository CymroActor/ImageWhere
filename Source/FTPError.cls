VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FTPError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mvarErrorNo As Long
Private mvarDescription As String

Public Property Get Description() As String
10        Description = mvarDescription
End Property

Public Sub Read(lngErrorNo As Long)
    
          Dim recFTPError As ADODB.Recordset
          Dim strSQL As String
    
10        mvarErrorNo = lngErrorNo
    
20        strSQL = "SELECT * " _
                 & "FROM FTPERRORS " _
                 & "WHERE ERRORNO = " & CStr(lngErrorNo)
    
30        Set recFTPError = New ADODB.Recordset
40        recFTPError.Open strSQL, gdbADO, adOpenStatic, adLockReadOnly
50        If recFTPError.EOF Then
60            mvarDescription = "*** Unknown Error ***"
70        Else
80            mvarDescription = recFTPError!Description
90        End If
100       recFTPError.Close
110       Set recFTPError = Nothing
  
120       Exit Sub
End Sub


