VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BusinessTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"BusinessType"
Attribute VB_Ext_KEY = "Member0" ,"BusinessType"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Option Compare Text

Private mCol As Collection
Private mvarCurrentName As String

Public Function Add(Name As String, Optional sKey As String) As BusinessType
          'create a new object
          Dim objNewMember As BusinessType
10        Set objNewMember = New BusinessType


          'set the properties passed into the method
20        objNewMember.Name = Name

30        If Len(sKey) = 0 Then
40            mCol.Add objNewMember
50        Else
60            mCol.Add objNewMember, sKey
70        End If


          'return the object created
80        Set Add = objNewMember
90        Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As BusinessType
Attribute Item.VB_UserMemId = 0
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
20        mvarCurrentName = ""
30        Refresh
End Sub

Private Sub Class_Terminate()
10        Set mCol = Nothing
End Sub

Public Sub Refresh()
          Dim strSQL As String
          Dim rstBusinessTypes As ADODB.Recordset
          Dim objBusinessType As BusinessType
          Dim blnCurrentNameFound As Boolean
          Dim strPreviousName As String
    
10        On Error GoTo ErrorProc
          '
          '   Clear all existing Business Types.
          '
20        While Not mCol.Count = 0
30            mCol.Remove 1
40        Wend
          '
          '   Refresh Business Types.
          '
50        strSQL = "SELECT * " _
                 & "FROM BUSINESS_TYPE " _
                 & "ORDER BY BUSINESS_TYPE"
60        Set rstBusinessTypes = New ADODB.Recordset
70        rstBusinessTypes.Open strSQL, gdbADO, adOpenForwardOnly, , adCmdText
80        With rstBusinessTypes
90            strPreviousName = ""
100           blnCurrentNameFound = False
110           While Not .EOF
120               Set objBusinessType = Add(!Business_Type, !Business_Type)
130               With objBusinessType
140                   .InitialReturnPeriod = IIf(NVL(rstBusinessTypes!InitialReturnPeriod, NVLNumeric) = 0, 30, NVL(rstBusinessTypes!InitialReturnPeriod, NVLNumeric))
150                   .SL1ReturnPeriod = IIf(NVL(rstBusinessTypes!SL1ReturnPeriod, NVLNumeric) = 0, 30, NVL(rstBusinessTypes!SL1ReturnPeriod, NVLNumeric))
160                   .SL2ReturnPeriod = IIf(NVL(rstBusinessTypes!SL2ReturnPeriod, NVLNumeric) = 0, 30, NVL(rstBusinessTypes!SL2ReturnPeriod, NVLNumeric))
170                   .Phone1ReturnPeriod = IIf(NVL(rstBusinessTypes!Phone1ReturnPeriod, NVLNumeric) = 0, 30, NVL(rstBusinessTypes!Phone1ReturnPeriod, NVLNumeric))
180                   .Phone2ReturnPeriod = IIf(NVL(rstBusinessTypes!Phone2ReturnPeriod, NVLNumeric) = 0, 30, NVL(rstBusinessTypes!Phone2ReturnPeriod, NVLNumeric))
190                   .LossFeeReturnPeriod = IIf(NVL(rstBusinessTypes!LossFee, NVLNumeric) = 0, 30, NVL(rstBusinessTypes!LossFee, NVLNumeric))
200                   objBusinessType.CheckPoint
210               End With
220               If Not blnCurrentNameFound Then
230                   If !Business_Type > mvarCurrentName Then
240                       If strPreviousName = "" Then
250                           mvarCurrentName = !Business_Type
260                           blnCurrentNameFound = True
270                       Else
280                           mvarCurrentName = strPreviousName
290                           blnCurrentNameFound = True
300                       End If
310                   ElseIf !Business_Type = mvarCurrentName Then
320                       blnCurrentNameFound = True
330                   End If
340               End If
350               strPreviousName = !Business_Type
360               .MoveNext
370           Wend
380       End With
390       If Not blnCurrentNameFound Then
400           mvarCurrentName = strPreviousName
410       End If
    
420       rstBusinessTypes.Close
430       Set rstBusinessTypes = Nothing
440       Exit Sub
ErrorProc:
450       Err.Raise Err.Number, "BusinessTypes: Refresh"
End Sub

Public Property Get CurrentName() As String
10        CurrentName = mvarCurrentName
End Property

Public Property Let CurrentName(vData As String)
10        mvarCurrentName = vData
End Property

