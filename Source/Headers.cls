VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GridColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"Column"
Attribute VB_Ext_KEY = "Member0" ,"Column"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable to hold collection
Private mCol As Collection
Private msgGrid As MSFlexGrid
'local variable(s) to hold property value(s)
Private mvarGrid As Object 'local copy

Friend Property Set Grid(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Grid = Form1
    Set mvarGrid = vData
End Property


Friend Property Get Grid() As Object
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Grid
    Set Grid = mvarGrid
End Property

Public Function Add(Key As String, _
                    Header As String, _
                    Value As Variant) As GridColumn
    'create a new object
    Dim objNewMember As GridColumn
    Set objNewMember = New GridColumn


    'set the properties passed into the method
    Set objNewMember.Grid = mvarGrid
    objNewMember.Key = Key
    objNewMember.Header = Header
    objNewMember.Value = Value
    mCol.Add objNewMember, Key

    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As GridColumn
Attribute Item.VB_UserMemId = 0
'    Refresh
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
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

Public Sub Refresh()
    Dim intCount As Integer
    
    Set mCol = Nothing
    Set mCol = New Collection
    
    For intCount = 1 To mvarGrid.Cols
        Add CStr(intCount), _
            mvarGrid.TextMatrix(0, intCount - 1), _
            mvarGrid.TextMatrix(IIf(mvarGrid.Rows = 1, 0, mvarGrid.row), intCount - 1)
    Next intCount
        
End Sub
