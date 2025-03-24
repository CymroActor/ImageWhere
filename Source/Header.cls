VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GridColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit
Public Enum SortedTypes
    smgNone = 0
    smgAscending = 1
    smgDescending = 2
End Enum

Private mstrKey As String

'local variable(s) to hold property value(s)
Private mvarHeader As String 'local copy
Private mvarGrid As MSFlexGrid 'local copy
Private mintWidth As Integer
Private mvarValue As Variant
Private mvarSorted As SortedTypes
Private mvarHeight As Integer

Public Property Let Header(ByVal vData As String)
    mvarHeader = vData
    mvarGrid.TextArray(fgIndex(mvarGrid, 0, CInt(Key) - 1)) = mvarHeader
'    With mvarGrid
'        .col = CInt(Key) - 1
'        .ColSel = .col
'        .row = .RowSel
'        .CellFontBold = True
'    End With
        
End Property


Public Property Get Header() As String
    Header = mvarHeader
End Property

Friend Property Set Grid(ByVal vData As Object)
    Set mvarGrid = vData
End Property

Public Property Get Width() As Integer
    Width = mvarGrid.ColWidth(CInt(Key) - 1)
End Property

Public Property Let Width(vData As Integer)
    mvarGrid.ColWidth(CInt(Key) - 1) = vData
End Property

Public Property Get Height() As Integer
    Height = mvarGrid.RowHeight(CInt(Key))
End Property

Public Property Let Height(vData As Integer)
    mvarGrid.RowHeight(CInt(Key)) = vData
End Property

Friend Property Get Key() As String
    Key = mstrKey
End Property

Friend Property Let Key(vData As String)
    mstrKey = vData
End Property

Public Property Get Value() As Variant
    Value = mvarValue
End Property

Public Property Let Value(vData As Variant)
    mvarValue = vData
End Property

Public Property Get Sorted() As SortedTypes
    Sorted = mvarSorted
End Property

Public Property Let Sorted(vData As SortedTypes)
    mvarSorted = vData
    With mvarGrid
        If .Rows <= 1 Then Exit Property        '   Don't bother sorting as there are no rows.
        .col = CInt(mstrKey) - 1
        .ColSel = .col
        .row = .RowSel
        .Sort = mvarSorted
    End With
End Property

Public Property Let Align(vData As MSFlexGridLib.AlignmentSettings)
    mvarGrid.ColAlignment(CInt(Key) - 1) = vData
End Property

Public Property Get Align() As MSFlexGridLib.AlignmentSettings
    Align = mvarGrid.ColAlignment(CInt(Key) - 1)
End Property
