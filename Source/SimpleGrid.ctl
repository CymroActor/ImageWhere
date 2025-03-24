VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl SimpleGrid 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4365
   KeyPreview      =   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4365
   Begin MSComctlLib.ImageList imlSimpleGrid 
      Left            =   2280
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SimpleGrid.ctx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid msgSimpleGrid 
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   4471
      _Version        =   393216
      Rows            =   1
      FixedCols       =   0
      WordWrap        =   -1  'True
      FocusRect       =   0
      SelectionMode   =   1
   End
   Begin VB.Label labTest 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "SimpleGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mcolColumns As GridColumns
Private mintKeyCol As Integer
Private mbolShiftPressed As Boolean
Private mstrToolTipText As String
Private mintCurrentCol As Integer
Private mlngCurrentLeft As Long
Private mintCurrentRow As Integer
Private mintPreSortRow As Integer
Private mintSortCol As Integer
Private mintSortSequence As Integer
Private mblnRowSelected     As Boolean
'
'   ToolTip properties
'
Private Const CONST_TOOLTIP_DELAY As Integer = 4
Private mdteToolTipDelayStart As Date
'
'   Column Properties.
'
Private mcolColumnProperties As Collection
'
'   Events.
'
Public Event Click()
Public Event DblClick()
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event RowChanged(CurrentRow As String)
Public Event MouseHover(CurrentKeyCol As String)
Public Event KeyPress(KeyAscii As Integer)

Private Sub msgSimpleGrid_Click()
    mintPreSortRow = msgSimpleGrid.row
    mblnRowSelected = True
End Sub

Private Sub msgSimpleGrid_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub msgSimpleGrid_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub msgSimpleGrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim intRow As Integer
    Dim intStart As Integer
    
    If Button = vbRightButton Then
        '
        '   Work out which row to select as this isn't automatic.
        '
        With msgSimpleGrid
            intStart = .RowHeight(0)
            For intRow = .TopRow To .Rows - 1
                If Not .RowIsVisible(intRow) Then
                    Exit For
                End If
                If Y > intStart And Y < intStart + .RowHeight(intRow) Then
                    SelectRow (intRow + 1)
                    RowChanged
                    Exit For
                End If
                intStart = intStart + .RowHeight(intRow)
            Next intRow
        End With
        RaiseEvent MouseDown(Button, Shift, X, Y)
    End If
End Sub

Private Sub msgSimpleGrid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReturnColRowMouseOver X, Y

    With msgSimpleGrid
        If .RowSel <> .row Then
            SelectRow (.RowSel + 1)
        End If
    End With
End Sub

Private Sub msgSimpleGrid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    If mintCurrentRow = 0 Then
        SortByColumn mintCurrentCol
    End If
    RaiseEvent Click
End Sub

Private Sub msgSimpleGrid_SelChange()
    RowChanged
End Sub

Private Sub RowChanged()
    With msgSimpleGrid
        If .Redraw = True Then
            If .RowSel <> .row Then
                SelectRow (.RowSel + 1)
            End If
            mcolColumns.Refresh
            If .row <> 0 Then
                RaiseEvent RowChanged(msgSimpleGrid.TextArray(fgIndex(msgSimpleGrid, msgSimpleGrid.row, mintKeyCol)))
            End If
        End If
    End With
    mblnRowSelected = True
End Sub

Private Sub UserControl_Initialize()
    mintSortCol = -1
    mintPreSortRow = 1
    mdteToolTipDelayStart = 0
    mblnRowSelected = True
    labTest.Visible = True
    With msgSimpleGrid
        .row = 0
        .CellAlignment = flexAlignLeftTop
    End With
    Set mcolColumns = New GridColumns
    Set mcolColumns.Grid = msgSimpleGrid
    mcolColumns.Refresh
    Set mcolColumnProperties = New Collection
End Sub

Private Sub UserControl_KeyDown(keycode As Integer, Shift As Integer)
    If Shift <> 0 And Shift <> vbAltMask Then
        keycode = 0
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    msgSimpleGrid.Cols = PropBag.ReadProperty("Columns", 1) + 1
    msgSimpleGrid.ColWidth(msgSimpleGrid.Cols - 1) = 0
    mintKeyCol = PropBag.ReadProperty("KeyCol", 1)
End Sub

Private Sub UserControl_Resize()
    
    On Error GoTo ErrorProc
    
    With msgSimpleGrid
        .Left = 0
        .Top = 0
        .Width = UserControl.Width
        .Height = UserControl.Height
    End With
    Exit Sub
ErrorProc:
    Err.Raise Err.Number
End Sub

Public Sub ResizeRows(Optional intResizeRow As Integer = 0)
    Dim intMaxHeight As Integer
    Dim intRow As Integer
    Dim intCount As Integer
    Dim intCurrentRow As Integer
    Dim intStart As Integer
    Dim intEnd As Integer
    
    On Error GoTo ErrorProc
    
    If intResizeRow = 0 Then
        intStart = 0
        intEnd = msgSimpleGrid.Rows - 1
    Else
        intStart = intResizeRow
        intEnd = intResizeRow
    End If
    
    With msgSimpleGrid
        .Redraw = False
        intCurrentRow = .row
        For intRow = intStart To intEnd
            .row = intRow
            intMaxHeight = 0
            For intCount = 0 To .Cols - 1
                labTest.Height = 200
                labTest.Width = 3000
                labTest.Caption = .TextMatrix(.row, intCount)
                'Debug.Print labTest.Caption
                If .ColWidth(intCount) > 100 Then
                    labTest.Width = .ColWidth(intCount) - 100
                Else
                    labTest.Width = 0
                End If
                labTest.AutoSize = True
                DoEvents
                If intMaxHeight < labTest.Height Then
                    intMaxHeight = labTest.Height
                End If
            Next intCount
            .RowHeight(msgSimpleGrid.row) = intMaxHeight
        Next intRow
        SelectRow (intCurrentRow + 1)
        .Redraw = True
    End With
    Exit Sub
ErrorProc:
    msgSimpleGrid.Redraw = True
    Err.Raise Err.Number
End Sub
Public Property Get Columns() As Integer
    Columns = msgSimpleGrid.Cols - 1 - msgSimpleGrid.FixedCols
End Property

Public Property Let Columns(ByVal vNewValue As Integer)
    Dim intCol As Integer
    Dim oColumnProperty As GridColumnProperty
    '
    '   Add an extra column on the end to hide the selection when there are no
    '   records in the grid e.g. after a clear.
    '
    msgSimpleGrid.Cols = vNewValue + 1 + msgSimpleGrid.FixedCols
    
    msgSimpleGrid.ColWidth(msgSimpleGrid.Cols - 1) = 0
    mcolColumns.Refresh
    '
    '   Initialise the alignment of each column.
    '
    For intCol = 0 To msgSimpleGrid.Cols - 1
        msgSimpleGrid.ColAlignment(intCol) = flexAlignLeftTop
    Next intCol
    '
    '   Select the Invisible Column.
    '
    With msgSimpleGrid
        .row = 0
        .RowSel = 0
        .col = .Cols - 1
        .ColSel = .Cols - 1
    End With
    '
    '   Add a Column Property object for each Column apart from the invisible one.
    '
    For intCol = 1 To vNewValue
        Set oColumnProperty = New GridColumnProperty
        oColumnProperty.PropertyType = smgPropTypeText  '   Default
        mcolColumnProperties.Add oColumnProperty
    Next intCol
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Columns", msgSimpleGrid.Cols - 1
    PropBag.WriteProperty "KeyCol", mintKeyCol, 1
End Sub

Public Sub AddRow(blnRefreshColumns As Boolean, _
                  ParamArray Values() As Variant)
    Dim strValues As String
    Dim intCount As Integer
    Dim intMaxHeight As Integer
    Dim blnTrueFalse As Boolean
    
    If UBound(Values) > msgSimpleGrid.Cols - 1 Then
        Err.Raise vbObjectError + 3, "SimpleGrid", "There are too many values"
        Exit Sub
    End If
    
    strValues = Values(0)
    For intCount = 1 To UBound(Values)
        strValues = strValues & vbTab & Values(intCount)
    Next intCount
    
    With msgSimpleGrid
        .AddItem strValues
        'If blnRefreshColumns Then
            SelectRow .Rows
            .TopRow = IIf(.Rows < 3, 1, .Rows - 2)
        'End If
    End With
    '
    '   Only refresh columns if required.
    '
    'If blnRefreshColumns Then
        mcolColumns.Refresh
    'End If
    '
    '   For each column where there is a graphic representation,
    '   add the graphic.
    '
    With msgSimpleGrid
        For intCount = 1 To .Cols - 1
            If mcolColumnProperties(intCount).PropertyType = smgPropTypeBooleanGraphic Then
                On Error Resume Next
                blnTrueFalse = CBool(.TextMatrix(.row, intCount - 1))
                'blnTrueFalse = CBool(Values(intCount - 1))
                If Err.Number = 13 Then
                    blnTrueFalse = False
                Else
                    On Error GoTo 0
                End If
                
                If blnTrueFalse = True Then
                    .col = intCount - 1
                    Set .CellPicture = imlSimpleGrid.ListImages(1).Picture
                    .CellPictureAlignment = flexAlignCenterTop
                End If
                .TextMatrix(.row, intCount - 1) = ""
            End If
        Next intCount
    End With
    '
    '   Decide whether to increase size of row if required.
    '
    If blnRefreshColumns Then
        ResizeRows msgSimpleGrid.Rows - 1
    End If
End Sub

Public Sub UpdateRow(ByVal intRow As Integer, _
                     ParamArray Values() As Variant)
    Dim strValues As String
    Dim intCount As Integer
    Dim intMaxHeight As Integer
    Dim blnTrueFalse As Boolean
    
    If UBound(Values) > msgSimpleGrid.Cols - 1 Then
        Err.Raise vbObjectError + 3, "SimpleGrid", "There are too many values"
        Exit Sub
    End If
    '
    '   For each column where there is a graphic representation,
    '   add the graphic.
    '
    With msgSimpleGrid
        For intCount = 1 To .Cols - 1
            .TextMatrix(intRow, intCount - 1) = Values(intCount - 1)
            If mcolColumnProperties(intCount).PropertyType = smgPropTypeBooleanGraphic Then
                On Error Resume Next
                blnTrueFalse = CBool(.TextMatrix(.row, intCount - 1))
                If Err.Number = 13 Then
                    blnTrueFalse = False
                Else
                    On Error GoTo 0
                End If

                .col = intCount - 1
                If blnTrueFalse = True Then
                    Set .CellPicture = imlSimpleGrid.ListImages(1).Picture
                    .CellPictureAlignment = flexAlignCenterTop
                Else
                    Set .CellPicture = Nothing
                End If
                .TextMatrix(.row, intCount - 1) = ""
            End If
        Next intCount
    End With
    '
    '   Refresh the internal collection.
    '
    ''mcolColumns.Refresh
    '
    '   Select the current row.
    '
    SelectRow intRow + 1
End Sub

Public Property Get Column(intCol As Integer) As GridColumn
    If intCol > msgSimpleGrid.Cols Then
        Err.Raise vbObjectError + 1, "SimpleGrid: Get Column", "Invalid Column Number"
        Exit Property
    End If
    
    Set Column = mcolColumns.Item(CStr(intCol))
End Property

Public Property Get ColumnProperty(intCol As Integer) As GridColumnProperty
    Set ColumnProperty = mcolColumnProperties(intCol)
End Property

Public Property Set ColumnProperty(intCol As Integer, vData As GridColumnProperty)
    Set mcolColumnProperties(intCol) = vData
End Property

Public Property Get RowHeight(intRow As Integer) As Integer
    If intRow > msgSimpleGrid.Rows Then
        Err.Raise vbObjectError + 2, "SimpleGrid: Get RowHeight", "Invalid Row Number"
        Exit Property
    End If
    
    RowHeight = msgSimpleGrid.RowHeight(intRow)
End Property

Public Property Get KeyCol() As Integer
    KeyCol = mintKeyCol + 1
End Property

Public Property Let KeyCol(vData As Integer)
    If vData = 0 Then
        MsgBox "KeyCol must be greater than zero.", vbCritical
        Exit Property
    End If
    mintKeyCol = vData - 1
End Property

Public Function GetKeyRow(strKey As String) As Boolean
    Dim intCount As Integer
    
    msgSimpleGrid.Redraw = False
    GetKeyRow = False
    With msgSimpleGrid
        For intCount = 2 To .Rows
            If LCase(.TextArray(fgIndex(msgSimpleGrid, intCount - 1, mintKeyCol))) = LCase(strKey) Then
                GetKeyRow = True
                SelectRow (intCount)
                Exit For
            End If
        Next intCount
    End With

    If Not GetKeyRow Then
        SelectRow (2)
    End If
    If msgSimpleGrid.row <> 0 Then
        msgSimpleGrid.TopRow = IIf(msgSimpleGrid.row < 3, 1, msgSimpleGrid.row - 2)
    End If
    msgSimpleGrid.Redraw = True
    mcolColumns.Refresh

End Function

Public Sub Deselect()
    With msgSimpleGrid
        .col = .Cols - 1
        .ColSel = .Cols - 1
        mblnRowSelected = False
    End With
End Sub

Private Sub SelectRow(intRow As Integer)
        
    If msgSimpleGrid.Rows = 1 Then Exit Sub
    
    With msgSimpleGrid
        .col = 0
        .row = intRow - 1
        .ColSel = .Cols - 1
        .RowSel = intRow - 1
    End With
End Sub

Public Sub DeleteRow()
    msgSimpleGrid.RemoveItem msgSimpleGrid.row
End Sub

Public Sub Clear()
    With msgSimpleGrid
        .Rows = 1
        .row = 0
        .col = .Cols - 1
        .ColSel = .Cols - 1
    End With
    
End Sub

Public Property Get Rows() As Integer
    Rows = msgSimpleGrid.Rows - 1
End Property

Public Property Get CurrentRow() As Integer
    If mblnRowSelected Then
        CurrentRow = msgSimpleGrid.row
    Else
        CurrentRow = 0
    End If
End Property

Public Property Let CurrentRow(vData As Integer)
    If msgSimpleGrid.Rows <> 1 Then
        msgSimpleGrid.row = vData
        SelectRow (vData + 1)
        mcolColumns.Refresh
    End If
End Property

Public Property Let Redraw(vData As Boolean)
    msgSimpleGrid.Redraw = vData
    If vData = True Then
        mcolColumns.Refresh
    End If
End Property

Public Property Let TopRow(vData As Integer)
    msgSimpleGrid.TopRow = vData
End Property

Public Property Get TopRow() As Integer
    TopRow = msgSimpleGrid.TopRow
End Property

Public Property Get HeaderHeight() As Integer
    HeaderHeight = msgSimpleGrid.RowHeight(0)
End Property

Public Property Let ToolTip(vData As String)
    mstrToolTipText = vData
    msgSimpleGrid.ToolTipText = vData
End Property

Public Property Get ToolTip() As String
    ToolTip = msgSimpleGrid.ToolTipText
End Property

Public Property Get KeyPress()
    
End Property

Private Sub ReturnColRowMouseOver(X As Single, Y As Single)
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim intCol As Integer
    Dim intRow As Integer
    Dim blnRow As Boolean
    Dim blnCol As Boolean
    Dim intCurrentRow As Integer
    Dim strKey As String
    
    mintCurrentCol = -1
    mintCurrentRow = -1
    lngLeft = 0
    lngTop = 0
    blnCol = False
    blnRow = False
        
    For intCol = 0 To msgSimpleGrid.Cols - 1
        If X > lngLeft And X <= lngLeft + msgSimpleGrid.ColWidth(intCol) Then
            blnCol = True
            mintCurrentCol = intCol
            Exit For
        End If
        lngLeft = lngLeft + msgSimpleGrid.ColWidth(intCol)
    Next intCol
            
    'lngTop = msgSimpleGrid.RowHeight(0)
    'For intRow = msgSimpleGrid.TopRow To msgSimpleGrid.Rows - 1
    For intRow = 0 To msgSimpleGrid.Rows - 1
        If Y > lngTop And Y <= (lngTop + msgSimpleGrid.RowHeight(intRow)) Then
            blnRow = True
            'mintCurrentRow = intRow
            Exit For
        End If
        lngTop = lngTop + msgSimpleGrid.RowHeight(intRow)
    Next intRow
    
    If blnCol And blnRow Then
        mintCurrentCol = intCol
        If mintCurrentRow <> intRow Then
            mintCurrentRow = intRow
            If mintCurrentCol <> 0 Then
                strKey = msgSimpleGrid.TextMatrix(mintCurrentRow, mintKeyCol)
                If mintCurrentRow <= 0 Then
                    RaiseEvent MouseHover(0)
                ElseIf strKey <> "0" And strKey <> "" Then
                    RaiseEvent MouseHover(strKey)
                End If
            End If
        End If
    Else
        RaiseEvent MouseHover(0)
    End If
End Sub

Private Sub SortByColumn(intCol As Integer)
    Dim strKey As String
    Dim intCount As Integer
    Dim blnDateConverted As Boolean
    
    With msgSimpleGrid
    
        If .Rows <= 1 Then Exit Sub
        
        strKey = .TextMatrix(mintPreSortRow, mintKeyCol)
        .col = intCol
        If IsDate(.TextMatrix(1, intCol)) And Val(.TextMatrix(1, intCol)) <> 0 Then
            blnDateConverted = True
            For intCount = 1 To .Rows - 1
''                .TextMatrix(intCount, intCol) = CDbl(CDate(.TextMatrix(intCount, intCol)))
                If IsDate(.TextMatrix(intCount, intCol)) Then
                    .TextMatrix(intCount, intCol) = CDbl(CDate(.TextMatrix(intCount, intCol)))
                Else
                    .TextMatrix(intCount, intCol) = 0
                End If
            Next intCount
        End If
        If mintSortCol = intCol Then
            If mintSortSequence = flexSortGenericAscending Then
                .Sort = flexSortGenericDescending
                mintSortSequence = flexSortGenericDescending
            Else
                .Sort = flexSortGenericAscending
                mintSortSequence = flexSortGenericAscending
            End If
        Else
            .Sort = flexSortGenericAscending
            mintSortSequence = flexSortGenericAscending
        End If
        If blnDateConverted Then
            For intCount = 1 To .Rows - 1
                If .TextMatrix(intCount, intCol) = 0 Then
                    .TextMatrix(intCount, intCol) = ""
                Else
                    Select Case mcolColumnProperties(intCol + 1).PropertyType
                        Case Is = smgPropTypeDateShort
                            .TextMatrix(intCount, intCol) = Format(.TextMatrix(intCount, intCol), "DD/MM/YYYY")
                        Case Is = smgPropTypeDateLong
                            .TextMatrix(intCount, intCol) = Format(.TextMatrix(intCount, intCol), "DD/MM/YYYY hh:MM:ss")
                        Case Else
                            .TextMatrix(intCount, intCol) = Format(.TextMatrix(intCount, intCol), "DD/MM/YYYY")
                    End Select
                End If
            Next intCount
        End If
        mintSortCol = intCol
        GetKeyRow strKey
    End With
End Sub

Public Property Let RowBold(ByVal lngRow As Long, ByVal vData As Boolean)
    Dim intCurrentRow As Integer
    
    With msgSimpleGrid
        .RowData(lngRow) = vData
        '
        '   Retain the current row.
        '
        intCurrentRow = CurrentRow
        '
        '   Grey out, or otherwise, the selected row.
        '
        'HighlightRow CInt(lngRow + 1)
        .FillStyle = flexFillRepeat
        .CellFontBold = vData
        SelectRow intCurrentRow + 1
    End With
End Property


