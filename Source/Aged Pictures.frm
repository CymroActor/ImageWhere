VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAgedPictures 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aged Pictures Report Parameters"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4140
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   4140
   StartUpPosition =   1  'CenterOwner
   Begin MSComCtl2.DTPicker dtpDateFrom 
      Height          =   315
      Left            =   2580
      TabIndex        =   2
      Top             =   780
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   67174403
      CurrentDate     =   36809
   End
   Begin VB.ComboBox cboAction 
      Height          =   315
      Left            =   1500
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   180
      Width           =   2415
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   555
      Left            =   3120
      Picture         =   "Aged Pictures.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2100
      Width           =   795
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   555
      Left            =   2100
      Picture         =   "Aged Pictures.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2100
      Width           =   795
   End
   Begin MSComCtl2.DTPicker dtpDateTo 
      Height          =   315
      Left            =   2580
      TabIndex        =   4
      Top             =   1380
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   67174403
      CurrentDate     =   36809
   End
   Begin VB.Label lblDateTo 
      Caption         =   "Date &To:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   1200
   End
   Begin VB.Label lblDateFrom 
      Caption         =   "Date &From:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1200
   End
   Begin VB.Label lblNextAction 
      Caption         =   "Next &Action:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1200
   End
End
Attribute VB_Name = "frmAgedPictures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
'***************************************
' Module/Form Name   : frmAgedPictures
'
' Procedure Name     : cmdOK_Click
'
' Purpose            :
'
' Date Created       : 28/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo cmdOK_Click_Error
'
'******** Code Starts Here *************
'
''    Dim oAccessDB As access.Application
    Dim oAccessDB As Object
    
    Screen.MousePointer = vbHourglass
    If Not ValidInput Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    CreateAgedDebtorsQuery cboAction, dtpDateFrom.Value, dtpDateTo.Value

    Set oAccessDB = GetObject(glo_dbname, "Access.Application")
    oAccessDB.Visible = True
    oAccessDB.DoCmd.Maximize
    oAccessDB.DoCmd.RunCommand 10                           'acCmdAppMaximize
    oAccessDB.DoCmd.OpenReport "REP_AGED_PICTURES", 2       'acPreview
    Set oAccessDB = Nothing
    Screen.MousePointer = vbDefault
'
'********* Code Ends Here **************
'
   Exit Sub
'
cmdOK_Click_Error:
    DisplayError , "frmAgedPictures.cmdOK_Click", vbExclamation
End Sub

Private Sub Form_Load()
'***************************************
' Module/Form Name   : frmAgedPictures
'
' Procedure Name     : Form_Load
'
' Purpose            :
'
' Date Created       : 28/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo Form_Load_Error
'
'******** Code Starts Here *************
'
    With cboAction
        .AddItem "<All>"
        .AddItem "SL1"
        .AddItem "SL2"
        .AddItem "Phone 1"
        .AddItem "Phone 2"
        .AddItem "Loss/Fee"
        .AddItem "Miscellaneous"
        LocateComboItem cboAction, GetSetting(App.Title, "AgedDebtors", "Action", "")
        If cboAction.ListIndex = -1 Then
            cboAction.ListIndex = 0
        End If
    End With

    dtpDateFrom.Value = GetSetting(App.Title, "AgedDebtors", "DateFrom", Now)
    dtpDateTo.Value = GetSetting(App.Title, "AgedDebtors", "DateTo", Now)
'
'********* Code Ends Here **************
'
   Exit Sub
'
Form_Load_Error:
    DisplayError , "frmAgedPictures.Form_Load", vbExclamation
End Sub

Private Sub Form_Unload(Cancel As Integer)
'***************************************
' Module/Form Name   : frmAgedPictures
'
' Procedure Name     : Form_Unload
'
' Purpose            :
'
' Date Created       : 28/12/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : Cancel - Integer
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo Form_Unload_Error
'
'******** Code Starts Here *************
'
    SaveSetting App.Title, "AgedDebtors", "Action", cboAction.Text
    SaveSetting App.Title, "AgedDebtors", "DateFrom", dtpDateFrom.Value
    SaveSetting App.Title, "AgedDebtors", "DateTo", dtpDateTo.Value
'
'********* Code Ends Here **************
'
   Exit Sub
'
Form_Unload_Error:
    DisplayError , "frmAgedPictures.Form_Unload", vbExclamation
End Sub

Private Sub CreateAgedDebtorsQuery(strAction As String, _
                                   dteDateFrom As Date, _
                                   dteDateTo As Date)
'***************************************
' Module/Form Name   : frmAgedPictures
'
' Procedure Name     : CreateAgedDebtorsQuery
'
' Purpose            :
'
' Date Created       : 28/12/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : strAction - String
'                    : dteDateFrom - Date
'                    : dteDateTo - Date
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo CreateAgedDebtorsQuery_Error
'
'******** Code Starts Here *************
'
    Dim qdfAgedDebtors As QueryDef
    Dim strSQL As String
    Dim strWhereClause As String
    Dim rstRecordRequest As DAO.Recordset
    '
    '   Set the report parameters.
    '
    strSQL = "SELECT * FROM REPORTREQUEST WHERE NAME = 'AGEDPICTURES'"
    Set rstRecordRequest = db.OpenRecordset(strSQL, dbOpenDynaset, dbPessimistic)
    With rstRecordRequest
        .Edit
        !Parameter1 = Format(dteDateFrom, "DD/MM/YYYY 00:00:00")
        !Parameter2 = Format(dteDateTo, "DD/MM/YYYY 23:59:59")
        !Parameter3 = strAction
        .update
    End With
    '
    '
    ' Delete the current query.
    '
    On Error Resume Next

    db.QueryDefs.Delete "QRY_AGE_DEBTORS"
    '
    ' Create the QueryDef.
    '
    strSQL = "SELECT (SELECT MAX(ID) FROM CHASER WHERE chaser.deliverynoteno = Delivery_note.delivery_note_no) AS LatestId, " _
           & "Delivery_note.delivery_note_no, Chaser.Id, Chaser.NextAction, Customer.customer_name, Chaser.ReturnByDate, Chaser.Contact, Chaser.Comment, Chaser.Action " _
           & "FROM ReportRequest, Customer INNER JOIN (Delivery_note INNER JOIN Chaser ON Delivery_note.delivery_note_no = Chaser.DeliveryNoteNo) ON Customer.customer_no = Delivery_note.customer_no " _
           & "WHERE (((Chaser.Id)=(SELECT MAX(ID) FROM CHASER WHERE chaser.deliverynoteno = Delivery_note.delivery_note_no)) AND ((Chaser.ReturnByDate)>=CDate([ReportRequest].[PARAMETER1]) And (Chaser.ReturnByDate)<=CDate([ReportRequest].[PARAMETER2]) And (Chaser.ReturnByDate) <> CDate('01-JAN-1980')) AND ((ReportRequest.Name) = 'AGEDPICTURES'))"

    If strAction <> "<All>" Then
        strWhereClause = " AND Chaser.NextAction = '" & strAction & "'"
    End If
    strSQL = strSQL & strWhereClause & " ORDER BY CUSTOMER.CUSTOMER_NAME"
    '
    On Error Resume Next
    Set qdfAgedDebtors = db.CreateQueryDef("QRY_AGE_DEBTORS", strSQL)

    If Err.Number <> 3012 Then
        If Err.Number <> 0 Then
            ErrorSave
            On Error GoTo CreateAgedDebtorsQuery_Error
            ErrorRestore
        End If
    End If

    On Error GoTo CreateAgedDebtorsQuery_Error
    Set rstRecordRequest = Nothing
'
'********* Code Ends Here **************
'
   Exit Sub
'
CreateAgedDebtorsQuery_Error:
    ErrorRaise "frmAgedPictures.CreateAgedDebtorsQuery"
End Sub

Private Function ValidInput() As Boolean
'***************************************
' Module/Form Name   : frmAgedPictures
'
' Procedure Name     : ValidInput
'
' Purpose            :
'
' Date Created       : 28/12/2002
'
' Author             : GARETH SAUNDERS
'
' Returns            : Boolean
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo ValidInput_Error
'
'******** Code Starts Here *************
'
    ValidInput = False
    If dtpDateFrom.Value > dtpDateTo.Value Then
        MsgBox "'Date To' must not be less than 'Date From'", vbExclamation
        dtpDateTo.SetFocus
        Exit Function
    End If

    ValidInput = True
'
'********* Code Ends Here **************
'
   Exit Function
'
ValidInput_Error:
    ErrorRaise "frmAgedPictures.ValidInput"
End Function
