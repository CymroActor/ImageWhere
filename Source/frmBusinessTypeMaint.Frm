VERSION 5.00
Begin VB.Form frmBusinessTypeMaint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Business Type"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmBusinessTypeMaint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraReturnPeriods 
      Caption         =   "Return Periods"
      Height          =   2655
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   4455
      Begin VB.TextBox txtLossFee 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         TabIndex        =   14
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txtPhone2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         TabIndex        =   12
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtPhone1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         TabIndex        =   10
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txtSL2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         TabIndex        =   8
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtSL1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         TabIndex        =   6
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtInitial 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblLossFee 
         Caption         =   "Loss Fee:"
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label lblPhone2 
         Caption         =   "Phone2:"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label lblPhone1 
         Caption         =   "Phone1:"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label lblSl2 
         Caption         =   "SL2:"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lblSL1 
         Caption         =   "SL1:"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblInitial 
         Caption         =   "&Initial:"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   3375
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   555
      Left            =   3720
      Picture         =   "frmBusinessTypeMaint.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3720
      Width           =   795
   End
   Begin VB.CommandButton cmbOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   555
      Left            =   2400
      Picture         =   "frmBusinessTypeMaint.frx":0316
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3720
      Width           =   795
   End
   Begin VB.Label labBusinessType 
      Caption         =   "&Name:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "frmBusinessTypeMaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event BusinessTypeUpdated(Name As String)
Public Event BusinessTypeAdded(Name As String)
Public mMode As UpdateMode
Private moBusinessType As BusinessType

Private Sub cmbOK_Click()
'***************************************
' Module/Form Name   : frmBusinessTypeMaint
'
' Procedure Name     : cmbOK_Click
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
On Error GoTo cmbOK_Click_Error
'
'******** Code Starts Here *************
'

    If Not ValidInput Then
        Exit Sub
    End If

    If mMode = Edit Then
        With moBusinessType
            .InitialReturnPeriod = txtInitial
            .SL1ReturnPeriod = txtSL1
            .SL2ReturnPeriod = txtSL2
            .Phone1ReturnPeriod = txtPhone1
            .Phone2ReturnPeriod = txtPhone2
            .LossFeeReturnPeriod = txtLossFee
            On Error Resume Next
            .update
            If Err.Number = vbObjectError + 1 Then
                DisplayError
                DisplayBusinessType
                Exit Sub
            Else
                If Err.Number <> 0 Then
                    ErrorSave
                    On Error GoTo cmbOK_Click_Error
                    ErrorRestore
                Else
                    On Error GoTo cmbOK_Click_Error
                End If
            End If
            RaiseEvent BusinessTypeUpdated(.Name)
        End With
    Else
        moBusinessType.create txtName, _
                              txtInitial, _
                              txtSL1, _
                              txtSL2, _
                              txtPhone1, _
                              txtPhone2, _
                              txtLossFee

        RaiseEvent BusinessTypeAdded(moBusinessType.Name)
    End If

    Unload Me
'
'********* Code Ends Here **************
'
   Exit Sub
'
cmbOK_Click_Error:
    DisplayError , "frmBusinessTypeMaint.cmbOK_Click", vbExclamation
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub txtInitial_GotFocus()
    HighLightText txtInitial
End Sub

Private Sub txtLossFee_GotFocus()
    HighLightText txtLossFee
End Sub

Private Sub txtName_Change()
    If Len(Trim(txtName.Text)) = 0 Then
        txtInitial.Enabled = False
        txtInitial.BackColor = vbButtonFace
    Else
        txtInitial.Enabled = True
        txtInitial.BackColor = vbWindowBackground
    End If
End Sub

Private Function ValidInput() As Boolean
'***************************************
' Module/Form Name   : frmBusinessTypeMaint
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

    If Len(Trim(txtName)) = 0 Then
        MsgBox "Please enter the Business Type Name.", vbExclamation
        HighLightText txtName
        txtName.SetFocus
        Exit Function
    End If

    If Len(Trim(txtInitial)) = 0 Then
        MsgBox "Please enter the Initial Return Period.", vbExclamation
        txtInitial.SetFocus
        Exit Function
    End If

    If CDbl(txtInitial) > 1500 Then
        MsgBox "Period must not be greater than 1500 days.", vbExclamation
        HighLightText txtInitial
        txtInitial.SetFocus
        Exit Function
    End If

    If CDbl(txtInitial) < 1 Then
        MsgBox "Period must not be less than 1 day.", vbExclamation
        HighLightText txtInitial
        txtInitial.SetFocus
        Exit Function
    End If

    If Len(Trim(txtSL1)) = 0 Then
        MsgBox "Please enter the SL1 Return Period.", vbExclamation
        txtSL1.SetFocus
        Exit Function
    End If

    If CDbl(txtSL1) > 1500 Then
        MsgBox "Period must not be greater than 1500 days.", vbExclamation
        HighLightText txtSL1
        txtSL1.SetFocus
        Exit Function
    End If

    If CDbl(txtSL1) < 1 Then
        MsgBox "Period must not be less than 1 day.", vbExclamation
        HighLightText txtSL1
        txtSL1.SetFocus
        Exit Function
    End If

    If Len(Trim(txtSL2)) = 0 Then
        MsgBox "Please enter the SL2 Return Period.", vbExclamation
        txtSL2.SetFocus
        Exit Function
    End If

    If CDbl(txtSL2) > 1500 Then
        MsgBox "Period must not be greater than 1500 days.", vbExclamation
        HighLightText txtSL2
        txtSL2.SetFocus
        Exit Function
    End If

    If CDbl(txtSL2) < 1 Then
        MsgBox "Period must not be less than 1 day.", vbExclamation
        HighLightText txtSL2
        txtSL2.SetFocus
        Exit Function
    End If

    If Len(Trim(txtPhone1)) = 0 Then
        MsgBox "Please enter the Phone1 Return Period.", vbExclamation
        txtPhone1.SetFocus
        Exit Function
    End If

    If CDbl(txtPhone1) > 1500 Then
        MsgBox "Period must not be greater than 1500 days.", vbExclamation
        HighLightText txtPhone1
        txtPhone1.SetFocus
        Exit Function
    End If

    If CDbl(txtPhone1) < 1 Then
        MsgBox "Period must not be less than 1 day.", vbExclamation
        HighLightText txtPhone1
        txtPhone1.SetFocus
        Exit Function
    End If

    If Len(Trim(txtPhone2)) = 0 Then
        MsgBox "Please enter the Phone2 Return Period.", vbExclamation
        txtPhone2.SetFocus
        Exit Function
    End If

    If CDbl(txtPhone2) > 1500 Then
        MsgBox "Period must not be greater than 1500 days.", vbExclamation
        HighLightText txtPhone2
        txtPhone2.SetFocus
        Exit Function
    End If

    If CDbl(txtPhone2) < 1 Then
        MsgBox "Period must not be less than 1 day.", vbExclamation
        HighLightText txtPhone2
        txtPhone2.SetFocus
        Exit Function
    End If

    If Len(Trim(txtLossFee)) = 0 Then
        MsgBox "Please enter the LossFee Return Period.", vbExclamation
        txtLossFee.SetFocus
        Exit Function
    End If

    If CDbl(txtLossFee) > 1500 Then
        MsgBox "Period must not be greater than 1500 days.", vbExclamation
        HighLightText txtLossFee
        txtLossFee.SetFocus
        Exit Function
    End If

    If CDbl(txtLossFee) < 1 Then
        MsgBox "Period must not be less than 1 day.", vbExclamation
        HighLightText txtLossFee
        txtLossFee.SetFocus
        Exit Function
    End If

    ValidInput = True

'
'********* Code Ends Here **************
'
   Exit Function
'
ValidInput_Error:
    ErrorRaise "frmBusinessTypeMaint.ValidInput"
End Function

Private Sub txtInitial_KeyPress(KeyAscii As Integer)
    KeyAscii = allow_numeric_only(KeyAscii)
End Sub

Private Sub txtLossFee_KeyPress(KeyAscii As Integer)
    KeyAscii = allow_numeric_only(KeyAscii)
End Sub

Private Sub txtName_GotFocus()
    HighLightText txtName
End Sub

Private Sub txtPhone1_GotFocus()
    HighLightText txtPhone1
End Sub

Private Sub txtPhone1_KeyPress(KeyAscii As Integer)
    KeyAscii = allow_numeric_only(KeyAscii)
End Sub

Private Sub txtPhone2_GotFocus()
    HighLightText txtPhone2
End Sub

Private Sub txtPhone2_KeyPress(KeyAscii As Integer)
    KeyAscii = allow_numeric_only(KeyAscii)
End Sub

Private Sub txtSL1_GotFocus()
    HighLightText txtSL1
End Sub

Private Sub txtSL1_KeyPress(KeyAscii As Integer)
    KeyAscii = allow_numeric_only(KeyAscii)
End Sub

Private Sub txtSL2_GotFocus()
    HighLightText txtSL2
End Sub

Private Sub txtSL2_KeyPress(KeyAscii As Integer)
    KeyAscii = allow_numeric_only(KeyAscii)
End Sub

Public Sub Display(Mode As UpdateMode, Optional oBusinessType As BusinessType)
'***************************************
' Module/Form Name   : frmBusinessTypeMaint
'
' Procedure Name     : Display
'
' Purpose            :
'
' Date Created       : 28/12/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : Mode - UpdateMode
'                    : oBusinessType - BusinessType
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo Display_Error
'
'******** Code Starts Here *************
'
    mMode = Mode
    If mMode = Edit Then
        Set moBusinessType = oBusinessType
        Me.Caption = "Edit " & Me.Caption & " '" & oBusinessType.Name & "'"
        With txtName
            .Locked = True
            .BackColor = vbButtonFace
        End With
        DisplayBusinessType
        SendKeys "{TAB}"
    Else
        Set moBusinessType = New BusinessType
        Me.Caption = "Add " & Me.Caption
        With txtName
            .Locked = False
            .BackColor = vbWindowBackground
        End With
    End If
    Me.Show 1
'
'********* Code Ends Here **************
'
   Exit Sub
'
Display_Error:
    ErrorRaise "frmBusinessTypeMaint.Display"
End Sub

Private Sub DisplayBusinessType()
    With moBusinessType
        txtName = .Name
        txtInitial = .InitialReturnPeriod
        txtSL1 = .SL1ReturnPeriod
        txtSL2 = .SL2ReturnPeriod
        txtPhone1 = .Phone1ReturnPeriod
        txtPhone2 = .Phone2ReturnPeriod
        txtLossFee = .LossFeeReturnPeriod
    End With
End Sub
