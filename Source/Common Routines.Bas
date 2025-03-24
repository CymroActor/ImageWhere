Attribute VB_Name = "common_routines"
Option Explicit
Option Compare Text
Global transaction_started As Boolean
'
'   Declare APIs
'
Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
Public Const HELP_COMMAND = &H102&
Public Const HELP_CONTENTS = &H3&
Public Const HELP_CONTEXT = &H1
Public Const HELP_CONTEXTPOPUP = &H8&
Public Const HELP_FORCEFILE = &H9&
Public Const HELP_HELPONHELP = &H4
Public Const HELP_INDEX = &H3
Public Const HELP_KEY = &H101
Public Const HELP_MULTIKEY = &H201&
Public Const HELP_PARTIALKEY = &H105&
Public Const HELP_QUIT = &H2
Public Const HELP_SETCONTENTS = &H5&
Public Const HELP_SETINDEX = &H5
Public Const HELP_SETWINPOS = &H203&
Public Const HELPMSGSTRING = "commdlg_help"

Declare Function GetActiveWindow Lib "user32" () As Long

Public Const SW_MAXIMIZE = 3
Declare Function ShowWindow Lib "user32" _
                           (ByVal hWnd As Long, _
                            ByVal nCmdShow As Long) As Long
Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hWnd&, ByVal HelpFile$, ByVal wCommand%, ByVal dwData As Any)
Declare Function apiGetDC Lib "user32" Alias "GetDC" _
   (ByVal hWnd As Long) As Long
Declare Function apiReleaseDC Lib "user32" Alias "ReleaseDC" (ByVal _
   hWnd As Long, ByVal hDC As Long) As Long
Declare Function apiGetDeviceCaps Lib "gdi32" Alias "GetDeviceCaps" _
   (ByVal hDC As Long, ByVal nIndex As Long) As Long
Global Const TWIPSPERINCH = 1440
'
'   Used for setting a form to stay on top.
'
Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Declare Function SetWindowPos Lib "user32" _
    (ByVal hWnd As Long, _
     ByVal hWndInsertAfter As Long, _
     ByVal X As Long, _
     ByVal Y As Long, _
     ByVal cx As Long, _
     ByVal cy As Long, _
     ByVal wFlags As Long) As Long
'
'   System metrics.
'
'Public Declare Function GetSystemMetrics% Lib "user32" (ByVal nIndex%)
'Global Const SM_CYCAPTION = 4
'
Public Enum UpdateMode
    Edit
    Add
    view
End Enum
                                
Public Enum NVLType
        NVLString = 1
        NVLNumeric = 2
End Enum
'
'   Used for generating errors after in-line error handling.
'
Private mlngErrorNumber As Long
Private mstrSource As String
Private mstrErrorDescription As String
'
'   Used for controlling the visibility of controls.
'
Private colInvisibleControls As Collection
'
'   Used for setting Tabstops in a ListBox.
'
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                (ByVal hWnd As Long, _
                 ByVal wMsg As Long, _
                 ByVal wParam As Long, _
                 lParam As Any) As Long

Public Const LB_SETTABSTOPS = &H192
'
'   Used to stop window flickering.
'
Public Declare Function LockWindowUpdate Lib "user32" _
    (ByVal hWnd As Long) As Long

Public Sub HighLightText(txtHighLight As TextBox)
    With txtHighLight
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
Public Function FlexGridIndex(msFlex As MSFlexGrid, row As Integer, col As Integer) As Long
     FlexGridIndex = row * msFlex.Cols + col
End Function
Public Function FlexGridHighlightRow(msFlex As MSFlexGrid, HighLightRow As Integer)
'
'   If the row to be highlighted is 0, there can be no row to highlight!
'
    If HighLightRow = 0 Then
        Exit Function
    End If
    
    With msFlex
        .col = 0
        .row = HighLightRow
        .ColSel = .Cols - 1
        .RowSel = HighLightRow
        .TopRow = HighLightRow
    End With
End Function

Public Sub close_recordset(rset_close)
'
'   This has been written to close a recordset and cope in the
'   case that it is already closed!
'
    On Error GoTo error_proc
    
    rset_close.Close

    Exit Sub
    
error_proc:
    If Err.Number = 91 Then
        Resume Next
    Else
        MsgBox Str(Err.Number) + " - " + Err.Description, vbCritical
    End If

End Sub

Public Function format_search_result(batch_no As Long, photograph_no As Long, batch_description As String) As String
'***************************************
' Module/Form Name   : common_routines
'
' Procedure Name     : format_search_result
'
' Purpose            :
'
' Date Created       : 07/07/2004
'
' Author             : ADMINISTRATOR
'
' Parameters         : batch_no - Long
'                    : photograph_no - Long
'                    : batch_description - String
'
' Returns            : String
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo format_search_result_Error
'
'******** Code Starts Here *************
'
    Dim description_string
    Dim strPhoto      As String
    
    If batch_description = "" Then
        description_string = ""
    Else
        description_string = "    " + batch_description
    End If
    '
    '   Don't output the photograph number if digital.
    '
    If photograph_no = 0 Then
        strPhoto = ""
    Else
        strPhoto = "." + LTrim(Str(photograph_no))
    End If

    format_search_result = Space(6 - Len(Trim(Str(batch_no)))) & Format(Str(batch_no), "@@@@@@") _
                           & strPhoto _
                           & Space(2 - Len(LTrim(Str(photograph_no)))) _
                           & description_string
'
'********* Code Ends Here **************
'
   Exit Function
'
format_search_result_Error:
    ErrorRaise "common_routines.format_search_result"
End Function

Public Function characters_after_space(number_string As String) As String
    characters_after_space = Mid(number_string, InStr(number_string, " ") + 1)
End Function

Public Function characters_after_decimal(number_string) As String
    If InStr(number_string, ".") = 0 Then
        characters_after_decimal = "0"
    Else
        characters_after_decimal = Mid(number_string, InStr(number_string, ".") + 1)
    End If
End Function

Public Function last_character(number_string As String) As String
    If Len(number_string) <> 0 Then
        last_character = Mid(number_string, Len(number_string), 1)
    Else
        last_character = ""
    End If
End Function

Public Function characters_before_space(number_string As String) As String
    If InStr(number_string, " ") = 0 Then
        characters_before_space = number_string
    Else
        characters_before_space = Mid(number_string, 1, InStr(number_string, " ") - 1)
    End If
End Function

Public Function allow_numeric_only(KeyAscii) As Integer
    If (KeyAscii < 48 Or KeyAscii > 57) And _
        Not KeyAscii = 8 Then
        allow_numeric_only = 0
        Beep
    Else
        allow_numeric_only = KeyAscii
    End If
End Function

Public Function AllowCurrencyOnly(KeyAscii As Integer, _
                                  strCurrency As String) As Integer
    '
    '   Allow a back space immediately.
    '
    If KeyAscii = vbKeyBack Then
        AllowCurrencyOnly = KeyAscii
        Exit Function
    End If
    '
    '   Don't allow more than one decimal point.
    '
    If InStr(strCurrency, ".") <> 0 And _
       KeyAscii = 46 Then
        AllowCurrencyOnly = 0
        Beep
        Exit Function
    End If
    '
    If (KeyAscii < 48 Or KeyAscii > 57) And _
        Not KeyAscii = 46 Then
        AllowCurrencyOnly = 0
        Beep
        Exit Function
    End If
    
    AllowCurrencyOnly = KeyAscii
End Function

Public Sub begin_trans()
    
    On Error Resume Next
    If gintTransCount = 0 Then
        ws.BeginTrans
    ''    DBEngine.Idle dbRefreshCache
        gdbADO.BeginTrans
        transaction_started = True
    End If
    gintTransCount = gintTransCount + 1
End Sub


Public Function characters_before_decimal(number_string) As String
    If InStr(number_string, ".") = 0 Then
        characters_before_decimal = number_string
    Else
        characters_before_decimal = Mid(number_string, 1, InStr(number_string, ".") - 1)
    End If
End Function

Public Sub com_position_form(position_form As Form)
    On Error GoTo ErrorProc
    
    If position_form.WindowState <> 0 Then
        Exit Sub
    End If
'
'   1100 is to allow for the '95 toolbar and the title and menu bars.
'
    If position_form.Top + position_form.Height + 1100 > _
        mdi_npls.Height - mdi_npls.tob_npls.Height - mdi_npls.stb_npls.Height Then
        position_form.Top = mdi_npls.Height - position_form.Height - _
                            mdi_npls.tob_npls.Height - 1100 - _
                            mdi_npls.stb_npls.Height
    End If
    If position_form.Left + position_form.Width > Screen.Width Then
        position_form.Left = Screen.Width - position_form.Width
    End If
    Exit Sub
ErrorProc:
    ErrorRaise "common_routines: com_position_form"
End Sub

Public Sub commit_trans()
    On Error Resume Next
    If gintTransCount = 1 Then
        ws.CommitTrans dbForceOSFlush
        gdbADO.CommitTrans
        transaction_started = False
    End If
    If gintTransCount > 0 Then
        gintTransCount = gintTransCount - 1
    End If
End Sub

Public Function is_form_loaded(form_name As String, _
                               Optional CallingForm As Variant, _
                               Optional FormCollection As Variant) As Boolean
    Dim loop_index As Integer
    
    is_form_loaded = False
    
    For loop_index = 0 To Forms.Count - 1
        If Forms(loop_index).Name = form_name And Forms(loop_index).Visible = True Then
            If Not IsMissing(CallingForm) Then
                If Not Forms(loop_index) Is CallingForm Then
                    is_form_loaded = True
                    If Not IsMissing(FormCollection) Then
                        FormCollection.Add Forms(loop_index)
                    End If
                End If
            Else
                is_form_loaded = True
                Exit For
            End If
        End If
    Next loop_index

End Function

Public Sub maximize_window()
    Dim handle As Integer
    
    MsgBox "Getting Active Window"
    handle = GetActiveWindow
    
    MsgBox "Maximizing Window"
    ShowWindow handle, SW_MAXIMIZE
    
End Sub

Public Sub roll_back()
    On Error Resume Next
    '
    '   There may be nested DAO transactions.
    '
    While Err.Number = 0
        ws.Rollback
    Wend
    '
    '   There may be nested ADO transactions.
    '
    Err.Clear
    While Err.Number = 0
        gdbADO.RollbackTrans
    Wend
    '
    Err.Clear
    transaction_started = False
    gintTransCount = 0
End Sub

Public Sub DisplayError(Optional ByVal LogError As Boolean = True, _
                        Optional ByVal Source As String = "Source not supplied", _
                        Optional ByVal Severity As VbMsgBoxStyle = vbCritical)
    Dim intCount            As Integer
    Dim strCallStack()      As String
    Dim strOriginalSource   As String
    Dim strError            As String
    Dim strFullSource       As String
    Dim f                   As Form
    
    If Erl <> 0 Then
        Source = Source & " at line " & CStr(Erl)
    End If
    
    strFullSource = Err.Source & "~" & Source
    
    strCallStack = Split(strFullSource, "~")
    strOriginalSource = strCallStack(1)
    
    If Err.Number >= vbObjectError And Err.Number < vbObjectError + 65536 Then
        strError = "Application Error (" & CStr(Err.Number - vbObjectError) & " - " & Err.Description & ") in " & strOriginalSource
        MsgBox strError, Severity, "Image Where?"
        If LogError Then
            goLog.WriteLog strError, strFullSource
        End If
    Else
        If Err.Number = 3146 Then
            '
            '   Database Error.
            '
            intCount = 0
            While intCount <= DBEngine.Errors.Count
                strError = DBEngine.Errors(intCount) & " in " & strOriginalSource
                MsgBox strError, Severity, "Image Where?"
                If LogError Then
                    goLog.WriteLog strError, strFullSource
                End If
                intCount = intCount + 1
            Wend
        Else
            strError = CStr(Err.Number) & " - " & Err.Description & " in " & strOriginalSource
            MsgBox strError, Severity, "Image Where?"
            If LogError Then
                goLog.WriteLog strError, strFullSource
            End If
        End If
    End If
    '
    Screen.MousePointer = vbDefault
    '
    '   Rollback if required.
    '
    On Error Resume Next
    roll_back
    gdbADO.RollbackTrans
    '
    '   Unlock all windows.
    '
    For Each f In Forms
        UnlockWindow
    Next f
    gblnResizeMaxedWindows = True

End Sub

Public Function NVL(varNVL As Variant, Optional NullType As NVLType = 0) As String
    Dim intType As NVLType
    
    If NullType = 0 Then        ' Not entered.
        If VarType(varNVL) = vbString Then
            intType = NVLString
        Else
            intType = NVLNumeric
        End If
    Else
        intType = NullType
    End If
    
    If IsNull(varNVL) Then
        If intType = NVLString Then
            NVL = ""
        Else
            NVL = 0
        End If
    Else
        NVL = varNVL
    End If
    
End Function


Public Sub HighLightRow(msfGrid As MSFlexGrid, intHighlightRow As Integer)
    '
    '   If row does not exist then abort.
    '
    If intHighlightRow = 0 Or msfGrid.Rows = 1 Then
        Exit Sub
    ElseIf intHighlightRow >= msfGrid.Rows Then
        intHighlightRow = msfGrid.Rows - 1
    End If
    
    With msfGrid
        .col = 0            ' Highlight required row.
        .row = intHighlightRow
        .ColSel = .Cols - 1
        .RowSel = intHighlightRow
        If Not .RowIsVisible(.row) Then
            .TopRow = IIf(.row < 3, 1, .row - 2)
            '.TopRow = intHighlightRow - Int((.Height) / (2 * .RowHeight(1)))
        End If
    End With
End Sub

Public Function ContactGridPosition(msfGrid As MSFlexGrid, ID As Long) As Integer
    Dim intPosition As Integer
    
    On Error GoTo ErrorProc
    
    If ID <> 0 Then
        For intPosition = 1 To msfGrid.Rows - 1
            If msfGrid.TextMatrix(intPosition, 8) = ID Then
                Exit For
            End If
        Next intPosition
    End If
    
    ContactGridPosition = intPosition
    
    Exit Function
ErrorProc:
    Err.Raise Err.Number
End Function

Public Function LocateComboItem(cboLocate As ComboBox, strLocateText As String) As Boolean
    Dim intCount As Integer
    On Error GoTo ErrorProc
    
    LocateComboItem = False
    For intCount = 0 To cboLocate.ListCount - 1
        If Trim(cboLocate.List(intCount)) = Trim(strLocateText) Then
            cboLocate.ListIndex = intCount
            LocateComboItem = True
            Exit For
        End If
    Next intCount
    
    Exit Function

ErrorProc:
    Err.Raise Err.Number
End Function

Public Function StripPath(strFullFileName) As String
    Dim intChar As Integer
    
    On Error GoTo ErrorProc
    
    If Trim(strFullFileName) = "" Then
        StripPath = ""
        Exit Function
    End If
    
    For intChar = 1 To Len(strFullFileName)
        If Mid(strFullFileName, Len(strFullFileName) - intChar, 1) = "\" Then
            Exit For
        End If
    Next intChar
    StripPath = Right(strFullFileName, intChar)
    Exit Function
ErrorProc:
    Err.Raise Err.Number
End Function

Public Function StripFile(strFullFileName) As String
    Dim intChar As Integer
    
    On Error GoTo ErrorProc
    
    If Trim(strFullFileName) = "" Then
        StripFile = ""
        Exit Function
    End If
    
    For intChar = 1 To Len(strFullFileName)
        If Mid(strFullFileName, Len(strFullFileName) - intChar, 1) = "\" Then
            Exit For
        End If
    Next intChar
    StripFile = Left(strFullFileName, Len(strFullFileName) - intChar - 1)
    Exit Function
ErrorProc:
    Err.Raise Err.Number
End Function

Sub ConvertPIXELSToTWIPS(X As Long, Y As Long)

   '*************************************************************
   ' PURPOSE: Converts the two pixel measurements passed as
   '          arguments to twips.
   ' ARGUMENTS:
   '    X, Y: Measurement variables in pixels. These will be
   '          converted to twips and returned through the same
   '          variables "by reference."
   '*************************************************************
   Dim hDC As Long, hWnd As Long, retval As Long
   Dim XPIXELSPERINCH, YPIXELSPERINCH
   Const LOGPIXELSX = 88
   Const LOGPIXELSY = 90
   
   ' Retrieve the current number of pixels per inch, which is
   ' resolution-dependent.
   hDC = apiGetDC(0)
   XPIXELSPERINCH = apiGetDeviceCaps(hDC, LOGPIXELSX)
   YPIXELSPERINCH = apiGetDeviceCaps(hDC, LOGPIXELSY)
   retval = apiReleaseDC(0, hDC)
   
   ' Compute and return the measurements in twips.
   X = (X / XPIXELSPERINCH) * TWIPSPERINCH
   Y = (Y / YPIXELSPERINCH) * TWIPSPERINCH

End Sub

Public Function Duplicate_character(character_string As String, character_duplicate) As String
    Dim character_pos As Integer
    Dim WorkCharacterString
'
'   Add an extra character to each one in the customer name to
'   override the underscore functionality.
'
    WorkCharacterString = character_string
    character_pos = 1
    While Not character_pos = 0
        character_pos = InStr(character_pos, WorkCharacterString, character_duplicate)
        If Not character_pos = 0 Then
            WorkCharacterString = Mid(WorkCharacterString, 1, character_pos) + Mid(WorkCharacterString, character_pos)
            character_pos = character_pos + 2
        End If
    Wend

    Duplicate_character = WorkCharacterString
End Function

Public Function DAOToADOType(DAOType As Integer) As Integer
    Select Case DAOType
        Case Is = dbText
            DAOToADOType = adChar
        Case Is = DAO.dbInteger
            DAOToADOType = adInteger
    End Select
End Function

Public Sub ErrorRaise(strSource As String)

    Dim ErrSource   As String
    Dim oFTPError As FTPError
    
    If Erl <> 0 Then
        strSource = strSource & " at line " & CStr(Erl)
    End If
    
    ErrSource = Err.Source & "~" & strSource
    '
    '   Get the Full description of the FTP error.
    '
    If Left(Err.Description, 9) = "FTP Error" Then
        mstrErrorDescription = Err.Description
        If IsNumeric(Mid(Trim(mstrErrorDescription), 11)) Then
            mlngErrorNumber = CLng(Mid(Trim(mstrErrorDescription), 11))
        Else
            mlngErrorNumber = Err.Number
        End If
        Set oFTPError = New FTPError
        On Error Resume Next
        oFTPError.Read mlngErrorNumber
        If Err.Number = 0 Then
            mstrErrorDescription = oFTPError.Description
        End If
        On Error GoTo 0
        Err.Raise mlngErrorNumber, ErrSource, mstrErrorDescription
    Else
        Err.Raise Err.Number, ErrSource
    End If
End Sub

Public Function SetTopMostWindow(TopForm As Form, TopMost As Boolean) As Long

    If TopMost = True Then 'Make the window topmost
       SetTopMostWindow = SetWindowPos(TopForm.hWnd, HWND_TOPMOST, 0, 0, 0, _
          0, FLAGS)
    Else
       SetTopMostWindow = SetWindowPos(TopForm.hWnd, HWND_NOTOPMOST, 0, 0, _
          0, 0, FLAGS)
       SetTopMostWindow = False
    End If

End Function

Public Sub ErrorSave()
    mlngErrorNumber = Err.Number
    mstrErrorDescription = Err.Description
    mstrSource = Err.Source
End Sub

Public Sub ErrorRestore(Optional strSource As String = "")
    If strSource = "" Then
        Err.Raise mlngErrorNumber, mstrSource, mstrErrorDescription
    Else
        mstrSource = mstrSource & "~" & strSource
        Err.Raise mlngErrorNumber, mstrSource, mstrErrorDescription
    End If
End Sub

Public Sub WriteLog(strRecord As String, Optional strErrorNo As String = "")
    'On Error Resume Next
    Print #1, """" & strErrorNo & """,""" & _
              strRecord & """,""" & _
              Format(Now, "General Date") & """"
End Sub

Function JetSQLFixup(TextIn)
    '***************************************
    ' Module/Form Name   : common_routines
    '
    ' Procedure Name     : JetSQLFixup
    '
    ' Purpose            :
    '
    ' Date Created       : 09/12/2001
    '
    ' Author             : GARETH
    '
    ' Parameters         : TextIn -
    '
    ' Returns            :
    '
    ' Amendment History  : Date       Author    Description
    '                    : --------------------------------
    '                    : 09/12/2001 GARETH
    '
    '***************************************
    '
On Error GoTo JetSQLFixup_Error
    '
    '******** Code Starts Here *************
    '
    Dim temp
temp = Replace(TextIn, "'", "''")
temp = Replace(temp, "|", "' & chr(124) & '")
    JetSQLFixup = Replace(temp, "%", "[%]")
    '
    '********* Code Ends Here **************
    '
Exit Function
    '
JetSQLFixup_Error:
    ErrorRaise "common_routines.JetSQLFixup"
End Function

Public Function Email(strTo As String, _
                      strSubject As String, _
                      strBody As String, _
                      Optional colAttachments As Collection) As Boolean
'***************************************
' Module/Form Name   : common_routines
'
' Procedure Name     : Email
'
' Purpose            :
'
' Date Created       : 18/04/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : strTo - String
'                    : strSubject - String
'                    : strBody - String
'
' Returns            : Boolean
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 18/04/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo Email_Error
'
'******** Code Starts Here *************
'
    Dim strFile As Variant
    'Dim objOutlook As Outlook.Application
    Dim objOutlook As Object
    'Dim objItem As Outlook.MailItem
    Dim objItem As Object
    'Dim colAttach As Outlook.Attachments
    Dim colAttach As Object
    'Dim oMailAttach As Outlook.Attachment
    Dim oMailAttach As Object

    Email = False
    'Create a Microsoft Outlook object.
    'Set objOutlook = New Outlook.Application
    
    'Create and open a new contact form for input.
    'Set objItem = objOutlook.CreateItem(olMailItem)
    Set objItem = objOutlook.CreateItem(0)
    With objItem
        .To = strTo
        .Subject = strSubject
        .Body = strBody
        If Not IsMissing(colAttachments) Then
            Set colAttach = objItem.Attachments
            For Each strFile In colAttachments
                Set oMailAttach = colAttach.Add(strFile)
            Next strFile
        End If
        .Send
    End With
    
    'Quit Microsoft Outlook.
    Set objOutlook = Nothing
    Email = True

    Exit Function
    '
    '********* Code Ends Here **************
    '
Email_Error:
    ErrorRaise "common_routines.Email"
End Function

Public Sub MakeAllControlsInvisible(invisibleControls As Form)
    Dim Ctl As Object
    
    On Error Resume Next
    For Each Ctl In invisibleControls.Controls
        Ctl.Visible = False
    Next Ctl

End Sub

Public Sub MakeAllControlsVisible(visibleControls As Form, invisibleControls As Collection)
    Dim Ctl As Object
    Dim InvisCtl As Object
    Dim makeVisible As Boolean
        
    On Error Resume Next
    For Each Ctl In visibleControls.Controls
        '
        '   Does the value exist in the Invisible Controls collection
        '
        makeVisible = True
        For Each InvisCtl In invisibleControls
            If Ctl.Name = InvisCtl.Name Then
                makeVisible = False
                Exit For
            End If
        Next InvisCtl
        
        If makeVisible Then
            Ctl.Visible = True
        End If
    Next Ctl

End Sub

Public Sub LockWindow(hWnd As Long)
    LockWindowUpdate hWnd
End Sub

Public Sub UnlockWindow()
    LockWindowUpdate 0
End Sub

Public Function ApplicationVersion() As String
    ApplicationVersion = App.Major & "." & Format(App.Minor, "00") & "." & Format(App.Revision, "0000")
End Function

Public Function IsInIDE() As Boolean
   Dim X As Long
   Debug.Assert Not TestIDE(X)
   IsInIDE = X = 1
End Function

Private Function TestIDE(X As Long) As Boolean
   X = 1
End Function

Public Function ChildWindowsMaxed() As Boolean
    Dim f As Form
    Dim blnChildWindowMaxed As Boolean
    
    blnChildWindowMaxed = False
    For Each f In Forms
        If f.Name <> "mdi_npls" Then
            If f.WindowState = vbMaximized Then
                blnChildWindowMaxed = True
                Exit For
            End If
        End If
    Next f

    ChildWindowsMaxed = blnChildWindowMaxed
End Function

Public Function ResizeForm(frmResize As Form) As Boolean
'***************************************
' Module/Form Name   : common_routines
'
' Procedure Name     : ResizeForm
'
' Purpose            :
'
' Date Created       : 19/04/2003
'
' Author             : GARETH SAUNDERS
'
' Parameters         : frmResize - Form
'
' Returns            : Boolean
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo ResizeForm_Error
'
'******** Code Starts Here *************
'
    ResizeForm = False
    '
    '   Resizing Rules for MDI Children.
    '
    ResizeMaxedWindows frmResize
    
    If frmResize.WindowState <> vbMaximized And _
       ChildWindowsMaxed Then
        Exit Function
    End If
    '
    If frmResize.WindowState <> vbMaximized And _
       Not ChildWindowsMaxed And _
       Not FormResizedToMax(frmResize) And gcolMaxedWindows.Count <> 0 Then
        Exit Function
    End If
    '
    If frmResize.WindowState = vbMaximized Then
        If FormResizedToMax(frmResize) Then
            Exit Function
        Else
            gcolMaxedWindows.Add frmResize, CStr(frmResize.hWnd)
        End If
    End If
    
    If gblnResizeMaxedWindows Then
        If Not ChildWindowsMaxed Then
            '
            '   Empty the collection.
            '
            While Not gcolMaxedWindows.Count = 0
                gcolMaxedWindows.Remove 1
            Wend
            Set gcolMaxedWindows = Nothing
            Set gcolMaxedWindows = New Collection
        End If
    End If
    
    ResizeForm = True
'
'********* Code Ends Here **************
'
   Exit Function
'
ResizeForm_Error:
    ErrorRaise "common_routines.ResizeForm"
End Function

Public Sub ResizeMaxedWindows(frmResize As Form)
'***************************************
' Module/Form Name   : common_routines
'
' Procedure Name     : ResizeMaxedWindows
'
' Purpose            :
'
' Date Created       : 20/04/2003
'
' Author             : GARETH SAUNDERS
'
' Parameters         : frmResize - Form
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo ResizeMaxedWindows_Error
'
'******** Code Starts Here *************
'
    Dim f As Form
    Dim f2 As Form
    Dim blnFound As Boolean
    
    If gblnResizeMaxedWindows Then
        If Not ChildWindowsMaxed Then
            For Each f In gcolMaxedWindows
                '
                '   Is the form still loaded?
                '
                blnFound = False
                For Each f2 In Forms
                    If f2 Is f Then
                        blnFound = True
                        Exit For
                    End If
                Next f2
                '
                If blnFound Then
                    If f.hWnd <> frmResize.hWnd Then
                        On Error Resume Next
                        f.ForceResize
                        gcolMaxedWindows.Remove CStr(f.hWnd)
                        On Error GoTo ResizeMaxedWindows_Error
                    End If
                End If
            Next f
        End If
    End If
'
'********* Code Ends Here **************
'
   Exit Sub
'
ResizeMaxedWindows_Error:
    ErrorRaise "common_routines.ResizeMaxedWindows"
End Sub

Private Function FormResizedToMax(frmMax As Form) As Boolean
'***************************************
' Module/Form Name   : common_routines
'
' Procedure Name     : FormResizedToMax
'
' Purpose            :
'
' Date Created       : 20/04/2003
'
' Author             : GARETH SAUNDERS
'
' Parameters         : frmMax - Form
'
' Returns            : Boolean
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo FormResizedToMax_Error
'
'******** Code Starts Here *************
'
    Dim f As Form
    
    FormResizedToMax = False
    
    On Error Resume Next
    Set f = gcolMaxedWindows.Item(CStr(frmMax.hWnd))
    If Err.Number <> 0 Then
        Exit Function
    End If
    
    FormResizedToMax = True
'
'********* Code Ends Here **************
'
   Exit Function
'
FormResizedToMax_Error:
    ErrorRaise "common_routines.FormResizedToMax"
End Function

Public Function RemoveIllegalFileAndFolderCharacters(ByVal pstrFileOrFolder As String) As String
'***************************************
' Module/Form Name   : common_routines
'
' Procedure Name     : RemoveIllegalFileAndFolderCharacters
'
' Purpose            :
'
' Date Created       : 01/07/2004
'
' Author             : ADMINISTRATOR
'
' Parameters         : pstrFileOrFolder - String
'
' Returns            : String
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo RemoveIllegalFileAndFolderCharacters_Error
'
'******** Code Starts Here *************
'
    Dim strPath     As String
    
    strPath = Replace(pstrFileOrFolder, "\", "")
    strPath = Replace(strPath, ":", "")
    strPath = Replace(strPath, "/", "")
    strPath = Replace(strPath, "*", "")
    strPath = Replace(strPath, "?", "")
    strPath = Replace(strPath, """", "")
    strPath = Replace(strPath, "<", "")
    strPath = Replace(strPath, ">", "")
    strPath = Replace(strPath, "|", "")
    
    RemoveIllegalFileAndFolderCharacters = strPath
'
'********* Code Ends Here **************
'
   Exit Function
'
RemoveIllegalFileAndFolderCharacters_Error:
    ErrorRaise "common_routines.RemoveIllegalFileAndFolderCharacters"
End Function

Public Function TopLevelControl(ByRef pfrmTop As Form, _
                                 ByRef pctl As Control) As Control
    Dim Ctl         As Control
    Dim ctlParent   As Control
    
    Set Ctl = pctl
    On Error Resume Next
    Set ctlParent = Ctl.Container
    If Err.Number = 0 Then
        Do While Not ctlParent.Container.Name = pfrmTop.Name
            Set Ctl = Ctl.Container
            Set ctlParent = ctlParent.Container
            If Err.Number <> 0 Then
                Exit Do
            End If
        Loop
    End If
    Set TopLevelControl = Ctl
End Function



