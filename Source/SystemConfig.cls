VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SystemConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarDirty                           As Boolean
Private mvarServerLocation                  As String
Private mvarImageHoldingArea                As String
Private mvarDefaultImageLocation            As String
Private mvarSageLink                        As Boolean
Private mvarCaptionHeight                   As Long
Private mvarVScrollBarWidth                 As Long
Private mvarPostedImages                    As Collection
Private mvarConfirmationImages              As Collection
Private mvarCancellationImages              As Collection
Private mvarPostedHTMLEmail                 As String
Private mvarConfirmationHTMLEmail           As String
Private mvarCancellationHTMLEmail           As String
Private mvarBundledHTMLEmail                As String
Private mvarDateLogFilePosted               As Date
Private mvarDatePreviousPostWarning         As Date
Private mvarDateHTMLEmailPosted             As Date
Private mvarLogFilePostingFrequency         As Integer
Private mvarSupportEmail                    As String
Private mvarWebSearchTestEmail              As String
Private mvarWebLocation                     As String
Private mvarHTMLEmailSupportRequestDate     As Date
Private mvarDateHTMLEmailUpgraded           As Date
Private mvarDiagnosticsRequestDate          As Date
Private mvarDiagnosticSQL                   As String
Private mvarDateDiagnosticsPosted           As Date
Private mvarSupportUser                     As String
Private mvarUpgradeCheckFrequency           As Integer
Private mvarLocations                       As Locations
Private mvarDetailedAuditing                As Boolean
Private mvarDateDetailedAuditingSet         As Date
Private mvarDateDetailedAuditingRequested   As Date
Private mvarDetailedAuditingRequest         As Boolean
Private mvarFTPUser                         As String       '   "clivenichols"
Private mvarFTPPassword                     As String       '   "clin88"
Private mvarFTPServer                       As String       '   "ftp.transcom.com"
Private mvarPostWebAddress                  As String       '   www.clivenichols.com
Private mstrIniFile                         As String
Private mstrServerIniFile                   As String
Private mvarWebSearchesEmailTo              As String
Private mvarWebSearchesEmailFrom            As String
Private mvarFuzzyKeywordSearch              As Boolean
Private mvarBasicImageWhere                 As Boolean
Private mvarMouseWheel                      As Boolean
Private mvarTooltipDelay                    As Integer
Private mvarLastPhotoNumberUsed             As Long
'
'   System metrics.
'
Private Declare Function GetSystemMetrics% Lib "user32" (ByVal nIndex%)
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Const SM_CYCAPTION = 4
Private Const SM_CXVSCROLL = 2
Private Const MAX_NO_OF_THUMBS = 5000
Private Const DEFAULT_NO_OF_THUMBS = 500
Private Const MAX_IMAGES_PER_PAGE = 500
Private Const DEFAULT_IMAGES_PER_PAGE = 20
Private Const DEFAULT_IMAGE_HEIGHT_WIDTH = 1500
Private Const MAX_IMAGE_HEIGHT_WIDTH = 2268
Private Const MIN_IMAGE_HEIGHT_WIDTH = 850
Private Const DEFAULT_OFFICE_VERSION = "97"
'
'   Active Internet Connection variables and APIs.
'
Private Const HKEY_LOCAL_MACHINE As Long = &H80000002
Private Const HKEY_USERS As Long = &H80000003

'Return codes from Registration functions
Private Const ERROR_SUCCESS As Long = 0
Private Const ERROR_BADDB As Long = 1009
Private Const ERROR_BADKEY As Long = 1010
Private Const ERROR_CANTOPEN As Long = 1011
Private Const ERROR_CANTREAD As Long = 1012
Private Const ERROR_CANTWRITE As Long = 1013
Private Const ERROR_OUTOFMEMORY As Long = 14
Private Const ERROR_INVALID_PARAMETER As Long = 87
Private Const ERROR_ACCESS_DENIED As Long = 5
Private Const ERROR_MORE_DATA As Long = 234
Private Const ERROR_NO_MORE_ITEMS As Long = 259
Private Const KEY_ALL_ACCESS As Long = &H3F
Private Const REG_OPTION_NON_VOLATILE As Long = 0

Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private Declare Function RegOpenKey Lib "advapi32" _
   Alias "RegOpenKeyA" _
  (ByVal hKey As Long, _
   ByVal sSubKey As String, _
   hKey As Long) As Long

Private Declare Function RegQueryValueEx Lib "advapi32" _
   Alias "RegQueryValueExA" _
  (ByVal hKey As Long, _
   ByVal sKeyValue As String, _
   ByVal lpReserved As Long, _
   lpType As Long, _
   lpData As Any, _
   nSizeData As Long) As Long

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Property Get Dirty() As Boolean
    Dirty = mvarDirty
End Property

Public Property Get ServerLocation() As String
    ServerLocation = mvarServerLocation
End Property

Public Property Let ServerLocation(ByVal vData As String)
    mvarServerLocation = vData
    mvarDirty = True
End Property

Public Property Get FuzzyKeywordSearch() As Boolean
    FuzzyKeywordSearch = mvarFuzzyKeywordSearch
End Property

Public Property Let FuzzyKeywordSearch(ByVal vData As Boolean)
    mvarFuzzyKeywordSearch = vData
    mvarDirty = True
End Property

Public Property Get BasicImageWhere() As Boolean
    BasicImageWhere = mvarBasicImageWhere
End Property

Public Property Let BasicImageWhere(ByVal vData As Boolean)
    mvarBasicImageWhere = vData
    mvarDirty = True
End Property

Public Property Let DefaultImageLocation(ByVal vData As String)
    mvarDefaultImageLocation = vData
    mvarDirty = True
End Property

Public Property Get DefaultImageLocation() As String
    DefaultImageLocation = mvarDefaultImageLocation
End Property

Public Property Let ImageHoldingArea(ByVal vData As String)
    mvarImageHoldingArea = vData
    mvarDirty = True
End Property

Public Property Get ImageHoldingArea() As String
    ImageHoldingArea = mvarImageHoldingArea
End Property

Public Property Get TitleBarHeight() As Long
    TitleBarHeight = ConvertPIXELSToTWIPS(mvarCaptionHeight)
End Property

Public Property Get VScrollBarWidth() As Long
    VScrollBarWidth = ConvertPIXELSToTWIPS(mvarVScrollBarWidth)
End Property

Public Property Get SageLink() As Boolean
    SageLink = mvarSageLink
End Property

Public Property Get ImagesReturned() As Long
    Dim strImagesReturned As String
    
    strImagesReturned = GetSetting(App.Title, "Options", "ImagesReturned", MAX_NO_OF_THUMBS)
    If Not IsNumeric(strImagesReturned) Then
        MsgBox "Stored value for 'Images Returned' is corrupt." & vbCrLf & "Resetting to " & CStr(DEFAULT_NO_OF_THUMBS), vbExclamation
        SaveSetting App.Title, "Options", "ImagesReturned", CStr(DEFAULT_NO_OF_THUMBS)
        strImagesReturned = CStr(DEFAULT_NO_OF_THUMBS)
    End If
    ImagesReturned = CLng(strImagesReturned)
End Property

Public Property Let ImagesReturned(vData As Long)
    SaveSetting App.Title, "Options", "ImagesReturned", CStr(vData)
    mvarDirty = True
End Property

Public Property Get ImageHeightWidth() As Double '   In Twips.
    Dim strImageHeightWidth As String
    
    strImageHeightWidth = GetSetting(App.Title, "Options", "ImageHeightWidth", DEFAULT_IMAGE_HEIGHT_WIDTH)
    If Not IsNumeric(strImageHeightWidth) Then
        MsgBox "Stored value for 'Image Height' is corrupt." & vbCrLf & "Resetting to default.", vbExclamation
        SaveSetting App.Title, "Options", "ImagesHeight", CStr(DEFAULT_IMAGE_HEIGHT_WIDTH)
        strImageHeightWidth = CStr(DEFAULT_IMAGE_HEIGHT_WIDTH)
    End If
    ImageHeightWidth = CDbl(strImageHeightWidth)
End Property

Public Property Get ImageHeightWidthCms() As Double '   In Cms.
    ImageHeightWidthCms = ((ImageHeightWidth / 1440) * 254) / 100
End Property

Public Property Let ImageHeightWidthCms(vData As Double) '   In Cms.
    SaveSetting App.Title, "Options", "ImageHeightWidth", CStr((vData / 2.54) * 1440)
    mvarDirty = True
End Property

Public Property Get ImagesPerPage() As Long
    Dim strImagesPerPage As String
    
    strImagesPerPage = GetSetting(App.Title, "Options", "MaxImagesPerPage", MAX_IMAGES_PER_PAGE)
    If Not IsNumeric(strImagesPerPage) Then
        MsgBox "Stored value for 'Maximum Images perPage' is corrupt." & vbCrLf & "Resetting to " & CStr(DEFAULT_IMAGES_PER_PAGE), vbExclamation
        SaveSetting App.Title, "Options", "MaxImagesPerPage", CStr(DEFAULT_IMAGES_PER_PAGE)
        strImagesPerPage = CStr(DEFAULT_IMAGES_PER_PAGE)
    End If
    ImagesPerPage = CLng(strImagesPerPage)
End Property

Public Property Let ImagesPerPage(vData As Long)
    SaveSetting App.Title, "Options", "MaxImagesPerPage", CStr(vData)
    mvarDirty = True
End Property

Public Property Get MaxImagesPerSearch() As Integer
    MaxImagesPerSearch = MAX_NO_OF_THUMBS
End Property

Public Property Get MaxImagesPerPage() As Integer
    MaxImagesPerPage = MAX_IMAGES_PER_PAGE
End Property

Public Property Get MaxImageHeightWidth() As Double    '   In Cms.
    MaxImageHeightWidth = Round(((MAX_IMAGE_HEIGHT_WIDTH / 1440) * 2.54), 2)
End Property

Public Property Get MinImageHeightWidth() As Double
    MinImageHeightWidth = Round(((MIN_IMAGE_HEIGHT_WIDTH / 1440) * 2.54), 2)
End Property

Public Property Get PostedHTMLEmail() As String
    PostedHTMLEmail = mvarPostedHTMLEmail
End Property

Public Property Let PostedHTMLEmail(vData As String)
    mvarPostedHTMLEmail = vData
    mvarDirty = True
End Property

Public Property Get ConfirmationHTMLEmail() As String
    ConfirmationHTMLEmail = mvarConfirmationHTMLEmail
End Property

Public Property Let ConfirmationHTMLEmail(vData As String)
    mvarConfirmationHTMLEmail = vData
    mvarDirty = True
End Property

Public Property Get CancellationHTMLEmail() As String
    CancellationHTMLEmail = mvarCancellationHTMLEmail
End Property

Public Property Let BundledHTMLEmail(vData As String)
    mvarBundledHTMLEmail = vData
    mvarDirty = True
End Property

Public Property Get BundledHTMLEmail() As String
    BundledHTMLEmail = mvarBundledHTMLEmail
End Property

Public Property Let CancellationHTMLEmail(vData As String)
    mvarCancellationHTMLEmail = vData
    mvarDirty = True
End Property

Private Sub Class_Initialize()
    Dim oFSO As New Scripting.FileSystemObject
    
    On Error GoTo ErrorProc
    '
    mvarSupportUser = False
    mvarDetailedAuditing = False
    mstrIniFile = App.Path & "\IW.ini"
    '
    With oFSO
        If Not .FileExists(mstrIniFile) Then
            Open mstrIniFile For Output As #1
            Close #1
        End If
    End With
    '
    Set oFSO = Nothing
    '
    Exit Sub
ErrorProc:
    Err.Raise Err.Number, "System Config: Initialise", Err.Description
End Sub

Private Function ConvertPIXELSToTWIPS(Y As Long) As Long
   Dim hDC As Long, hWnd As Long, retval As Long
   Dim YPIXELSPERINCH
   Const LOGPIXELSY = 90
   ' Retrieve the current number of pixels per inch, which is
   ' resolution-dependent.
   hDC = apiGetDC(0)
   YPIXELSPERINCH = apiGetDeviceCaps(hDC, LOGPIXELSY)
   retval = apiReleaseDC(0, hDC)
   ' Compute and return the measurements in twips.
   ConvertPIXELSToTWIPS = (Y / YPIXELSPERINCH) * TWIPSPERINCH
End Function

Public Property Get ComputerName() As String
          Dim intResponse As Integer
          Dim strComputerName As String
          Dim lngSize As Long

          strComputerName = Space(16)
          lngSize = 16
          intResponse = GetComputerName(strComputerName, lngSize)
          ComputerName = Left(strComputerName, lngSize)
End Property

Public Property Get UserName() As String
    UserName = VBA.Environ("USERNAME")
End Property

Public Property Get PostedImages() As Collection
    Set PostedImages = mvarPostedImages
End Property

Public Property Get ConfirmationImages() As Collection
    Set ConfirmationImages = mvarConfirmationImages
End Property

Public Property Get CancellationImages() As Collection
    Set CancellationImages = mvarCancellationImages
End Property

Public Property Let DateLogFilePosted(vData As Date)
    mvarDateLogFilePosted = vData
    mvarDirty = True
End Property

Public Property Get DateLogFilePosted() As Date
    DateLogFilePosted = mvarDateLogFilePosted
End Property

Public Property Let DatePreviousPostWarning(vData As Date)
    mvarDatePreviousPostWarning = vData
    mvarDirty = True
End Property

Public Property Get DatePreviousPostWarning() As Date
    DatePreviousPostWarning = mvarDatePreviousPostWarning
End Property

Public Property Get DateHTMLEmailPosted() As Date
    DateHTMLEmailPosted = mvarDateHTMLEmailPosted
End Property

Public Property Let DateHTMLEmailPosted(vData As Date)
    mvarDateHTMLEmailPosted = vData
    mvarDirty = True
End Property

Public Property Get DateDiagnosticsPosted() As Date
    DateDiagnosticsPosted = mvarDateDiagnosticsPosted
End Property

Public Property Let DateDiagnosticsPosted(vData As Date)
    mvarDateDiagnosticsPosted = vData
    mvarDirty = True
End Property

Public Property Get DateHTMLEmailUpgraded() As Date
    DateHTMLEmailUpgraded = mvarDateHTMLEmailUpgraded
End Property

Public Property Let DateHTMLEmailUpgraded(vData As Date)
    mvarDateHTMLEmailUpgraded = vData
    mvarDirty = True
End Property

Public Property Let LogFilePostingFrequency(vData As Integer)
    mvarLogFilePostingFrequency = vData
    mvarDirty = True
End Property

Public Property Get LogFilePostingFrequency() As Integer
    LogFilePostingFrequency = mvarLogFilePostingFrequency
End Property

Public Property Let SupportEmail(vData As String)
    mvarSupportEmail = vData
    mvarDirty = True
End Property

Public Property Get SupportEmail() As String
    SupportEmail = mvarSupportEmail
End Property

Public Property Let WebSearchTestEmail(vData As String)
    mvarWebSearchTestEmail = vData
    If mvarWebSearchTestEmail = "" Then
        mvarWebLocation = "/websearches"
    Else
        mvarWebLocation = "/websearches/test"
    End If
    mvarDirty = True
End Property

Public Property Get WebSearchTestEmail() As String
    WebSearchTestEmail = mvarWebSearchTestEmail
End Property

Public Property Get WebLocation() As String
    WebLocation = mvarWebLocation
End Property

Public Property Get HTMLEmailSupportRequestDate() As Date
    HTMLEmailSupportRequestDate = mvarHTMLEmailSupportRequestDate
End Property

Public Property Get DiagnosticsRequestDate() As Date
    DiagnosticsRequestDate = mvarDiagnosticsRequestDate
End Property

Public Property Get DiagnosticSQL() As String
    DiagnosticSQL = mvarDiagnosticSQL
End Property

Public Property Let DiagnosticSQL(vData As String)
    mvarDiagnosticSQL = vData
    mvarDirty = True
End Property

Public Property Get SupportUser() As Boolean
    SupportUser = mvarSupportUser
End Property

Public Property Let SupportUser(vData As Boolean)
    mvarSupportUser = vData
    mvarDirty = True
End Property

Public Property Get MouseWheel() As Boolean
    MouseWheel = mvarMouseWheel
End Property

Public Property Let MouseWheel(vData As Boolean)
    mvarMouseWheel = vData
    mvarDirty = True
End Property

Public Property Get TooltipDelay() As Integer
    TooltipDelay = mvarTooltipDelay
End Property

Public Property Let TooltipDelay(vData As Integer)
    mvarTooltipDelay = vData
    mvarDirty = True
End Property

Public Property Get LastPhotoNumberUsed() As Long
    LastPhotoNumberUsed = mvarLastPhotoNumberUsed
End Property

Public Property Let LastPhotoNumberUsed(vData As Long)
    mvarLastPhotoNumberUsed = vData
    mvarDirty = True
End Property

Public Sub Refresh()
'***************************************
' Module/Form Name   : SystemConfig
'
' Procedure Name     : Refresh
'
' Purpose            :
'
' Date Created       : 02/06/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 02/06/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo Refresh_Error
'
'******** Code Starts Here *************
'
    Dim retval              As Long
    Dim strSQL              As String
    Dim strBuffer           As String
    Dim rstSystemConfig     As ADODB.Recordset
    Dim oFSO                As Scripting.FileSystemObject
    Dim strDetailedAuditing As String
    Dim objUpgradeChecker   As Object

    On Error Resume Next
    Set objUpgradeChecker = CreateObject("IWUpgradeChecker2.MainControl")
    If Err.Number = 0 Then
        mvarSupportUser = True
    Else
        mvarSupportUser = False
    End If
    Set objUpgradeChecker = Nothing
    '
    strSQL = "SELECT * FROM SYSTEMCONFIG"
        Set rstSystemConfig = New ADODB.Recordset
    rstSystemConfig.Open strSQL, gdbADO, adOpenStatic
    '
    With rstSystemConfig
        If .EOF Then
            Err.Raise vbObjectError + 1, "SystemConfig: Initialize", "System Configuration Record Missing"
        Else
            .MoveLast
            .MoveFirst
            If .RecordCount > 1 Then
                Err.Raise vbObjectError + 2, , "More than one System Configuration record found"
            Else
                .MoveFirst
            End If
        End If
    
        mvarDefaultImageLocation = NVL(!DefImageFileLocation, NVLString)
        mvarImageHoldingArea = NVL(!HoldingAreaLocation, NVLString)
        mvarSageLink = NVL(!SageLink, NVLString)
        mvarPostedHTMLEmail = NVL(!PostedHTMLEmail, NVLString)
        WebSearchTestEmail = NVL(!WebSearchTestEmail, NVLString)
        mvarConfirmationHTMLEmail = NVL(!ConfirmationHTMLEmail, NVLString)
        mvarCancellationHTMLEmail = NVL(!CancellationHTMLEmail, NVLString)
        mvarBundledHTMLEmail = NVL(!BundledHTMLEmail, NVLString)
        mvarDateLogFilePosted = NVL(!DateLogFilePosted, NVLNumeric)
        mvarLogFilePostingFrequency = NVL(!LogFilePostingFrequency, NVLNumeric)
        mvarSupportEmail = Trim(NVL(!SupportEmail, NVLString))
    End With
    '
    '   registry Settings.
    '
    mvarDateHTMLEmailPosted = CDate(GetSetting("Image Where", "Options", "DateHTMLEmailPosted", 0))
    mvarDateHTMLEmailUpgraded = CDate(GetSetting("Image Where", "Options", "DateHTMLEmailUpgraded", 0))
    mvarDateDiagnosticsPosted = CDate(GetSetting("Image Where", "Options", "DateDiagnosticsPosted", 0))
    mvarUpgradeCheckFrequency = CInt(GetSetting("Image Where", "Upgrade", "CheckInterval", "60"))
    mvarDateDetailedAuditingSet = CDate(GetSetting("Image Where", "Options", "DateDetailedAuditingSet", 0))
    mvarDatePreviousPostWarning = CDate(GetSetting("Image Where", "Options", "DatePreviousPostWarning", 0))
    '
    '
    '   Check that the folders exist.
    '
    On Error Resume Next
    Set oFSO = New Scripting.FileSystemObject
    With oFSO
        'If Not .FolderExists(.BuildPath(App.Path, mvarDefaultImageLocation)) Then
        If Not .FolderExists(.BuildPath(goSystemConfig.ServerLocation, mvarDefaultImageLocation)) Then
            On Error GoTo 0
            'Err.Raise vbObjectError + 3, , "Default Image Location does not exist: " & App.Path & "\" & mvarDefaultImageLocation
            Err.Raise vbObjectError + 3, , "Default Image Location does not exist: " & goSystemConfig.ServerLocation & "\" & mvarDefaultImageLocation
        End If
        '
        'If Not .FolderExists(.BuildPath(App.Path, mvarImageHoldingArea)) Then
        If Not .FolderExists(.BuildPath(goSystemConfig.ServerLocation, mvarImageHoldingArea)) Then
            On Error GoTo 0
            '.CreateFolder .BuildPath(App.Path, mvarImageHoldingArea)
            .CreateFolder .BuildPath(goSystemConfig.ServerLocation, mvarImageHoldingArea)
            'MsgBox "Image Holding Area has been created:" & vbCrLf & App.Path & "\" & mvarImageHoldingArea, vbInformation
            MsgBox "Image Holding Area has been created:" & vbCrLf & goSystemConfig.ServerLocation & "\" & mvarImageHoldingArea, vbInformation
        End If
    End With

    rstSystemConfig.Close
    Set rstSystemConfig = Nothing
    '
    '   Set up the system metrics.
    '
    mvarVScrollBarWidth = GetSystemMetrics(SM_CXVSCROLL)
    mvarCaptionHeight = GetSystemMetrics(SM_CYCAPTION)
    '
    '   Set up the HTML Images collection for use in emails.
    '
    SetUpHTMLImages
    '
    If mvarSupportUser Then
        SupportUserFlags
    End If
    '
    strBuffer = Space(10)
    retval = GetPrivateProfileString("Support", "DetailedAuditing", "", strBuffer, 10, mstrServerIniFile)
    strDetailedAuditing = Left(strBuffer, InStr(strBuffer, Chr(0)) - 1)
    mvarDetailedAuditing = CBool(strDetailedAuditing)
    '
    strBuffer = Space(100)
    retval = GetPrivateProfileString("FTP", "Server", "", strBuffer, 100, mstrServerIniFile)
    mvarFTPServer = Left(strBuffer, InStr(strBuffer, Chr(0)) - 1)
    '
    strBuffer = Space(100)
    retval = GetPrivateProfileString("FTP", "User", "", strBuffer, 100, mstrServerIniFile)
    mvarFTPUser = Left(strBuffer, InStr(strBuffer, Chr(0)) - 1)
    '
    strBuffer = Space(50)
    retval = GetPrivateProfileString("FTP", "Password", "", strBuffer, 50, mstrServerIniFile)
    mvarFTPPassword = Left(strBuffer, InStr(strBuffer, Chr(0)) - 1)
    '
    strBuffer = Space(100)
    retval = GetPrivateProfileString("FTP", "PostWebAddress", "", strBuffer, 100, mstrServerIniFile)
    mvarPostWebAddress = Left(strBuffer, InStr(strBuffer, Chr(0)) - 1)
    '
    strBuffer = Space(50)
    retval = GetPrivateProfileString("FTP", "WebSearchesEmailTo", "", strBuffer, 50, mstrServerIniFile)
    mvarWebSearchesEmailTo = Left(strBuffer, InStr(strBuffer, Chr(0)) - 1)
    '
    strBuffer = Space(50)
    retval = GetPrivateProfileString("FTP", "WebSearchesEmailFrom", "", strBuffer, 50, mstrServerIniFile)
    mvarWebSearchesEmailFrom = Left(strBuffer, InStr(strBuffer, Chr(0)) - 1)
    '
    strBuffer = Space(10)
    retval = GetPrivateProfileString("General", "FuzzyKeywordSearch", "", strBuffer, 10, mstrServerIniFile)
    mvarFuzzyKeywordSearch = CBool(Left(strBuffer, InStr(strBuffer, Chr(0)) - 1))
    '
    strBuffer = Space(10)
    retval = GetPrivateProfileString("General", "BasicImageWhere", "", strBuffer, 10, mstrServerIniFile)
    mvarBasicImageWhere = CBool(Left(strBuffer, InStr(strBuffer, Chr(0)) - 1))
    '
    mvarMouseWheel = CBool(GetSetting("Image Where", "Options", "MouseWheel", "True"))
    '
    mvarTooltipDelay = CInt(GetSetting("Image Where", "Options", "TooltipDelay", "1"))
    '
    Set oFSO = Nothing
    '
    '********* Code Ends Here **************
    '
    Exit Sub
    '
Refresh_Error:
    ErrorRaise "SystemConfig.Refresh"
End Sub

Private Sub SupportUserFlags()
'***************************************
' Module/Form Name   : SystemConfig
'
' Procedure Name     : SupportUserFlags
'
' Purpose            :
'
' Date Created       : 08/07/2004
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo SupportUserFlags_Error
'
'******** Code Starts Here *************
'
    Dim strSupportFile      As String
    Dim strBuffer           As String
    Dim retval              As Long
    
    strSupportFile = gstrUpgradePath & "\iwversion.ini"
    '
    '   Read the upgrade folder's 'iwversion.ini' file for further fields.
    '
    strBuffer = Space(20)
    retval = GetPrivateProfileString("Support", "DateDetailedAuditingRequested", "", strBuffer, 20, strSupportFile)
    strBuffer = Left(strBuffer, InStr(strBuffer, Chr(0)) - 1)
    If IsDate(strBuffer) Then
        mvarDateDetailedAuditingRequested = CDate(strBuffer)
    Else
        mvarDateDetailedAuditingRequested = 0
    End If
    '
    strBuffer = Space(20)
    retval = GetPrivateProfileString("Support", "DetailedAuditingRequest", "", strBuffer, 20, strSupportFile)
    strBuffer = Left(strBuffer, InStr(strBuffer, Chr(0)) - 1)
    mvarDetailedAuditingRequest = CBool(strBuffer)
    '
    strBuffer = Space(20)
    retval = GetPrivateProfileString("Support", "HTMLEmailSupportRequestDate", "", strBuffer, 20, strSupportFile)
    strBuffer = Left(strBuffer, InStr(strBuffer, Chr(0)) - 1)
    If IsDate(strBuffer) Then
        mvarHTMLEmailSupportRequestDate = CDate(strBuffer)
    Else
        mvarHTMLEmailSupportRequestDate = 0
    End If
    '
    '   Diagnostics Request Date
    '
    strBuffer = Space(20)
    retval = GetPrivateProfileString("Support", "DiagnosticsRequestDate", "", strBuffer, 20, strSupportFile)
    strBuffer = Left(strBuffer, InStr(strBuffer, Chr(0)) - 1)
    If IsDate(strBuffer) Then
        mvarDiagnosticsRequestDate = CDate(strBuffer)
    Else
        mvarDiagnosticsRequestDate = 0
    End If
    '
    '   Diagnostic SQL
    '
    strBuffer = Space(1000)
    retval = GetPrivateProfileString("Support", "DiagnosticSQL", "", strBuffer, 1000, strSupportFile)
    '
    mvarDiagnosticSQL = Left(strBuffer, InStr(strBuffer, Chr(0)) - 1)
'
'********* Code Ends Here **************
'
   Exit Sub
'
SupportUserFlags_Error:
    ErrorRaise "SystemConfig.SupportUserFlags"
End Sub

Public Sub UpdateServerLocation()
'***************************************
' Module/Form Name   : SystemConfig
'
' Procedure Name     : UpdateServerLocation
'
' Purpose            :
'
' Date Created       : 10/03/2025
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo UpdateServerLocation_Error
'
'******** Code Starts Here *************
'
    Dim retval          As Long
    
    retval = WritePrivateProfileString("General", "ServerLocation", CStr(mvarServerLocation), mstrIniFile)
'
'********* Code Ends Here **************
'
    Exit Sub
    '
UpdateServerLocation_Error:
    ErrorRaise "SystemConfig.UpdateServerLocation"
End Sub

Public Sub RefreshServerLocation()
'***************************************
' Module/Form Name   : SystemConfig
'
' Procedure Name     : RefreshServerLocation
'
' Purpose            :
'
' Date Created       : 11/03/2025
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo RefreshServerLocation_Error
'
'******** Code Starts Here *************
'
    Dim retval          As Long
    Dim strBuffer           As String
    
    strBuffer = Space(200)
    retval = GetPrivateProfileString("General", "ServerLocation", "", strBuffer, 200, mstrIniFile)
    If Len(Trim(strBuffer)) <= 1 Then
        mvarServerLocation = App.Path
        mstrServerIniFile = App.Path & "\IW.ini"
    Else
        mvarServerLocation = Left(strBuffer, InStr(strBuffer, Chr(0)) - 1)
        mstrServerIniFile = mvarServerLocation & "\IW.ini"
    End If

'
'********* Code Ends Here **************
'
    Exit Sub
    '
RefreshServerLocation_Error:
    ErrorRaise "SystemConfig.UpdateServerLocation"
End Sub

Public Sub update()
'***************************************
' Module/Form Name   : SystemConfig
'
' Procedure Name     : Update
'
' Purpose            :
'
' Date Created       : 02/06/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 02/06/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo update_Error
'
'******** Code Starts Here *************
'
    Dim strSQL          As String
    Dim rstSysConfig    As ADODB.Recordset
    Dim strBuffer       As String
    Dim retval          As Long
    
    strSQL = "SELECT * FROM SYSTEMCONFIG"
    Set rstSysConfig = New ADODB.Recordset
    rstSysConfig.Open strSQL, gdbADO, adOpenKeyset, adLockPessimistic
    With rstSysConfig
        !DateLogFilePosted = mvarDateLogFilePosted
        !LogFilePostingFrequency = mvarLogFilePostingFrequency
        !SupportEmail = mvarSupportEmail
        !WebSearchTestEmail = mvarWebSearchTestEmail
        !PostedHTMLEmail = mvarPostedHTMLEmail
        !ConfirmationHTMLEmail = mvarConfirmationHTMLEmail
        !CancellationHTMLEmail = mvarCancellationHTMLEmail
        !BundledHTMLEmail = mvarBundledHTMLEmail
        .update
    End With
    rstSysConfig.Close
    Set rstSysConfig = Nothing
    '
    '   Update registry Settings.
    '
    SaveSetting "Image Where", "Options", "DateHTMLEmailPosted", Format(mvarDateHTMLEmailPosted, "dd/mm/yyyy hh:nn:ss")
    SaveSetting "Image Where", "Options", "DateHTMLEmailUpgraded", Format(mvarDateHTMLEmailUpgraded, "dd/mm/yyyy hh:nn:ss")
    SaveSetting "Image Where", "Options", "DateDiagnosticsPosted", Format(mvarDateDiagnosticsPosted, "dd/mm/yyyy hh:nn:ss")
    SaveSetting "Image Where", "Upgrade", "Checkinterval", CStr(mvarUpgradeCheckFrequency)
    SaveSetting "Image Where", "Options", "DateDetailedAuditingSet", Format(mvarDateDetailedAuditingSet, "dd/mm/yyyy hh:nn:ss")
    SaveSetting "Image Where", "Options", "DatePreviousPostWarning", Format(mvarDatePreviousPostWarning, "dd/mm/yyyy hh:nn:ss")
    '
    retval = WritePrivateProfileString("Support", "DetailedAuditing", CStr(mvarDetailedAuditing), mstrServerIniFile)
    retval = WritePrivateProfileString("FTP", "Server", mvarFTPServer, mstrServerIniFile)
    retval = WritePrivateProfileString("FTP", "User", mvarFTPUser, mstrServerIniFile)
    retval = WritePrivateProfileString("FTP", "Password", mvarFTPPassword, mstrServerIniFile)
    retval = WritePrivateProfileString("FTP", "PostWebAddress", mvarPostWebAddress, mstrServerIniFile)
    retval = WritePrivateProfileString("FTP", "WebSearchesEmailTo", mvarWebSearchesEmailTo, mstrServerIniFile)
    retval = WritePrivateProfileString("FTP", "WebSearchesEmailFrom", mvarWebSearchesEmailFrom, mstrServerIniFile)
    retval = WritePrivateProfileString("General", "FuzzyKeywordSearch", CStr(mvarFuzzyKeywordSearch), mstrServerIniFile)
    retval = WritePrivateProfileString("General", "BasicImageWhere", CStr(mvarBasicImageWhere), mstrServerIniFile)
    SaveSetting "Image Where", "Options", "MouseWheel", CStr(mvarMouseWheel)
    SaveSetting "Image Where", "Options", "TooltipDelay", CStr(mvarTooltipDelay)

    mvarDirty = False
'
'********* Code Ends Here **************
'
    Exit Sub
    '
update_Error:
    ErrorRaise "SystemConfig.Update"
End Sub

Public Property Get WebNonFeedbackAlertPeriod() As Integer
    WebNonFeedbackAlertPeriod = 5
End Property

Public Property Get WebUseByPeriod() As Integer
    WebUseByPeriod = 15
End Property

Public Property Get WebUseByPeriodDigital() As Integer
    WebUseByPeriodDigital = 365
End Property

Public Property Get UpgradeCheckFrequency() As Integer
    UpgradeCheckFrequency = mvarUpgradeCheckFrequency
End Property

Public Property Let UpgradeCheckFrequency(vData As Integer)
    mvarUpgradeCheckFrequency = vData
    mvarDirty = True
End Property

Public Property Get ActiveConnection() As Boolean
   
  'function checks registry for
  'an active connection
   
   Dim hKey As Long
   Dim lpData As Long
   Dim nSizeData As Long
   
   Const sSubKey = "System\CurrentControlSet\Services\RemoteAccess"
   Const sKeyValue = "Remote Connection"

   If RegOpenKey(HKEY_LOCAL_MACHINE, _
                 sSubKey, _
                 hKey) = ERROR_SUCCESS Then
   
      lpData = 0&
      nSizeData = Len(lpData)

      If RegQueryValueEx(hKey, _
                         sKeyValue, _
                         0&, _
                         0&, _
                         lpData, _
                         nSizeData) = ERROR_SUCCESS Then

         ActiveConnection = lpData <> 0

      End If

      Call RegCloseKey(hKey)
   
   End If

End Property

Public Property Get ListViewColumnHeaderHeight() As Long
    ListViewColumnHeaderHeight = 355
End Property

Private Sub SetUpHTMLImages()
    SetUpEmailHTMLImages mvarPostedHTMLEmail, mvarPostedImages
    SetUpEmailHTMLImages mvarConfirmationHTMLEmail, mvarConfirmationImages
    SetUpEmailHTMLImages mvarCancellationHTMLEmail, mvarCancellationImages
End Sub

Private Sub SetUpEmailHTMLImages(ByRef strEmailBody As String, _
                                 ByRef colImages As Collection)
    '***************************************
    ' Module/Form Name   : SystemConfig
    '
    ' Procedure Name     : SetUpEmailHTMLImages
    '
    ' Purpose            :
    '
    ' Date Created       : 28/05/2002
    '
    ' Author             : GARETH SAUNDERS
    '
    ' Parameters         : strEmailBody - String
    '                    : colImages - Collection
    '
    ' Amendment History  : Date       Author    Description
    '                    : --------------------------------
    '                    : 28/05/2002 GARETH SAUNDERS
    '
    '***************************************
    '
On Error GoTo SetUpEmailHTMLImages_Error
    '
    '******** Code Starts Here *************
    '
    Dim strImage As String
    Dim intStart As Integer
    Dim intEnd As Integer

    '
    '   Create the collection of images to be embedded into the HTML Email.
    '
Set colImages = Nothing
Set colImages = New Collection
    '
intStart = InStr(1, strEmailBody, "<IWIMAGE ")
While Not intStart = 0
  intEnd = InStr(intStart, strEmailBody, ">")
  '
  '   We don't want extra images so don't check for duplicate key errors.
  '
  strImage = Mid(strEmailBody, intStart + 9, intEnd - intStart - 9)
  On Error Resume Next
  colImages.Add strImage, strImage
  On Error GoTo SetUpEmailHTMLImages_Error
  '
  '   Replace the IWIMAGE tag.
  '
  strEmailBody = Left(strEmailBody, intStart) & _
                 "IMG align=baseline border = 0 hspace = 0 src=cid:" & _
                 Mid(strEmailBody, intStart + 9)
  '
  '   Search for the next tag.
  '
  intStart = InStr(intEnd, strEmailBody, "<IWIMAGE ")
Wend
    '
    '********* Code Ends Here **************
    '
Exit Sub
    '
SetUpEmailHTMLImages_Error:
    ErrorRaise "SystemConfig.SetUpEmailHTMLImages"
End Sub

Public Property Get LogFileShouldBePosted() As Boolean
    LogFileShouldBePosted = False
    If mvarSupportUser <> "True" Then Exit Sub
    
    If mvarLogFilePostingFrequency = 0 Then Exit Sub
  
    If Trim(mvarSupportEmail) = "" Then Exit Sub
    
    If DateDiff("d", mvarDateLogFilePosted, Now()) < mvarLogFilePostingFrequency Then Exit Sub
    
    If DateAdd("d", 1, mvarDatePreviousPostWarning) > Now() Then Exit Sub
    
    LogFileShouldBePosted = True
    
End Property

Public Property Get Locations() As Locations
    If mvarLocations Is Nothing Then
        Set mvarLocations = New Locations
        mvarLocations.Refresh
    End If

    Set Locations = mvarLocations
End Property

Public Property Set Locations(vData As Locations)
    Set mvarLocations = vData
    mvarDirty = True
End Property

Public Function PhotoTypeDescription(ByVal pstrPhotoType As String)
    Select Case pstrPhotoType
    Case Is = "D"
        PhotoTypeDescription = "Digital"
    Case Is = "T"
        PhotoTypeDescription = "Transpncy"
    End Select
End Function

Public Property Get DetailedAuditing() As Boolean
    DetailedAuditing = mvarDetailedAuditing
End Property

Public Property Let DetailedAuditing(vData As Boolean)
    mvarDetailedAuditing = vData
    mvarDirty = True
End Property

Public Property Get DateDetailedAuditingSet() As Date
    DateDetailedAuditingSet = mvarDateDetailedAuditingSet
End Property

Public Property Let DateDetailedAuditingSet(vData As Date)
    mvarDateDetailedAuditingSet = vData
    mvarDirty = True
End Property

Public Property Get DateDetailedAuditingRequested() As Date
    DateDetailedAuditingRequested = mvarDateDetailedAuditingRequested
End Property

Public Property Let DateDetailedAuditingRequested(vData As Date)
    mvarDateDetailedAuditingRequested = vData
    mvarDirty = True
End Property

Public Property Get DetailedAuditingRequest() As String
    DetailedAuditingRequest = mvarDetailedAuditingRequest
End Property

Public Property Let DetailedAuditingRequest(vData As String)
    mvarDetailedAuditingRequest = vData
    mvarDirty = True
End Property

Public Property Get FTPServer() As String
    FTPServer = mvarFTPServer
End Property

Public Property Let FTPServer(vData As String)
    mvarFTPServer = vData
    mvarDirty = True
End Property

Public Property Get FTPUser() As String
    FTPUser = mvarFTPUser
End Property

Public Property Let FTPUser(vData As String)
    mvarFTPUser = vData
    mvarDirty = True
End Property

Public Property Get FTPPassword() As String
    FTPPassword = mvarFTPPassword
End Property

Public Property Let FTPPassword(vData As String)
    mvarFTPPassword = vData
    mvarDirty = True
End Property

Public Property Get PostWebAddress() As String
    PostWebAddress = mvarPostWebAddress
End Property

Public Property Let PostWebAddress(vData As String)
    mvarPostWebAddress = vData
    mvarDirty = True
End Property

Public Property Get WebSearchesEmailTo() As String
    WebSearchesEmailTo = IIf(mvarWebSearchesEmailTo = "", "pictures@clivenichols.co.uk", mvarWebSearchesEmailTo)
End Property

Public Property Let WebSearchesEmailTo(vData As String)
    mvarWebSearchesEmailTo = vData
    mvarDirty = True
End Property

Public Property Get WebSearchesEmailFrom() As String
    WebSearchesEmailFrom = IIf(mvarWebSearchesEmailFrom = "", "websearches@clivenichols.co.uk", mvarWebSearchesEmailFrom)
End Property

Public Property Let WebSearchesEmailFrom(vData As String)
    mvarWebSearchesEmailFrom = vData
    mvarDirty = True
End Property

