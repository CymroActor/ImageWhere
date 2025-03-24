VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdi_npls 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Image Where"
   ClientHeight    =   4080
   ClientLeft      =   1995
   ClientTop       =   2145
   ClientWidth     =   6690
   Icon            =   "NPLS.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar tob_npls 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6690
      _ExtentX        =   11800
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Customer"
            Object.ToolTipText     =   "Customer Maintenance"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Photograph"
            Object.ToolTipText     =   "Photograph Maintenance"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Job"
            Object.ToolTipText     =   "Job Maintenance"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delivery Note"
            Object.ToolTipText     =   "Delivery Note Maintenance"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Photo Labels"
            Object.ToolTipText     =   "Photograph Label Print"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "KeywordClustering"
            Object.ToolTipText     =   "Keyword Clustering"
         EndProperty
      EndProperty
   End
   Begin VB.Timer Tim_number_of_photos 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   2880
      Top             =   1560
   End
   Begin MSComctlLib.StatusBar stb_npls 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   3825
      Width           =   6690
      _ExtentX        =   11800
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            AutoSize        =   2
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            AutoSize        =   2
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            AutoSize        =   2
            Enabled         =   0   'False
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "11/03/2025"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList iml_npls 
      Left            =   1440
      Top             =   1020
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483636
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NPLS.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NPLS.frx":0624
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NPLS.frx":093E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NPLS.frx":0C58
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NPLS.frx":0F72
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NPLS.frx":128C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      NegotiatePosition=   1  'Left
      Begin VB.Menu print_photo_labels 
         Caption         =   "&Print photograph labels"
      End
      Begin VB.Menu File_hyphen_1 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu view 
      Caption         =   "&View"
      Begin VB.Menu photograph_find_maint 
         Caption         =   "Photograph &Find"
      End
      Begin VB.Menu hyphen 
         Caption         =   "-"
      End
      Begin VB.Menu tool_bar 
         Caption         =   "&Tool Bar"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu Maintenance 
      Caption         =   "&Maintenance"
      Begin VB.Menu Customer_maint 
         Caption         =   "&Customer"
      End
      Begin VB.Menu Photograph_maint 
         Caption         =   "&Photograph"
      End
      Begin VB.Menu job_maint 
         Caption         =   "&Job"
      End
      Begin VB.Menu delivery_note_maint 
         Caption         =   "&Delivery Note"
      End
   End
   Begin VB.Menu mnuJobs 
      Caption         =   "&Jobs"
      Visible         =   0   'False
      Begin VB.Menu mnuJobsNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuJobsEdit 
         Caption         =   "&Edit"
      End
      Begin VB.Menu mnuJobsDelete 
         Caption         =   "&Delete"
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "&Search"
      Begin VB.Menu mnuSearchEditPhoto 
         Caption         =   "Edit &photograph"
      End
      Begin VB.Menu mnuRemoveItem 
         Caption         =   "&Remove item(s) from search"
      End
   End
   Begin VB.Menu mnuReturn 
      Caption         =   "&Return"
      Begin VB.Menu mnuReturnImage 
         Caption         =   "Return &Image(s)"
      End
      Begin VB.Menu mnuReturnEditPhoto 
         Caption         =   "&Edit Photograph"
      End
   End
   Begin VB.Menu mnuDeliveryNoteMaint 
      Caption         =   "Delivery Note Maintenance"
      Visible         =   0   'False
      Begin VB.Menu mnuEditJobReference 
         Caption         =   "Edit Job Reference"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
      Visible         =   0   'False
      Begin VB.Menu mnuReportAgedPictures 
         Caption         =   "&Aged Pictures"
      End
   End
   Begin VB.Menu mnuChasers 
      Caption         =   "Chasers"
      Visible         =   0   'False
      Begin VB.Menu mnuChasersAdd 
         Caption         =   "&Add"
      End
      Begin VB.Menu mnuChasersPrint 
         Caption         =   "&Print"
      End
   End
   Begin VB.Menu mnuPhotographEditPopups 
      Caption         =   "Photograph Edit Popups"
      Visible         =   0   'False
      Begin VB.Menu mnuPhotographEditPopupsNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuPhotographEditPopupsEdit 
         Caption         =   "&Edit"
      End
      Begin VB.Menu mnuPhotographEditPopupsDelete 
         Caption         =   "&Delete"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsKeywordMaintenance 
         Caption         =   "&Keyword Maintenance"
         Begin VB.Menu mnuToolsApplyKeywordsToPhotographs 
            Caption         =   "Apply Keywords To &Photographs"
         End
         Begin VB.Menu mnuKeywordClustering 
            Caption         =   "Keyword &Clustering"
         End
         Begin VB.Menu mnuToolsKeywordExclusions 
            Caption         =   "Keyword &Exclusions"
         End
      End
      Begin VB.Menu mnuToolhyphen 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsOptions 
         Caption         =   "&Options"
      End
   End
   Begin VB.Menu Window_list 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu Cascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu Tile 
         Caption         =   "&Tile"
      End
      Begin VB.Menu ArrangeIcons 
         Caption         =   "&Arrange Icons"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnuHelpSearch 
         Caption         =   "&Search For Help On"
      End
      Begin VB.Menu mnuHelpSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu About 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "mdi_npls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
'   Used for determining whether Image Where has loaded.
'
Private Image_Where_loaded As Boolean
'
'   Used to determine which Search Form last requested a Pop Up Menu.
'
Public fSearch As frm_search
Public fPopUp As Form

Private Sub About_Click()
'***************************************
' Module/Form Name   : mdi_npls
'
' Procedure Name     : About_Click
'
' Purpose            :
'
' Date Created       : 31/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo About_Click_Error
'
'******** Code Starts Here *************
'
    AbSplash.About Me
'
'********* Code Ends Here **************
'
   Exit Sub
'
About_Click_Error:
    DisplayError , "mdi_npls.About_Click", vbExclamation
End Sub

Private Sub ArrangeIcons_Click()
    mdi_npls.Arrange vbArrangeIcons
End Sub

Private Sub Cascade_Click()
    mdi_npls.Arrange vbCascade
End Sub

Private Sub Customer_maint_Click()
'***************************************
' Module/Form Name   : mdi_npls
'
' Procedure Name     : Customer_maint_Click
'
' Purpose            :
'
' Date Created       : 31/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo Customer_maint_Click_Error
'
'******** Code Starts Here *************
'
    Load frm_cust_maint
'
'********* Code Ends Here **************
'
   Exit Sub
'
Customer_maint_Click_Error:
    DisplayError , "mdi_npls.Customer_maint_Click", vbExclamation
End Sub

Private Sub delivery_note_maint_Click()
'***************************************
' Module/Form Name   : mdi_npls
'
' Procedure Name     : delivery_note_maint_Click
'
' Purpose            :
'
' Date Created       : 31/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo delivery_note_maint_Click_Error
'
'******** Code Starts Here *************
'
    frm_delivery_note.load_delivery_notes 0, "0", 0
'
'********* Code Ends Here **************
'
   Exit Sub
'
delivery_note_maint_Click_Error:
    DisplayError , "mdi_npls.delivery_note_maint_Click", vbExclamation
End Sub

Public Sub display_last_photo_number_used()
'***************************************
' Module/Form Name   : mdi_npls
'
' Procedure Name     : display_last_photo_number_used
'
' Purpose            :
'
' Date Created       : 31/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo display_last_photo_number_used_Error
'
'******** Code Starts Here *************
'
    Dim pnlX As Panel
    Dim lngLastNumberUsed As Long
    
    Set pnlX = stb_npls.Panels.Item(1)
    On Error Resume Next
    lngLastNumberUsed = photograph.get_last_number_used
    goSystemConfig.LastPhotoNumberUsed = lngLastNumberUsed
    If Err.Number = 0 Then
        pnlX.Text = "Last photograph number:  " & CStr(photograph.get_last_number_used)
    End If
'
'********* Code Ends Here **************
'
   Exit Sub
'
display_last_photo_number_used_Error:
    ErrorRaise "mdi_npls.display_last_photo_number_used"
End Sub

Private Sub exit_Click()
    Unload Me
End Sub

Private Sub job_maint_Click()
'***************************************
' Module/Form Name   : mdi_npls
'
' Procedure Name     : job_maint_Click
'
' Purpose            :
'
' Date Created       : 31/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo job_maint_Click_Error
'
'******** Code Starts Here *************
'
    Screen.MousePointer = ccHourglass
    Load frm_job_maint
    Screen.MousePointer = ccDefault
'
'********* Code Ends Here **************
'
   Exit Sub
'
job_maint_Click_Error:
    DisplayError , "mdi_npls.job_maint_Click", vbExclamation
End Sub

Private Sub MDIForm_Load()
'***************************************
' Module/Form Name   : mdi_npls
'
' Procedure Name     : MDIForm_Load
'
' Purpose            :
'
' Date Created       : 02/06/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo MDIForm_Load_Error
'
'******** Code Starts Here *************
'
    Dim strErrorNo          As String
    Dim objUpgradeChecker   As Object
    Dim oFSO                As Scripting.FileSystemObject
    Dim fDB                 As File
    Dim intDBSizeMB         As Integer
    
    If IsInIDE Then
        gblnInhibitSubClassing = True
    Else
        gblnInhibitSubClassing = False
    End If
    '
    '   Initialise the lower level error flag.
    '
    gLowerLevelError = False
    '
    AbSplash.SplashOn Me, 5000
    AbSplash.SplashOff
    Tim_number_of_photos.Enabled = True
    '
    '   POPUP Menus are invisible to begin with.
    '
    mnuSearch.Visible = False
    mnuReturn.Visible = False
    '
    '   Initialise system variables.
    '
    Screen.MousePointer = ccHourglass
    Image_Where_loaded = False
    gblnResizeMaxedWindows = True
    gintTransCount = 0
    
    Action = "Initialising Variables"
    scroll_bar_width = 364
    Set tob_npls.ImageList = iml_npls
    tob_npls.Buttons(1).Image = 1
    tob_npls.Buttons(2).Image = 2
    tob_npls.Buttons(3).Image = 3
    tob_npls.Buttons(4).Image = 4
    tob_npls.Buttons(5).Image = 5
    tob_npls.Buttons(6).Image = 6
        
    tob_npls.Refresh
    DoEvents
    '
    '   Create System Configuration Object and retrieve the Server Location.
    '
    goSystemConfig.RefreshServerLocation
    If goSystemConfig.ServerLocation = "" Then
        glo_dbname = App.Path & "\Database\Library7.mdb"
    Else
        glo_dbname = goSystemConfig.ServerLocation & "\Database\Library7.mdb"
    End If
    '
    '   Does the Database Exist?
    '
    Set oFSO = New Scripting.FileSystemObject

    If Not oFSO.FileExists(glo_dbname) Then
        MsgBox "Database not found. Please locate the Server location where the current Dtatbase folder resides.", vbExclamation, "Image Where!"
        Dim spath               As String
        Dim strPathReturned     As String
    
        spath = ""
        strPathReturned = BrowseForFolderByPath(Me, spath)
        '
        If strPathReturned <> "" Then
            goSystemConfig.ServerLocation = strPathReturned
            goSystemConfig.UpdateServerLocation
            glo_dbname = oFSO.BuildPath(strPathReturned, "Database\Library7.mdb")
        Else
            Unload Me
            Exit Sub
        End If
    End If
    Set oFSO = Nothing
    '
    '   Set up the path variables.
    '
    'gstrAppPath = GetSetting(CONST_APPLICATION, "Locations", "Application", App.Path)
    gstrAppPath = goSystemConfig.ServerLocation
    gstrUpgradePath = goSystemConfig.ServerLocation & "\Upgrade"
    '
    '   Open the Database
    '
    Action = "Opening Database"
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(glo_dbname)
    DBEngine.SetOption dbFlushTransactionTimeout, 1
    '
    '   Create ADO database Object.
    '
    Action = ""
    Set gdbADO = New ADODB.Connection
    'gADOConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & glo_dbname & ";Persist Security Info=False"
    gADOConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & glo_dbname
    gdbADO.Open gADOConnection
    transaction_started = False
    '
    '   Reresh the System Config with the now available DB info.
    '
    goSystemConfig.Refresh
    '
    '   Set the Maxed Windows collection.
    '
    Set gcolMaxedWindows = New Collection
    '
    '   Create Log object.
    '
''    Set goLog = New Log
    '
    '   Record the application starting.
    '
    goLog.WriteLog "Application Starting", ""
    '
    '   Basic Image Where?
    '
    If goSystemConfig.BasicImageWhere Then
        tob_npls.Buttons(5).Visible = False
        tob_npls.Buttons(6).Visible = False
    End If
    '
    '   Create Company Information Object.
    '
    Set goCompanyInfo = New CompanyInfo
    goCompanyInfo.Refresh
    '
    Me.Caption = Me.Caption & " for " & goCompanyInfo.CompanyName
    '
    '   Is this PC set up to for upgrading IW.
    '
    On Error Resume Next
    Set objUpgradeChecker = CreateObject("IWUpgradeChecker2.MainControl")
    If Err.Number = 0 Then
        '
        '   Check if there is an upgrade.
        '
        On Error GoTo MDIForm_Load_Error
        If UpgradeExists Then
            On Error Resume Next
            Unload Me
            Set mdi_npls = Nothing
            End
            Exit Sub
        Else
            objUpgradeChecker.Startup
        End If
    Else
        On Error GoTo MDIForm_Load_Error
    End If
    Set objUpgradeChecker = Nothing
    '
    '   Display the last photograph number used. Subsequently a timer control will
    '   continually display this.
    '
    display_last_photo_number_used
    '
    '   Record application has having loaded.
    '
    Image_Where_loaded = True
    Screen.MousePointer = ccDefault
    Close #1
    '
    '   Initialise settings.
    '
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    tool_bar.Checked = GetSetting(App.Title, "Settings", "Toolbar", True)
    tob_npls.Visible = tool_bar.Checked
    '
    Resize.Hook Me.hWnd, 700, 500
    '
    Exit Sub

MDIForm_Load_Error:
      Screen.MousePointer = ccDefault
      Select Case Err.Number
          Case Is = 3044
              MsgBox "Database not found. Possible causes are:" + vbCr + _
                    "    1. The database server is not running." + vbCr + _
                    "    2. The network is down.", vbCritical, "Image Where!"
          Case Else
              DisplayError , "mdi_npls.MDIForm_Load", vbCritical
      End Select
      Unload Me
      Set mdi_npls = Nothing
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'***************************************
' Module/Form Name   : mdi_npls
'
' Procedure Name     : MDIForm_QueryUnload
'
' Purpose            :
'
' Date Created       : 31/12/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : Cancel - Integer
'                    : UnloadMode - Integer
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo MDIForm_QueryUnload_Error
'
'******** Code Starts Here *************
'
    Dim msg ' Declare variable.
    Dim frm As Form
    '
    '   If application has failed even to load, don't ask the question?
    '
    If Not Image_Where_loaded Then
        Exit Sub
    End If
    '
    ' Set the message text.
    '
    msg = "Do you really want to exit the application?"
    '
    ' If user clicks the No button, stop QueryUnload.
    '
    If MsgBox(msg, 36, App.Title) = vbNo Then
        Cancel = True
        Exit Sub
    End If

    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If

    SaveSetting App.Title, "Settings", "Toolbar", tool_bar.Checked
    db.Close
    Set db = Nothing
    ws.Close
    Set ws = Nothing
    For Each frm In Forms
        If frm.Name <> "mdi_npls" Then
            Unload frm
            Set frm = Nothing
        End If
    Next frm
    Unload mdi_npls
    Set mdi_npls = Nothing
'
'********* Code Ends Here **************
'
   Exit Sub
'
MDIForm_QueryUnload_Error:
    DisplayError , "mdi_npls.MDIForm_QueryUnload", vbExclamation
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    '
    '   Record the application starting.
    '
    goLog.WriteLog "Application Finishing", ""
    '
    Set goSystemConfig = Nothing
    Set goCompanyInfo = Nothing
    Set goLog = Nothing
    Resize.UnHook Me.hWnd
End Sub

Private Sub mnuChasersAdd_Click()
'***************************************
' Module/Form Name   : mdi_npls
'
' Procedure Name     : mnuChasersAdd_Click
'
' Purpose            :
'
' Date Created       : 31/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo mnuChasersAdd_Click_Error
'
'******** Code Starts Here *************
'
          fPopUp.AddChaser
'
'********* Code Ends Here **************
'
   Exit Sub
'
mnuChasersAdd_Click_Error:
    DisplayError , "mdi_npls.mnuChasersAdd_Click", vbExclamation
End Sub

Private Sub mnuChasersPrint_Click()
'***************************************
' Module/Form Name   : mdi_npls
'
' Procedure Name     : mnuChasersPrint_Click
'
' Purpose            :
'
' Date Created       : 31/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo mnuChasersPrint_Click_Error
'
'******** Code Starts Here *************
'
          fPopUp.PrintChaser
'
'********* Code Ends Here **************
'
   Exit Sub
'
mnuChasersPrint_Click_Error:
    DisplayError , "mdi_npls.mnuChasersPrint_Click", vbExclamation
End Sub

Private Sub mnuHelpContents_Click()
'***************************************
' Module/Form Name   : mdi_npls
'
' Procedure Name     : mnuHelpContents_Click
'
' Purpose            :
'
' Date Created       : 31/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo mnuHelpContents_Click_Error
'
'******** Code Starts Here *************
'
    Dim nRet As Integer
    Dim strKey As String * 255

    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    
        
    
    
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        strKey = ""
        nRet = OSWinHelp(Me.hWnd, App.HelpFile, 3, strKey)
        If Err Then
            MsgBox Err.Description
        End If
    End If
'
'********* Code Ends Here **************
'
   Exit Sub
'
mnuHelpContents_Click_Error:
    DisplayError , "mdi_npls.mnuHelpContents_Click", vbExclamation
End Sub

Private Sub mnuHelpSearch_Click()
'***************************************
' Module/Form Name   : mdi_npls
'
' Procedure Name     : mnuHelpSearch_Click
'
' Purpose            :
'
' Date Created       : 31/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo mnuHelpSearch_Click_Error
'
'******** Code Starts Here *************
'

    Dim nRet As Integer
    Dim strKey As String * 255

    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        strKey = ""
        nRet = OSWinHelp(Me.hWnd, App.HelpFile, 261, strKey)
        If Err Then
            MsgBox Err.Description
        End If
    End If
'
'********* Code Ends Here **************
'
   Exit Sub
'
mnuHelpSearch_Click_Error:
    DisplayError , "mdi_npls.mnuHelpSearch_Click", vbExclamation
End Sub

Private Sub mnuJobsDelete_Click()
'***************************************
' Module/Form Name   : mdi_npls
'
' Procedure Name     : mnuJobsDelete_Click
'
' Purpose            :
'
' Date Created       : 14/04/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 14/04/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo mnuJobsDelete_Click_Error
'
'******** Code Starts Here *************
'
    fPopUp.DeleteJob
'
'********* Code Ends Here **************
'
    Exit Sub
    '
mnuJobsDelete_Click_Error:
    DisplayError , "mdi_npls.mnuJobsDelete_Click", vbExclamation
End Sub

Private Sub mnuJobsEdit_Click()
'***************************************
' Module/Form Name   : mdi_npls
'
' Procedure Name     : mnuJobsEdit_Click
'
' Purpose            :
'
' Date Created       : 31/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo mnuJobsEdit_Click_Error
'
'******** Code Starts Here *************
'
    fPopUp.EditJob
'
'********* Code Ends Here **************
'
   Exit Sub
'
mnuJobsEdit_Click_Error:
    DisplayError , "mdi_npls.mnuJobsEdit_Click", vbExclamation
End Sub

Private Sub mnuJobsNew_Click()
'***************************************
' Module/Form Name   : mdi_npls
'
' Procedure Name     : mnuJobsNew_Click
'
' Purpose            :
'
' Date Created       : 31/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo mnuJobsNew_Click_Error
'
'******** Code Starts Here *************
'
    fPopUp.NewJob
'
'********* Code Ends Here **************
'
   Exit Sub
'
mnuJobsNew_Click_Error:
    DisplayError , "mdi_npls.mnuJobsNew_Click", vbExclamation
End Sub

Private Sub mnuKeywordClustering_Click()
'***************************************
' Module/Form Name   : mdi_npls
'
' Procedure Name     : mnuKeywordClustering_Click
'
' Purpose            :
'
' Date Created       : 25/07/2005
'
' Author             : Gareth Saunders
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                      25/07/2005 Gareth Saunders
'
'***************************************
'
    On Error GoTo ErrorHandler:
'
'******** Code Starts Here *************
'
    Call KeywordClustering
'
'********* Code Ends Here **************
'
    On Error GoTo 0
    Exit Sub
'
ErrorHandler:
    DisplayError , "mdi_npls.mnuKeywordClustering_Click", vbExclamation
End Sub

Private Sub mnuPhotographEditPopupsDelete_Click()
    fPopUp.cmdDeleteKeyword.Value = True
End Sub

Private Sub mnuPhotographEditPopupsEdit_Click()
    fPopUp.cmdEditKeyword.Value = True
End Sub

Private Sub mnuPhotographEditPopupsNew_Click()
    fPopUp.cmdNewKeyword.Value = True
End Sub

Private Sub mnuRemoveItem_Click()
'***************************************
' Module/Form Name   : mdi_npls
'
' Procedure Name     : mnuRemoveItem_Click
'
' Purpose            :
'
' Date Created       : 31/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo mnuRemoveItem_Click_Error
'
'******** Code Starts Here *************
'
    fSearch.cmdDelete.Value = True
'
'********* Code Ends Here **************
'
   Exit Sub
'
mnuRemoveItem_Click_Error:
    DisplayError , "mdi_npls.mnuRemoveItem_Click", vbExclamation
End Sub

Private Sub mnuReportAgedPictures_Click()
'***************************************
' Module/Form Name   : mdi_npls
'
' Procedure Name     : mnuReportAgedPictures_Click
'
' Purpose            :
'
' Date Created       : 31/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo mnuReportAgedPictures_Click_Error
'
'******** Code Starts Here *************
'
    frmAgedPictures.Show 1
'
'********* Code Ends Here **************
'
   Exit Sub
'
mnuReportAgedPictures_Click_Error:
    DisplayError , "mdi_npls.mnuReportAgedPictures_Click", vbExclamation
End Sub

Private Sub mnuReturnEditPhoto_Click()
'***************************************
' Module/Form Name   : mdi_npls
'
' Procedure Name     : mnuReturnEditPhoto_Click
'
' Purpose            :
'
' Date Created       : 31/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo mnuReturnEditPhoto_Click_Error
'
'******** Code Starts Here *************
'
    frm_delivery_note_return_photos.edit_photo
'
'********* Code Ends Here **************
'
   Exit Sub
'
mnuReturnEditPhoto_Click_Error:
    DisplayError , "mdi_npls.mnuReturnEditPhoto_Click", vbExclamation
End Sub

Private Sub mnuReturnImage_Click()
'***************************************
' Module/Form Name   : mdi_npls
'
' Procedure Name     : mnuReturnImage_Click
'
' Purpose            :
'
' Date Created       : 31/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo mnuReturnImage_Click_Error
'
'******** Code Starts Here *************
'
    frm_delivery_note_return_photos.cmdReturn.Value = True
'
'********* Code Ends Here **************
'
   Exit Sub
'
mnuReturnImage_Click_Error:
    DisplayError , "mdi_npls.mnuReturnImage_Click", vbExclamation
End Sub

Private Sub mnuSearchEditPhoto_Click()
'***************************************
' Module/Form Name   : mdi_npls
'
' Procedure Name     : mnuSearchEditPhoto_Click
'
' Purpose            :
'
' Date Created       : 31/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo mnuSearchEditPhoto_Click_Error
'
'******** Code Starts Here *************
'
    fSearch.edit_photo
'
'********* Code Ends Here **************
'
   Exit Sub
'
mnuSearchEditPhoto_Click_Error:
    DisplayError , "mdi_npls.mnuSearchEditPhoto_Click", vbExclamation
End Sub

Private Sub mnuToolsApplyKeywordsToPhotographs_Click()
    frmApplyKeywordsToPhotographs.Show vbModal
End Sub

Private Sub mnuToolsKeywordExclusions_Click()
    Dim fMaintainKeywordExclusions  As frmKeywordExclusions
    
    Set fMaintainKeywordExclusions = New frmKeywordExclusions
    fMaintainKeywordExclusions.Display
    Set fMaintainKeywordExclusions = Nothing
End Sub

Private Sub mnuToolsOptions_Click()
'***************************************
' Module/Form Name   : mdi_npls
'
' Procedure Name     : mnuToolsOptions_Click
'
' Purpose            :
'
' Date Created       : 31/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo mnuToolsOptions_Click_Error
'
'******** Code Starts Here *************
'
    frmToolsOptions.Show 1
'
'********* Code Ends Here **************
'
   Exit Sub
'
mnuToolsOptions_Click_Error:
    DisplayError , "mdi_npls.mnuToolsOptions_Click", vbExclamation
End Sub

Private Sub photograph_find_maint_Click()
'***************************************
' Module/Form Name   : mdi_npls
'
' Procedure Name     : photograph_find_maint_Click
'
' Purpose            :
'
' Date Created       : 31/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo photograph_find_maint_Click_Error
'
'******** Code Starts Here *************
'
    Screen.MousePointer = ccHourglass
    frm_photograph_find_maint.Show
    Screen.MousePointer = ccDefault
'
'********* Code Ends Here **************
'
   Exit Sub
'
photograph_find_maint_Click_Error:
    DisplayError , "mdi_npls.photograph_find_maint_Click", vbExclamation
End Sub

Private Sub Photograph_maint_Click()
'***************************************
' Module/Form Name   : mdi_npls
'
' Procedure Name     : Photograph_maint_Click
'
' Purpose            :
'
' Date Created       : 31/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo Photograph_maint_Click_Error
'
'******** Code Starts Here *************
'
    frm_photograph_maint.Show
'
'********* Code Ends Here **************
'
   Exit Sub
'
Photograph_maint_Click_Error:
    DisplayError , "mdi_npls.Photograph_maint_Click", vbExclamation
End Sub

Private Sub print_photo_labels_Click()
'***************************************
' Module/Form Name   : mdi_npls
'
' Procedure Name     : print_photo_labels_Click
'
' Purpose            :
'
' Date Created       : 31/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo print_photo_labels_Click_Error
'
'******** Code Starts Here *************
'
    frm_print_photo_labels.Show 1
'
'********* Code Ends Here **************
'
   Exit Sub
'
print_photo_labels_Click_Error:
    DisplayError , "mdi_npls.print_photo_labels_Click", vbExclamation
End Sub

Public Sub RefreshCaption()
    Me.Caption = "Image Where for " & goCompanyInfo.CompanyName
End Sub

Private Sub Tile_Click()
    mdi_npls.Arrange vbTileHorizontal
End Sub

Private Sub Tim_number_of_photos_Timer()
'***************************************
' Module/Form Name   : mdi_npls
'
' Procedure Name     : Tim_number_of_photos_Timer
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
On Error GoTo Tim_number_of_photos_Timer_Error
'
'******** Code Starts Here *************
'
    display_last_photo_number_used
    '
    '   Use this timer to decide whether to post a log file to support.
    '
    If Not IsInIDE Then
        If goSystemConfig.LogFileShouldBePosted Then
            MsgBox "The Log File was last posted on '" & Format(goSystemConfig.DateLogFilePosted, "Short Date") & "'" & vbCrLf & _
                   "Please go to tools/options, select the Support Tab and click on 'Post'", vbInformation
            goSystemConfig.DatePreviousPostWarning = Now()
            goSystemConfig.Update
        End If
        '
        If goSystemConfig.SupportUser Then
            SetDetailedAuditing
        End If
    End If
'
'********* Code Ends Here **************
'
    Exit Sub
    '
Tim_number_of_photos_Timer_Error:
    DisplayError , "mdi_npls.Tim_number_of_photos_Timer", vbExclamation
End Sub

Private Sub SetDetailedAuditing()
'***************************************
' Module/Form Name   : mdi_npls
'
' Procedure Name     : SetDetailedAuditing
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
On Error GoTo SetDetailedAuditing_Error
'
'******** Code Starts Here *************
'
    Dim oFSO                        As Scripting.FileSystemObject
    Dim strIniFile                  As String
    Dim strOutput                   As String
    Dim intFileNo                   As Integer
    
    Set oFSO = New Scripting.FileSystemObject
    '
    strIniFile = oFSO.BuildPath(gstrAppPath, "IW.ini")
    intFileNo = FreeFile
    '
    '   Does the auditing need to be set?
    '
    If goSystemConfig.DateDetailedAuditingRequested <> 0 And goSystemConfig.DateDetailedAuditingRequested < Now Then
        If goSystemConfig.DateDetailedAuditingSet < goSystemConfig.DateDetailedAuditingRequested Then
            goSystemConfig.DetailedAuditing = goSystemConfig.DetailedAuditingRequest
            goSystemConfig.DateDetailedAuditingSet = Now()
            goSystemConfig.Update
        End If
    End If
                                    
'
'********* Code Ends Here **************
'
   Exit Sub
'
SetDetailedAuditing_Error:
    ErrorRaise "mdi_npls.SetDetailedAuditing"
End Sub

Private Sub tob_npls_ButtonClick(ByVal Button As Button)
'***************************************
' Module/Form Name   : mdi_npls
'
' Procedure Name     : tob_npls_ButtonClick
'
' Purpose            :
'
' Date Created       : 19/05/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : Button - Button
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 19/05/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo tob_npls_ButtonClick_Error
'
'******** Code Starts Here *************
'
    Select Case Button.Key
        Case "Customer"
            frm_cust_maint.Show
            frm_cust_maint.SetFocus
        Case "Photograph"
            Screen.MousePointer = ccHourglass
            With frm_photograph_maint
                .Show
                .SetFocus
            End With
            Screen.MousePointer = ccDefault
        Case "Job"
            Screen.MousePointer = ccHourglass
            frm_job_maint.Display
            On Error Resume Next
            frm_job_maint.SetFocus
            On Error GoTo tob_npls_ButtonClick_Error
            Screen.MousePointer = ccDefault
        Case "Delivery Note"
            frm_delivery_note.load_delivery_notes 0, "0", 0
        Case "Photo Labels"
            frm_print_photo_labels.Show 1
        Case "KeywordClustering"
            Call KeywordClustering
    End Select
'
'********* Code Ends Here **************
'
    Exit Sub
    '
tob_npls_ButtonClick_Error:
    DisplayError , "mdi_npls.tob_npls_ButtonClick", vbExclamation
End Sub

Private Sub KeywordClustering()
    Dim fKeywordClustering  As frmKeywordClustering
    
'***************************************
' Module/Form Name   : mdi_npls
'
' Procedure Name     : KeywordClustering
'
' Purpose            :
'
' Date Created       : 25/07/2005
'
' Author             : Gareth Saunders
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                      25/07/2005 Gareth Saunders
'
'***************************************
'
    On Error GoTo ErrorHandler:
'
'******** Code Starts Here *************
'
    Set fKeywordClustering = New frmKeywordClustering
    
    fKeywordClustering.Display
    
    Set fKeywordClustering = Nothing
'
'********* Code Ends Here **************
'
    On Error GoTo 0
    Exit Sub
'
ErrorHandler:
    ErrorRaise "mdi_npls.KeywordClustering"
End Sub

Private Sub tool_bar_Click()
    If tob_npls.Visible = True Then
        tob_npls.Visible = False
        tool_bar.Checked = False
    Else
        tob_npls.Visible = True
        tool_bar.Checked = True
    End If
End Sub

Private Function UpgradeExists() As Boolean
'***************************************
' Module/Form Name   : mdi_npls
'
' Procedure Name     : UpgradeExists
'
' Purpose            :
'
' Date Created       : 05/06/2002
'
' Author             : GARETH SAUNDERS
'
' Returns            : Boolean
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 22/12/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo UpgradeExists_Error
'
'******** Code Starts Here *************
'
    Dim oFSO As Scripting.FileSystemObject
''    Dim oDNCVersion As DNCFunctions.Version
    Dim oDNCVersion As Object
    Dim oUpgrade As Object
    Dim strNewVersion As String
    Dim strCurrentVersion As String
    Dim strNewExe As String
    '
    '   Indicate that the user is a Support User.
    '
    goSystemConfig.SupportUser = True
    '
    '   Apply any database updates here.
    '
''    ApplyDBUpdates
    '
    UpgradeExists = False

    Set oFSO = New Scripting.FileSystemObject
''    Set oDNCVersion = New DNCFunctions.Version
    Set oDNCVersion = CreateObject("DNCFunctions.Version")
    '
    '   Check if iw.exe exists.
    '
    strNewExe = oFSO.BuildPath(gstrUpgradePath, "iw.exe")
    If Not oFSO.FileExists(strNewExe) Then Exit Function
    '
    '   Check if the version is different.
    '
    strNewVersion = oDNCVersion.GetFileVersion(strNewExe)
    strCurrentVersion = oDNCVersion.GetFileVersion(gstrAppPath & "\" & App.EXEName & ".exe")
    If strCurrentVersion >= strNewVersion Then Exit Function
    '
    '   Invoke the Upgrade component if required.
    '
    If MsgBox("A new version of Image Where exists. Do you wish to automatically upgrade?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Function
    MsgBox "Please ensure that all other users are logged out of the system before continuing.", vbInformation, App.Title
    '
'    Set oUpgrade = New IWUpgrade.Upgrade
    Set oUpgrade = CreateObject("IWUpgrade.Upgrade")
    oUpgrade.Execute ApplicationVersion, strNewVersion
    Set oUpgrade = Nothing
    UpgradeExists = True
    '
    '********* Code Ends Here **************
    '
    Exit Function
    '
UpgradeExists_Error:
    ErrorRaise "mdi_npls.UpgradeExists"
End Function

Private Sub ApplyDBUpdates()
    Dim blnFullyUpgraded As Boolean
    
''    Dim oAccessDB As access.Application
    Dim oAccessDB As Object

    Set oAccessDB = GetObject(glo_dbname, "Access.Application")
    oAccessDB.Visible = True
    oAccessDB.DoCmd.Maximize
    oAccessDB.Modules.Item(0).ReplaceLine 17, "           & ""WHERE SEARCH_RESULT.DELIVERY_NOTE_NO = "" & delivery_note_no & "" AND (DATE_RETURNED IS NULL OR DATE_RETURNED = 0)"
    Set oAccessDB = Nothing
    '
    '   Audit the change.
    '
    If blnFullyUpgraded Then
        WriteUpgradeLog "Report has been fixed"
    End If
'
'********* Code Ends Here **************
'
   Exit Sub
'
ApplyDBUpdates_Error:
    ErrorRaise "mdi_npls.ApplyDBUpdates"
End Sub

Public Sub WriteUpgradeLog(strDescription As String)
'***************************************
' Module/Form Name   : mdi_npls
'
' Procedure Name     : WriteUpgradeLog
'
' Purpose            :
'
' Date Created       : 25/02/2003
'
' Author             : GARETH SAUNDERS
'
' Parameters         : strDescription - String
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo WriteUpgradeLog_Error
'
'******** Code Starts Here *************
'
    Dim intFileNo As Integer
    '
    '   Open the log file.
    '
    intFileNo = FreeFile(1)
    Open gstrUpgradePath & "\Upgrade.log" For Append As intFileNo

    Write #intFileNo, _
          Format(Now, "dd/mm/yyyy hh:mm:ss"), _
          strDescription
    '
    Close intFileNo
'
'********* Code Ends Here **************
'
   Exit Sub
'
WriteUpgradeLog_Error:
    ErrorRaise "mdi_npls.WriteUpgradeLog"
End Sub

