VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CFC13920-9EF4-11D0-B72F-0000C04D4C0A}#6.0#0"; "mswless.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_search_photograph 
   Caption         =   "Find photographs"
   ClientHeight    =   6000
   ClientLeft      =   3600
   ClientTop       =   3375
   ClientWidth     =   7935
   ClipControls    =   0   'False
   Icon            =   "Search Photographs.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6000
   ScaleWidth      =   7935
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraSearch 
      Height          =   975
      Left            =   120
      TabIndex        =   21
      Top             =   0
      Width           =   7575
      Begin VB.OptionButton optKeywords 
         Caption         =   "Keywords"
         Height          =   255
         Left            =   6240
         TabIndex        =   25
         Top             =   600
         Width           =   1200
      End
      Begin VB.OptionButton optCaption 
         Caption         =   "Caption"
         Height          =   255
         Left            =   6240
         TabIndex        =   24
         Top             =   240
         Value           =   -1  'True
         Width           =   1200
      End
      Begin VB.TextBox txt_description 
         Height          =   615
         HelpContextID   =   4
         Left            =   1080
         MaxLength       =   120
         MultiLine       =   -1  'True
         TabIndex        =   22
         Top             =   240
         Width           =   5055
      End
      Begin VB.Label lab_description 
         Caption         =   "&Description:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   300
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "F&ind"
      Height          =   675
      Left            =   480
      Picture         =   "Search Photographs.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4800
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "S&elect"
      Height          =   675
      Left            =   3120
      Picture         =   "Search Photographs.frx":0B04
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4800
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
      Height          =   675
      Left            =   4320
      Picture         =   "Search Photographs.frx":107E
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4800
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   675
      Left            =   6720
      Picture         =   "Search Photographs.frx":1740
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4800
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   675
      Left            =   5520
      Picture         =   "Search Photographs.frx":1CD2
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4800
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.Timer tim_Tooltip 
      Interval        =   1000
      Left            =   2520
      Top             =   4560
   End
   Begin TabDlg.SSTab sstFind 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   5953
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Text"
      TabPicture(0)   =   "Search Photographs.frx":2394
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblPhoto"
      Tab(0).Control(1)=   "lblDescription"
      Tab(0).Control(2)=   "lis_photographs"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Thumbnail"
      TabPicture(1)   =   "Search Photographs.frx":23B0
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblCurrentPage"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "picFind"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdNext"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdPrevious"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "&Previous"
         Enabled         =   0   'False
         Height          =   315
         Left            =   300
         TabIndex        =   2
         Top             =   420
         Width           =   795
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next"
         Enabled         =   0   'False
         Height          =   315
         Left            =   1260
         TabIndex        =   3
         Top             =   420
         Width           =   795
      End
      Begin VB.PictureBox picFind 
         ClipControls    =   0   'False
         Height          =   2355
         Left            =   240
         ScaleHeight     =   2295
         ScaleWidth      =   6915
         TabIndex        =   7
         Top             =   900
         Width           =   6975
         Begin VB.TextBox txtToolTip 
            BackColor       =   &H80000018&
            Height          =   375
            Left            =   4260
            MultiLine       =   -1  'True
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   1380
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.PictureBox picFindLarge 
            BorderStyle     =   0  'None
            ClipControls    =   0   'False
            Height          =   2055
            Left            =   120
            ScaleHeight     =   2055
            ScaleWidth      =   6075
            TabIndex        =   10
            Top             =   60
            Width           =   6075
            Begin VB.CommandButton cmdScreenRes 
               Height          =   255
               Index           =   0
               Left            =   2520
               MaskColor       =   &H00FF00FF&
               Picture         =   "Search Photographs.frx":23CC
               Style           =   1  'Graphical
               TabIndex        =   20
               Top             =   1440
               UseMaskColor    =   -1  'True
               Width           =   255
            End
            Begin VB.CheckBox chkSelected 
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   5
               Top             =   1560
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.Line linExclusive2 
               BorderWidth     =   5
               Index           =   0
               Visible         =   0   'False
               X1              =   4920
               X2              =   4080
               Y1              =   600
               Y2              =   960
            End
            Begin VB.Line linExclusive1 
               BorderWidth     =   5
               Index           =   0
               Visible         =   0   'False
               X1              =   4080
               X2              =   4920
               Y1              =   600
               Y2              =   960
            End
            Begin VB.Image imgA4 
               Height          =   255
               Index           =   0
               Left            =   3000
               Picture         =   "Search Photographs.frx":26DE
               Stretch         =   -1  'True
               Top             =   1560
               Width           =   255
            End
            Begin VB.Label lblStretch 
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   4080
               TabIndex        =   12
               Top             =   120
               Visible         =   0   'False
               Width           =   1455
               WordWrap        =   -1  'True
            End
            Begin VB.Label lblFind 
               BorderStyle     =   1  'Fixed Single
               Height          =   1455
               Index           =   0
               Left            =   2280
               TabIndex        =   11
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Shape shpSelected 
               BorderColor     =   &H8000000D&
               BorderWidth     =   4
               Height          =   495
               Index           =   0
               Left            =   1440
               Top             =   1440
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.Image imgFind 
               Height          =   1455
               Index           =   0
               Left            =   0
               Stretch         =   -1  'True
               Top             =   0
               Visible         =   0   'False
               Width           =   1455
            End
         End
         Begin MSWLess.WLVScroll vscFind 
            Height          =   2595
            Left            =   6625
            TabIndex        =   6
            Top             =   60
            Width           =   230
            _ExtentX        =   397
            _ExtentY        =   4577
            _Version        =   393216
         End
      End
      Begin VB.ListBox lis_photographs 
         Height          =   2400
         Left            =   -74880
         MultiSelect     =   1  'Simple
         TabIndex        =   1
         Top             =   720
         WhatsThisHelpID =   5
         Width           =   7095
      End
      Begin VB.Label lblDescription 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Description"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -73200
         TabIndex        =   19
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lblPhoto 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Photo"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -74880
         TabIndex        =   18
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lblCurrentPage 
         Caption         =   "Page 1 of 1"
         Height          =   195
         Left            =   2280
         TabIndex        =   4
         Top             =   480
         Width           =   1335
      End
   End
   Begin MSComDlg.CommonDialog cdgHelp 
      Left            =   1920
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      HelpContext     =   1
   End
   Begin MSComctlLib.StatusBar stbSearchFind 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   5745
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4471
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4471
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4471
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_search_photograph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const CONST_IMAGE_TOP = 200
Private Const CONST_SCROLL_BAR_LEFT_OFFSET = 40
Private Const CONST_SCROLL_BAR_HEIGHT_OFFSET = 40

Private snap_batch As DAO.Recordset
Private first_time As Boolean
Public Event PhotographSelected(oSearch As Search2, oDigitalImage As DigitalImage)
Public Event PhotographSelectionComplete()
Private moDigitalImages As DigitalImages
'
'   Used to determine the Digital Images collection indexes.
'
Private mintPictureImagesCount As Integer
Private mintPictureTextCount As Integer
Private mintPicturesDisplayed As Integer
'
'   Used to determine the Digital Images collection indexes.
'
Private mintPageImagesCount As Integer
Private mintPageTextCount As Integer
Private mintPageDisplayed As Integer
'
Private mintMaxImagesPerRow As Single
Private mdblRowsOfImages As Double
Private mdblRowsDisplayed As Double
Private mCtlCln As Collection
Private blnThumbClicked As Boolean
Private lngTimeTipHidden As Long
Private mLogicalImagesHeight As Double
Private mintImageHeight As Integer
Private mintImageWidth As Integer
Private mintCurrentPage As Integer
Private mintNoOfPages As Integer
'
'   Used to control whether the chkSelected_click event does any processing.
'
Private mblnSelectText As Boolean
'
Public Enum enmSearchRefine
    srSsearch
    srRefine
    srreview
End Enum
Private mblnEventRequired       As Boolean
Private mSearchRefine           As enmSearchRefine
Private moSearch                As Search2
Private moJob                   As Job2
Private mblnUpdateScreenOnly    As Boolean
Private mcolDeDupe              As New Collection
'
'   Tool Tip parameters.
'
Private mblnToolTipDisplayed    As Boolean
Private mdteTooltipStart        As Date
Private mblnTooltipClock        As Boolean
Private mblnTooltipEnabled      As Boolean
Private mdteTooltipLeftImage    As Date

Public Sub display_screen(ByVal pSearchRefine As enmSearchRefine, _
                          ByRef poSearch As Search2, _
                          ByRef poJob As Job2)
'***************************************
' Module/Form Name   : frm_search_photograph
'
' Procedure Name     : display_screen
'
' Purpose            :
'
' Date Created       : 25/06/2004
'
' Author             : ADMINISTRATOR
'
' Parameters         : pSearchRefine - enmSearchRefine
'                    : poSearch - Search2
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo display_screen_Error
'
'******** Code Starts Here *************
'
    '
    '   Set the module level variables.
    '
    mSearchRefine = pSearchRefine
    Set moSearch = poSearch
    Set moJob = poJob
    '
    If mSearchRefine = srreview Then
        cmdFind.Enabled = False
        cmdSelect.Enabled = False
        cmdClear.Enabled = False
        cmdHelp.Enabled = False
    ElseIf mSearchRefine = srRefine Then
        cmdClear.Enabled = False
    End If
    '
    If mSearchRefine = srRefine Or _
       mSearchRefine = srreview Then
        SetUpScreenForRefining
    End If
    Me.Show 1
'
'********* Code Ends Here **************
'
   Exit Sub
'
display_screen_Error:
    DisplayError , "frm_search_photograph.display_screen", vbExclamation
End Sub

Private Sub SetUpScreenForRefining()
'***************************************
' Module/Form Name   : frm_search_photograph
'
' Procedure Name     : SetUpScreenForRefining
'
' Purpose            :
'
' Date Created       : 24/06/2004
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo SetUpScreenForRefining_Error
'
'******** Code Starts Here *************
'
    Dim intIdx      As Integer
    
    If mSearchRefine = srRefine Then
        Me.Caption = "Refine Photographs for Search '" & moSearch.Description & "'"
    Else
        Me.Caption = "Review Photographs for Search '" & moSearch.Description & "'"
    End If
    '
    '   Need to set the Images Per Row constant as the form hasn't been displayed yet.
    '
    mintMaxImagesPerRow = Int((picFind.Width - 100 - vscFind.Width - CONST_SCROLL_BAR_LEFT_OFFSET) / (mintImageWidth + 100))
    '
    mblnUpdateScreenOnly = True
    DisplayPhotographs False
    mblnUpdateScreenOnly = False
    '
    '   Select all the ListBox items.
    '
    mblnEventRequired = False
    For intIdx = 0 To lis_photographs.ListCount - 1
        lis_photographs.Selected(intIdx) = True
    Next intIdx
    mblnEventRequired = True
'
'********* Code Ends Here **************
'
   Exit Sub
'
SetUpScreenForRefining_Error:
    ErrorRaise "frm_search_photograph.SetUpScreenForRefining"
End Sub

Private Function format_photograph(batch_no As Long, batch_description As String) As String
10        format_photograph = CStr(batch_no) & vbTab & batch_description
End Function

Private Sub chkSelected_Click(Index As Integer)
'***************************************
' Module/Form Name   : frm_search_photograph
'
' Procedure Name     : chkSelected_Click
'
' Purpose            :
'
' Date Created       : 15/02/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : Index - Integer
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 15/02/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo chkSelected_Click_Error
'
'******** Code Starts Here *************
'
    If Not mblnEventRequired Then Exit Sub
    '
    '   If only reviewing, don't allow selection or deselection.
    '
    If mSearchRefine = srreview Then
        Exit Sub
    End If
    '
    '   Display message if image is exclusive.
    '
    If Not mblnUpdateScreenOnly Then
        '
        '   If selecting the phtograph, check for exclusivity.
        '
        If chkSelected(Index).Value = vbChecked Then
            If moDigitalImages(CStr(Index + (mintCurrentPage - 1) * goSystemConfig.ImagesPerPage)).Exclusive Then
                MsgBox "This is an Exclusive Image. Please select another Image.", vbExclamation
                mblnEventRequired = False
                chkSelected(Index).Value = vbUnchecked
                mblnEventRequired = True
                Exit Sub
            End If
        End If
    End If
    '
    If mblnSelectText Then
        blnThumbClicked = True
        first_time = False
        lis_photographs.ListIndex = Index + (mintCurrentPage - 1) * goSystemConfig.ImagesPerPage
        first_time = True
        If Not lis_photographs.Selected(Index + (mintCurrentPage - 1) * goSystemConfig.ImagesPerPage) = CBool(chkSelected(Index).Value) Then
            lis_photographs.Selected(Index + (mintCurrentPage - 1) * goSystemConfig.ImagesPerPage) = chkSelected(Index).Value
        End If
        ToggleDigitalImageSelection Index
        blnThumbClicked = False
    End If
'
'********* Code Ends Here **************
'
    Exit Sub
        '
chkSelected_Click_Error:
    DisplayError , "frm_search_photograph.chkSelected_Click", vbExclamation
End Sub

Private Sub chkSelected_GotFocus(Index As Integer)
'***************************************
' Module/Form Name   : frm_search_photograph
'
' Procedure Name     : chkSelected_GotFocus
'
' Purpose            :
'
' Date Created       : 15/02/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : Index - Integer
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 15/02/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo chkSelected_GotFocus_Error
'
'******** Code Starts Here *************
'
    first_time = False
    lis_photographs.ListIndex = Index
    UpdateStatusBar
'
'********* Code Ends Here **************
'
    Exit Sub
    '
chkSelected_GotFocus_Error:
    DisplayError , "frm_search_photograph.chkSelected_GotFocus", vbExclamation
End Sub

Private Sub cmdNext_Click()
'***************************************
' Module/Form Name   : frm_search_photograph
'
' Procedure Name     : cmdNext_Click
'
' Purpose            :
'
' Date Created       : 15/02/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 15/02/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo cmdNext_Click_Error
'
'******** Code Starts Here *************
'
    Dim intImageIndex As Integer
    
    Screen.MousePointer = vbHourglass
    SetUpPage 1
    For intImageIndex = (mintCurrentPage - 1) * goSystemConfig.ImagesPerPage To (mintCurrentPage) * goSystemConfig.ImagesPerPage - 1
        If intImageIndex > moDigitalImages.Count - 1 Then
            Exit For
        End If
        DisplayDigitalImage moDigitalImages(CStr(intImageIndex))
    Next intImageIndex
    ConfigureScrollBar
    Screen.MousePointer = vbDefault
'
'********* Code Ends Here **************
'
    Exit Sub
    '
cmdNext_Click_Error:
    DisplayError , "frm_search_photograph.cmdNext_Click", vbExclamation
End Sub

Private Sub cmdPrevious_Click()
'***************************************
' Module/Form Name   : frm_search_photograph
'
' Procedure Name     : cmdPrevious_Click
'
' Purpose            :
'
' Date Created       : 15/02/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 15/02/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo cmdPrevious_Click_Error
'
'******** Code Starts Here *************
'
    Dim intImageIndex As Integer

    Screen.MousePointer = vbHourglass
    SetUpPage -1
    For intImageIndex = (mintCurrentPage - 1) * goSystemConfig.ImagesPerPage To (mintCurrentPage) * goSystemConfig.ImagesPerPage - 1
        If intImageIndex > moDigitalImages.Count - 1 Then
            Exit For
        End If
        DisplayDigitalImage moDigitalImages(CStr(intImageIndex))
    Next intImageIndex

    ConfigureScrollBar
    Screen.MousePointer = vbDefault
    '
    '********* Code Ends Here **************
    '
    Exit Sub
    '
cmdPrevious_Click_Error:
    DisplayError , "frm_search_photograph.cmdPrevious_Click", vbExclamation
End Sub

Private Sub cmdScreenRes_Click(Index As Integer)
''    frmScreenRes.Display moDigitalImages.Item(Index + 1 + (mintCurrentPage - 1) * goSystemConfig.ImagesPerPage).FileLocation("SC")
    frmScreenRes.Display moDigitalImages.Item(Index + 1 + (mintCurrentPage - 1) * goSystemConfig.ImagesPerPage)
End Sub

Private Sub Form_Initialize()
    mblnEventRequired = True
    mblnUpdateScreenOnly = False
End Sub

Private Sub Form_Load()
'***************************************
' Module/Form Name   : frm_search_photograph
'
' Procedure Name     : Form_Load
'
' Purpose            :
'
' Date Created       : 15/02/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 15/02/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo Form_Load_Error
'
'******** Code Starts Here *************
'
    Dim TabStop As Long

    msg_title = "Locate Picture"

    tim_Tooltip.Interval = Time_When_ToolTip_Not_Displayed
    com_position_form Me
    '
    '   Set the boolean 'first_time' etc. to true on loading used in photograph selection
    '
    first_time = True
    blnThumbClicked = False
    lngTimeTipHidden = 0
    '
    '   Create the Digital Image Collection and Initialise Counts.
    '
    Set moDigitalImages = New DigitalImages
    mintPictureImagesCount = 0
    mintPictureTextCount = 0
    mintPicturesDisplayed = 0
    mintImageHeight = goSystemConfig.ImageHeightWidth
    mintImageWidth = goSystemConfig.ImageHeightWidth
    mintCurrentPage = 1
    mintNoOfPages = 1
    mblnSelectText = True
    '
    '   Set Tab Stops for the ListView.
    '
    TabStop = 45
    'clear then set the tabs
    Call SendMessage(lis_photographs.hWnd, LB_SETTABSTOPS, 0&, ByVal 0&)
    Call SendMessage(lis_photographs.hWnd, LB_SETTABSTOPS, 1, TabStop)
    lis_photographs.Refresh
    '
    '   Set form size and tab and Picture container.
    '
    txt_description.Text = GetSetting(App.Title, "SearchFind", "Description", "")
    optCaption.Value = IIf(GetSetting(App.Title, "SearchFind", "Caption", "Y") = "Y", True, False)
    optKeywords.Value = Not optCaption.Value
    sstFind.Tab = GetSetting(App.Title, "SearchFind", "Tab", 0)
    Me.WindowState = GetSetting(App.Title, "SearchFind", "WindowState", 0)
    If GetSetting(App.Title, "SearchFind", "WindowState", 0) <> 2 Then
        Me.Width = GetSetting(App.Title, "SearchFind", "Width", 8025)
        Me.Height = GetSetting(App.Title, "SearchFind", "Height", 6000)
    End If
    picFindLarge.Move 0, 0, picFind.Width - CONST_SCROLL_BAR_LEFT_OFFSET - goSystemConfig.VScrollBarWidth, picFind.Height
    '
    '   Detect MouseWheel events.
    '
    If goSystemConfig.MouseWheel Then
        MouseWheel.Hook Me.hWnd
    End If
    '
    Screen.MousePointer = vbDefault
    '
    '********* Code Ends Here **************
    '
    Exit Sub
    '
Form_Load_Error:
    DisplayError , "frm_search_photograph.Form_Load", vbExclamation
End Sub

Private Sub Form_Resize()
'***************************************
' Module/Form Name   : frm_search_photograph
'
' Procedure Name     : Form_Resize
'
' Purpose            :
'
' Date Created       : 15/02/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 15/02/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo Form_Resize_Error
'
'******** Code Starts Here *************
'
    Dim MIN_WIDTH As Integer
    Dim MIN_HEIGHT As Integer
    Dim oDigitalImage As DigitalImage
    Dim intImageIndex As Integer

    MIN_WIDTH = 7395
    MIN_HEIGHT = 5800

    If WindowState = 1 Then
        Exit Sub
    End If

    If WindowState = 0 Then
        If Width < MIN_WIDTH Then Width = MIN_WIDTH
        If Height < MIN_HEIGHT Then Height = MIN_HEIGHT
    End If

    sstFind.Visible = False
    stbSearchFind.Visible = False
    DoEvents
    '
    '   Determine those controls with bad left properties!
    '
    FixBadControls
    cmdPrevious.Move 200, 400
    cmdNext.Move cmdPrevious.Left + cmdPrevious.Width + 200, cmdPrevious.Top
    lblCurrentPage.Move cmdNext.Left + cmdNext.Width + 400, cmdNext.Top

    cmdSelect.Left = (Me.Width - cmdSelect.Width - 230) / 4
    cmdSelect.Top = Me.Height - cmdSelect.Height - 550 - stbSearchFind.Height

    cmdCancel.Left = Me.Width - cmdCancel.Width - 230
    cmdCancel.Top = cmdSelect.Top

    cmdHelp.Move (Me.Width - cmdSelect.Width - 230) * 3 / 4, _
               cmdSelect.Top

    cmdClear.Move (Me.Width - cmdSelect.Width - 230) * 2 / 4, _
                cmdSelect.Top

    cmdFind.Left = lis_photographs.Left
    cmdFind.Top = cmdSelect.Top

    fraSearch.Move fraSearch.Left, fraSearch.Top, Me.Width - 2 * lis_photographs.Left - 100, 900
    optCaption.Move fraSearch.Width - optCaption.Width - 100, txt_description.Top
    optKeywords.Move optCaption.Left, optCaption.Top + optCaption.Height
    txt_description.Move txt_description.Left, txt_description.Top, optCaption.Left - txt_description.Left - 100, 500
    
    sstFind.Width = fraSearch.Width
    sstFind.Height = cmdFind.Top - sstFind.Top - 100
    sstFind.Top = fraSearch.Top + fraSearch.Height + 100

    lblPhoto.Move lblPhoto.Left, _
                sstFind.TabHeight + 100, 1000

    lblDescription.Move lblPhoto.Left + lblPhoto.Width, _
                      lblPhoto.Top, _
                      sstFind.Width - 2 * lblPhoto.Left - lblDescription.Left + 100

    lis_photographs.Move lblPhoto.Left, _
                       lblPhoto.Top + lblPhoto.Height, _
                       lblPhoto.Width + lblDescription.Width, _
                       sstFind.Height - lblPhoto.Top - lblPhoto.Height - 100

    With picFind
        .Top = cmdNext.Top + cmdNext.Height + 100
        .Left = 100
        .Width = sstFind.Width - 200
        .Height = sstFind.Height - .Top - 100
    End With
    '
    '   If the image tab is showing, reposition the images and text boxes and position and
    '   configure the scroll bar. First calculate the maximum number of images per row.
    '
    mintMaxImagesPerRow = Int((picFind.Width - 100 - vscFind.Width - CONST_SCROLL_BAR_LEFT_OFFSET) / (mintImageWidth + 100))
    ConfigureScrollBar
    DoEvents
    For intImageIndex = (mintCurrentPage - 1) * goSystemConfig.ImagesPerPage To (mintCurrentPage) * goSystemConfig.ImagesPerPage - 1
        If intImageIndex >= moDigitalImages.Count Then
            Exit For
        End If
        Set oDigitalImage = moDigitalImages(CStr(intImageIndex))
        If oDigitalImage.ImageAvailable Then
            PositionImageControl imgFind(oDigitalImage.ImageControlIndex), _
                                 chkSelected(oDigitalImage.DigitalImageIndex - (mintCurrentPage - 1) * goSystemConfig.ImagesPerPage), _
                                 shpSelected(oDigitalImage.DigitalImageIndex - (mintCurrentPage - 1) * goSystemConfig.ImagesPerPage), _
                                 linExclusive1(oDigitalImage.DigitalImageIndex - (mintCurrentPage - 1) * goSystemConfig.ImagesPerPage), _
                                 linExclusive2(oDigitalImage.DigitalImageIndex - (mintCurrentPage - 1) * goSystemConfig.ImagesPerPage), _
                                 cmdScreenRes(oDigitalImage.DigitalImageIndex - (mintCurrentPage - 1) * goSystemConfig.ImagesPerPage), _
                                 imgA4(oDigitalImage.DigitalImageIndex - (mintCurrentPage - 1) * goSystemConfig.ImagesPerPage), _
                                 oDigitalImage
        Else
            PositionImageControl lblFind(oDigitalImage.TextControlIndex), _
                                 chkSelected(oDigitalImage.DigitalImageIndex - (mintCurrentPage - 1) * goSystemConfig.ImagesPerPage), _
                                 shpSelected(oDigitalImage.DigitalImageIndex - (mintCurrentPage - 1) * goSystemConfig.ImagesPerPage), _
                                 linExclusive1(oDigitalImage.DigitalImageIndex - (mintCurrentPage - 1) * goSystemConfig.ImagesPerPage), _
                                 linExclusive2(oDigitalImage.DigitalImageIndex - (mintCurrentPage - 1) * goSystemConfig.ImagesPerPage), _
                                 cmdScreenRes(oDigitalImage.DigitalImageIndex - (mintCurrentPage - 1) * goSystemConfig.ImagesPerPage), _
                                 imgA4(oDigitalImage.DigitalImageIndex - (mintCurrentPage - 1) * goSystemConfig.ImagesPerPage), _
                                 oDigitalImage
        End If
    Next intImageIndex
    '
    UnFixBadControls
    stbSearchFind.Visible = True
    sstFind.Visible = True
'
'********* Code Ends Here **************
'
    Exit Sub
    '
Form_Resize_Error:
    stbSearchFind.Visible = True
    sstFind.Visible = True
    DisplayError , "frm_search_photograph.Form_Resize", vbExclamation
End Sub

Private Sub Form_Unload(Cancel As Integer)
'***************************************
' Module/Form Name   : frm_search_photograph
'
' Procedure Name     : Form_Unload
'
' Purpose            :
'
' Date Created       : 15/02/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : Cancel - Integer
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 15/02/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo Form_Unload_Error
'
'******** Code Starts Here *************
'
    Screen.MousePointer = vbHourglass
    SaveSetting App.Title, "SearchFind", "Description", txt_description.Text
    SaveSetting App.Title, "SearchFind", "Tab", sstFind.Tab
    SaveSetting App.Title, "SearchFind", "Width", Me.Width
    SaveSetting App.Title, "SearchFind", "Height", Me.Height
    SaveSetting App.Title, "SearchFind", "WindowState", Me.WindowState
    SaveSetting App.Title, "SearchFind", "Caption", IIf(optCaption.Value, "Y", "N")
    ClearDisplay
    DoEvents
    '
    '   Unhook from MouseWheel events.
    '
    If goSystemConfig.MouseWheel Then
        MouseWheel.UnHook
    End If
    '
    Screen.MousePointer = vbDefault
'
'********* Code Ends Here **************
'
    Exit Sub
    '
Form_Unload_Error:
    Cancel = False
    DisplayError , "frm_search_photograph.Form_Unload", vbExclamation
End Sub

Private Sub imgFind_Click(Index As Integer)
'***************************************
' Module/Form Name   : frm_search_photograph
'
' Procedure Name     : imgFind_Click
'
' Purpose            :
'
' Date Created       : 27/12/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : Index - Integer
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo imgFind_Click_Error
'
'******** Code Starts Here *************
'
    '
    '   If only reviewing, don't allow selection or deselection.
    '
    If mSearchRefine = srreview Then
        Exit Sub
    End If
    '
    '   Display message if image is exclusive.
    '
    If Not moDigitalImages(imgFind(Index).Tag).Selected Then
        If moDigitalImages(imgFind(Index).Tag).Exclusive Then
            MsgBox "This is an Exclusive Image. Please select another Image.", vbExclamation
            Exit Sub
        End If
    End If
    '
    first_time = False
    lis_photographs.ListIndex = CInt(imgFind(Index).Tag)
    first_time = True
    lis_photographs.Selected(CInt(imgFind(Index).Tag)) = Not lis_photographs.Selected(CInt(imgFind(Index).Tag))
    '
    '   If refining a search then, set the Pending flag accordingly.
    '
''    If mSearchRefine = Refine Then
''        MsgBox CInt(imgFind(Index).Tag)
''    End If
'
'********* Code Ends Here **************
'
   Exit Sub
'
imgFind_Click_Error:
    DisplayError , "frm_search_photograph.imgFind_Click", vbExclamation
End Sub

Private Sub lis_photographs_Click()
'***************************************
' Module/Form Name   : frm_search_photograph
'
' Procedure Name     : lis_photographs_Click
'
' Purpose            :
'
' Date Created       : 27/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo lis_photographs_Click_Error
'
'******** Code Starts Here *************
'
    Dim oDigitalSearch      As Search2
    Dim oDigitalImage       As DigitalImage
    Dim blnInStock          As Boolean
    Dim oSearchResult       As SearchResult
    Dim blnValid            As Boolean
    '
    '   Check if this event is required.
    '
    If Not mblnEventRequired Then Exit Sub
    '
    '   Select the Digital Image also provided this subroutine hasn't been invoked by
    '   clicking on the checkbox itself.
    '
    If blnThumbClicked Then
        blnThumbClicked = False
    Else
        '
        '   Don't select image if it is not currently displayed.
        '
        If lis_photographs.ListIndex <= (mintCurrentPage * goSystemConfig.ImagesPerPage) - 1 And _
           lis_photographs.ListIndex >= (mintCurrentPage - 1) * goSystemConfig.ImagesPerPage Then
            If Not lis_photographs.Selected(lis_photographs.ListIndex) = CBool(chkSelected(lis_photographs.ListIndex - (mintCurrentPage - 1) * goSystemConfig.ImagesPerPage).Value) Then
                If chkSelected(lis_photographs.ListIndex - (mintCurrentPage - 1) * goSystemConfig.ImagesPerPage).Value = 0 Then
                    chkSelected(lis_photographs.ListIndex - (mintCurrentPage - 1) * goSystemConfig.ImagesPerPage).Value = 1
                Else
                    chkSelected(lis_photographs.ListIndex - (mintCurrentPage - 1) * goSystemConfig.ImagesPerPage).Value = 0
                End If
            End If
        Else
            '
            '   Make sure the digital image is selected.
            '
            moDigitalImages(CStr(lis_photographs.ListIndex)).Selected = lis_photographs.Selected(lis_photographs.ListIndex)
            UpdateStatusBar
        End If
    End If
    '
    '   To get round the problem of this code being invoked again
    '   when the item is deselected, an item is checked to see which
    '   of the two invocations it is.
    '
    If Not first_time Then
        first_time = True
        Exit Sub
    End If
    '
    If mblnUpdateScreenOnly Then
        Exit Sub
    End If
    '
    '   This is the part that does the work with the underlying objects.
    '
    
    blnInStock = False
    Set oDigitalImage = moDigitalImages.Item(lis_photographs.ListIndex + 1)
    If lis_photographs.Selected(lis_photographs.ListIndex) Then
        '
        '   If Search Photo Type is Transparency then check if transparencies in stock.
        '
        If moSearch.PhotoType = "T" Then
            If oDigitalImage.Pending = pdgClean And Not photograph.photos_in_stock(lis_photographs.ItemData(lis_photographs.ListIndex)) Then
                '
                '   If no transparencies in stock, ask user if a digital image is required.
                '
                If oDigitalImage.FileLocation("A4") = "" Then
                    '
                    '   Simply inform user there are no transparencies.
                    '
                    MsgBox "There are no transparencies in stock.", vbExclamation
                Else
                    '
                    If MsgBox("There are no transparencies in stock." & vbCrLf & "Do you wish to send a digital image instead?", vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
                        '
                        '   Check if an A4 digital image exists.
                        '
                        If oDigitalImage.FileLocation("A4") = "" Then
                            MsgBox "There is no 'A4' digital image available for this photograph.", vbExclamation
                        Else
                            '
                            '   Update the pending update flag.
                            '
                            oDigitalImage.PhotoTypeRequired = "D"
                            If oDigitalImage.Pending = pdgClean Or oDigitalImage.Pending = pdgDelete Then
                                If oDigitalImage.SearchResultKey = 0 Then
                                    oDigitalImage.Pending = pdgCreate
                                Else
                                    oDigitalImage.Pending = pdgUpdate
                                End If
                            End If
                            blnInStock = True
                        End If
                    End If
                End If
            Else
                '
                '   Update the pending update flag.
                '
                With oDigitalImage
                    If .SearchResultKey = 0 Then
                        If .Pending = pdgClean Then
                            .Pending = pdgCreate
                        End If
                    Else
                        If .Pending = pdgClean Then
                            .Pending = pdgUpdate
                        ElseIf .Pending = pdgDelete Then
                            .Pending = pdgClean
                        End If
                    End If
                End With
                blnInStock = True
            End If
        Else
            '
            '   Is the photograph already on a Digital Search?
            '
            blnValid = True
            For Each oSearchResult In moSearch.SearchResults
                If oSearchResult.BatchNo = oDigitalImage.BatchNo Then
                    MsgBox "Digital Photograph '" & CStr(oDigitalImage.BatchNo) & "' has already been selected for this Search.", vbExclamation
                    blnValid = False
                End If
            Next oSearchResult
            If blnValid Then
                '
                '   Check if an A4 digital image exists.
                '
                If oDigitalImage.FileLocation("A4") = "" Then
                    MsgBox "There is no 'A4' digital image available for this photograph.", vbExclamation
                Else
                    '
                    '
                    '   Update the pending update flag.
                    '
                    With oDigitalImage
                        If .SearchResultKey = 0 Then
                            If .Pending = pdgClean Then
                                .Pending = pdgCreate
                            End If
                        Else
                            If .Pending = pdgClean Then
                                .Pending = pdgUpdate
                            ElseIf .Pending = pdgDelete Then
                                .Pending = pdgClean
                            End If
                        End If
                    End With
                    blnInStock = True
                End If
            End If
        End If
        If Not blnInStock Then
            '
            '   Just want the screen to be updated not the Digital Image object etc.
            '
            mblnUpdateScreenOnly = True
            lis_photographs.Selected(lis_photographs.ListIndex) = False
            mblnUpdateScreenOnly = False
            '
            '   Don't deselect image if it is not currently displayed.
            '
            If lis_photographs.ListIndex <= (mintCurrentPage * goSystemConfig.ImagesPerPage) - 1 And _
               lis_photographs.ListIndex >= (mintCurrentPage - 1) * goSystemConfig.ImagesPerPage Then
                chkSelected(lis_photographs.ListIndex - (mintCurrentPage - 1) * goSystemConfig.ImagesPerPage).Value = 0
            End If
        End If
    Else
        With oDigitalImage
            Select Case .Pending
            Case Is = pdgClean, pdgDelete
                .Pending = pdgDelete
            Case Is = pdgCreate, pdgUpdate
                .Pending = pdgClean
            End Select
        End With
    End If
'
'********* Code Ends Here **************
'
   Exit Sub
'
lis_photographs_Click_Error:
    DisplayError , "frm_search_photograph.lis_photographs_Click", vbExclamation
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
'***************************************
' Module/Form Name   : frm_search_photograph
'
' Procedure Name     : cmdFind_Click
'
' Purpose            :
'
' Date Created       : 27/10/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 27/10/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo cmdFind_Click_Error
'
'******** Code Starts Here *************
'

    If Len(LTrim(txt_description)) = 0 Then
        MsgBox "Enter a string to search for", vbExclamation
        txt_description.SetFocus
        Exit Sub
    End If
    
    cmdFind.Enabled = False
    Me.Caption = "Find photographs containing '" + txt_description + "'"
    
    photograph.search txt_description.Text, snap_batch, dbOpenSnapshot, abort, optKeywords.Value
    If abort Then Exit Sub
    '
    DisplayPhotographs
'
'********* Code Ends Here **************
'
    Exit Sub
    '
cmdFind_Click_Error:
    cmdFind.Enabled = True
    DisplayError , "frm_search_photograph.cmdFind_Click", vbExclamation
End Sub

Private Sub DisplayPhotographs(Optional pblnNewPhotographs As Boolean = True)
'***************************************
' Module/Form Name   : frm_search_photograph
'
' Procedure Name     : DisplayPhotographs
'
' Purpose            :
'
' Date Created       : 23/06/2004
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo DisplayPhotographs_Error
'
'******** Code Starts Here *************
'
    Dim doevents_count          As Integer
    Dim intPhotographs          As Integer
    Dim oDigitalImage           As DigitalImage
    Dim blnBeginningOfData      As Boolean
    Dim intIdx                  As Integer
    Dim lngBatchNo              As Long
    Dim strDescription          As String
    Dim strCredit               As String
    Dim varDeDupe               As Variant
    Dim blnExclusive            As Boolean
    
    Screen.MousePointer = vbHourglass
    tim_Tooltip.Enabled = False
    '
    '   Put returned batches into the listbox.
    '
    If pblnNewPhotographs Then
        If snap_batch.RecordCount = 0 Then
            MsgBox "No photographs found", vbExclamation, msg_title
            Screen.MousePointer = vbDefault
            cmdFind.Enabled = True
            Exit Sub
        End If
    End If
    
    If pblnNewPhotographs Then snap_batch.MoveLast
    '
    '   Give a warning if too many images have been returned.
    '
    If pblnNewPhotographs Then
        If snap_batch.RecordCount + moDigitalImages.Count > goSystemConfig.ImagesReturned Then
            If goSystemConfig.ImagesReturned > moDigitalImages.Count Then
                MsgBox Format(snap_batch.RecordCount, "###,###,##0") & " photographs have been returned by your query." & vbCrLf & _
                       "Only the first " & goSystemConfig.ImagesReturned - moDigitalImages.Count & " will be shown.", vbInformation
            Else
                MsgBox "The maximum number of records have already been shown. To change the maximum " & vbCrLf _
                     & "go to Tools & Options and select the 'Digital Images' tab.", vbInformation
                Screen.MousePointer = vbDefault
                cmdFind.Enabled = True
                Exit Sub
            End If
        End If
    End If
    
    intPhotographs = moDigitalImages.Count
    '
    '   Calculate the number of pages to be shown. Use 'Fix' to round up.
    '
    If pblnNewPhotographs Then
        If snap_batch.RecordCount + moDigitalImages.Count > goSystemConfig.ImagesReturned Then
            mintNoOfPages = Int(((goSystemConfig.ImagesReturned) / goSystemConfig.ImagesPerPage) * -1) * -1
        Else
            mintNoOfPages = Int(((moDigitalImages.Count + snap_batch.RecordCount) / goSystemConfig.ImagesPerPage) * -1) * -1
        End If
    Else
        mintNoOfPages = Int(((moSearch.SearchResults.Count) / goSystemConfig.ImagesPerPage) * -1) * -1
    End If
    '
    blnBeginningOfData = False
    If Not pblnNewPhotographs Then
        intIdx = moSearch.SearchResults.Count
        If intIdx = 0 Then
            blnBeginningOfData = True
        End If
    End If
    While Not blnBeginningOfData And intPhotographs < goSystemConfig.ImagesReturned
        If pblnNewPhotographs Then
            lngBatchNo = snap_batch!batch_no
            strDescription = snap_batch!Description
            strCredit = snap_batch!Credit
            blnExclusive = snap_batch!Exclusive
        Else
            lngBatchNo = moSearch.SearchResults.Item(intIdx).BatchNo
            '
            strDescription = moSearch.SearchResults.Item(intIdx).PhotoDesc
            strCredit = moSearch.SearchResults.Item(intIdx).Credit
            blnExclusive = moSearch.SearchResults.Item(intIdx).Exclusive
        End If
        '
        '   Deduplicate the Search Results collection first.
        '
        varDeDupe = lngBatchNo
        On Error Resume Next
        mcolDeDupe.Add varDeDupe, CStr(lngBatchNo)
        If Err.Number <> 0 Then
            On Error GoTo DisplayPhotographs_Error
            GoTo PreviousSearchResult
        End If
        On Error GoTo DisplayPhotographs_Error
        '
        lis_photographs.AddItem format_photograph(lngBatchNo, strDescription)
        lis_photographs.ItemData(lis_photographs.NewIndex) = lngBatchNo
        '
        '   Add the Digital Image to the module level collection.
        '
        Set oDigitalImage = moDigitalImages.Add(lngBatchNo, _
                                                strDescription, _
                                                False, _
                                                mintPicturesDisplayed, _
                                                CStr(mintPicturesDisplayed))
        oDigitalImage.Credit = strCredit
        oDigitalImage.Exclusive = blnExclusive
        If Not pblnNewPhotographs Then
            oDigitalImage.Selected = True
            oDigitalImage.SearchResultKey = moSearch.SearchResults.Item(intIdx).SearchResultKey
        End If
        '
        '   What type of Photo does the Digital Image represent physically?
        '
        oDigitalImage.PhotoTypeRequired = moSearch.PhotoType
        '
        '   Update the counts.
        '
        mintPicturesDisplayed = mintPicturesDisplayed + 1
        If oDigitalImage.ImageAvailable Then
            mintPictureImagesCount = mintPictureImagesCount + 1
        Else
            mintPictureTextCount = mintPictureTextCount + 1
        End If
        '
        '   Add Image to Images Tab.
        '
        If moDigitalImages.Count <= (mintCurrentPage * goSystemConfig.ImagesPerPage) And _
           moDigitalImages.Count > ((mintCurrentPage - 1) * goSystemConfig.ImagesPerPage) Then
            DisplayDigitalImage oDigitalImage
        End If
        
PreviousSearchResult:
        If pblnNewPhotographs Then
            snap_batch.MovePrevious
            If snap_batch.BOF Then
                blnBeginningOfData = True
            End If
        Else
            intIdx = intIdx - 1
            If intIdx = 0 Then
                blnBeginningOfData = True
            End If
        End If
        intPhotographs = intPhotographs + 1
    Wend
    '
    '   Enable or otherwise the Next & Previous buttons.
    '
    EnableNextPrevious
    '
    ConfigureScrollBar
    UpdateStatusBar
    '
    If pblnNewPhotographs Then
        '
        '   If text tab selected then fix controls so that the images will display!
        '
        If sstFind.Tab = 0 Then
            FixBadControls
            UnFixBadControls
        End If
        '
        cmdSelect.Default = True
        chkSelected(0).SetFocus
    End If
    '
    If pblnNewPhotographs Then
        snap_batch.Close
        Set snap_batch = Nothing
    End If
    tim_Tooltip.Enabled = True
    If mSearchRefine <> srreview Then
        cmdFind.Enabled = True
    End If
    Screen.MousePointer = vbDefault
'
'********* Code Ends Here **************
'
   Exit Sub
'
DisplayPhotographs_Error:
    ErrorRaise "frm_search_photograph.DisplayPhotographs"
End Sub

Private Sub cmdSelect_Click()
'***************************************
' Module/Form Name   : frm_search_photograph
'
' Procedure Name     : cmdSelect_Click
'
' Purpose            :
'
' Date Created       : 15/02/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo cmdSelect_Click_Error
'
'******** Code Starts Here *************
'
    Dim i                       As Long
    Dim oDigitalImage           As DigitalImage
    Dim oDigitalSearch          As Search2
    Dim oSearchToUpdate         As Search2
    Dim blnPendingUpdatesFound  As Boolean
    
    '
    '   Stop the double selection of images!
    '
    cmdSelect.Enabled = False
    cmdFind.Enabled = False
    cmdClear.Enabled = False
    
    Screen.MousePointer = vbHourglass
    LockWindow Me.hWnd
''    For i = 0 To lis_photographs.ListCount - 1
''        If lis_photographs.Selected(i) Then
    
    blnPendingUpdatesFound = False
    For Each oDigitalImage In moDigitalImages
        '
        '   Is there a pending update required.
        '
        If oDigitalImage.Pending <> pdgClean Then
            '
            '   Is a different type of Search required for the Photograph.
            '
            If moSearch.PhotoType <> oDigitalImage.PhotoTypeRequired Then
                '
                '   Search result needs to go on a different Search.
                '
                If moSearch.PhotoType = "T" Then
                    '
                    '   Has a Digital Search already been created for this Search?
                    '
                    Set oDigitalSearch = New Search2
                    begin_trans
                        '
                        '   Create a new Search to hold the digital images.
                        '
                        With oDigitalSearch
                            .PhotoType = "D"
                            On Error Resume Next
                            .Add moSearch.JobNo, moSearch.Description, moSearch.SearchDate
                            '
                            If Err.Number = vbObjectError + 1 Then
                                '
                                '   Search already exists, so use this one.
                                '
                                On Error GoTo cmdSelect_Click_Error
                                .Read plngJobNo:=moSearch.JobNo, _
                                      pstrDescription:=moSearch.Description, _
                                      pstrPhotoType:="D"
                            ElseIf Err.Number <> 0 Then
                                ErrorSave
                                On Error GoTo cmdSelect_Click_Error
                                ErrorRestore
                            Else
                                On Error GoTo cmdSelect_Click_Error
                            End If
                        End With
                        '
                        '   Create a Refine activity.
                        '
                        CreateDigitialActivity oDigitalSearch, "REFS"
                        moSearch.DigitalSearchNo = oDigitalSearch.SearchNo
                        moSearch.update
                    commit_trans
                    Set oSearchToUpdate = oDigitalSearch
                    '
                Else
                    '
                    '   Something similar may need to go in here for when a digital image isn't available but a Transparency is.
                    '
                    MsgBox "The scenario whereby a transparency is required on a Digital Search is not currently catered for.", vbExclamation
                End If
                '
                '   Raise an event for the calling form to process the Search Result creation.
                '
            Else
                Set oSearchToUpdate = moSearch
            End If
''                RaiseEvent PhotographSelected(lis_photographs.ItemData(i), txt_description.Text)
            'frm_search.add_photographs lis_photographs.ItemData(i), False
            RaiseEvent PhotographSelected(oSearchToUpdate, oDigitalImage)
            blnPendingUpdatesFound = True
        End If
NextDigitalImage:
    Next oDigitalImage
    '
    If blnPendingUpdatesFound Then
        RaiseEvent PhotographSelectionComplete
    End If
    '
    UnlockWindow
    Unload Me
    Set frm_search_photograph = Nothing
    Screen.MousePointer = vbDefault
'
'********* Code Ends Here **************
'
    Exit Sub
    '
cmdSelect_Click_Error:
    DisplayError , "frm_search_photograph.cmdSelect_Click", vbExclamation
End Sub

Private Sub CreateDigitialActivity(ByRef poSearch As Search2, _
                                   ByVal pstrActivityType As String)
'***************************************
' Module/Form Name   : frm_search_photograph
'
' Procedure Name     : CreateDigitialActivity
'
' Purpose            :
'
' Date Created       : 01/07/2004
'
' Author             : ADMINISTRATOR
'
' Parameters         : poSearch - Search2
'                    : pstrActivityType - String
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo CreateDigitialActivity_Error
'
'******** Code Starts Here *************
'
    Dim oActivity       As New Activity
    
    If poSearch.PhotoType = "D" Then
        '
        '   Does an activity already exist?
        '
        poSearch.Activities.ActivityType = pstrActivityType
        poSearch.Activities.Refresh
        If poSearch.Activities.Count = 0 Then
            With oActivity
                .ActivityType = pstrActivityType
                .StartDate = Now()
                .SearchNo = poSearch.SearchNo
                .CustomerNo = moJob.CustomerNo
                .JobNo = moJob.JobNo
                .Status = "0"
                .Create
            End With
            moSearch.Activities.Refresh
        End If
    End If
'
'********* Code Ends Here **************
'
   Exit Sub
'
CreateDigitialActivity_Error:
    ErrorRaise "frm_search_photograph.CreateDigitialActivity"
End Sub

Private Sub cmdClear_Click()
'***************************************
' Module/Form Name   : frm_search_photograph
'
' Procedure Name     : cmdClear_Click
'
' Purpose            :
'
' Date Created       : 27/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo cmdClear_Click_Error
'
'******** Code Starts Here *************
'
    If mSearchRefine = srRefine And moSearch.SearchResults.Count <> 0 Then
        If MsgBox("Note that this will remove all Photographs from the current Search?" & vbCrLf & _
                   "Do you wish to continue?", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
            Exit Sub
        End If
    End If
    '
    Screen.MousePointer = vbHourglass
    ClearDisplay
    mintCurrentPage = 1
    mintNoOfPages = 1
    SetUpPage 0
    HighLightText txt_description
    txt_description.SetFocus
    mintPictureImagesCount = 0
    mintPictureTextCount = 0
    mintPicturesDisplayed = 0
    ConfigureScrollBar
    UpdateStatusBar
    Set mcolDeDupe = Nothing
    Set mcolDeDupe = New Collection
    Screen.MousePointer = vbDefault
'
'********* Code Ends Here **************
'
   Exit Sub
'
cmdClear_Click_Error:
    DisplayError , "frm_search_photograph.cmdClear_Click", vbExclamation
End Sub

Private Sub cmdHelp_Click()
'***************************************
' Module/Form Name   : frm_search_photograph
'
' Procedure Name     : cmdHelp_Click
'
' Purpose            :
'
' Date Created       : 27/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo cmdHelp_Click_Error
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
        strKey = "Searching"
        nRet = OSWinHelp(Me.hWnd, App.HelpFile, HELP_KEY, strKey)
        If Err Then
            MsgBox Err.Description
        End If
    End If

'
'********* Code Ends Here **************
'
   Exit Sub
'
cmdHelp_Click_Error:
    DisplayError , "frm_search_photograph.cmdHelp_Click", vbExclamation
End Sub

Private Sub DisplayDigitalImage(oDigitalImage As DigitalImage)
'***************************************
' Module/Form Name   : frm_search_photograph
'
' Procedure Name     : DisplayDigitalImage
'
' Purpose            :
'
' Date Created       : 27/12/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : oDigitalImage - DigitalImage
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo DisplayDigitalImage_Error
'
'******** Code Starts Here *************
'
    '
    '   Update the counts.
    '
    mintPageDisplayed = mintPageDisplayed + 1
    If oDigitalImage.ImageAvailable Then
       mintPageImagesCount = mintPageImagesCount + 1
    Else
        mintPageTextCount = mintPageTextCount + 1
    End If
    '
    If oDigitalImage.ImageAvailable Then
        oDigitalImage.ImageControlIndex = mintPageImagesCount - 1
        If mintPageImagesCount <> 1 Then
            Load imgFind(mintPageImagesCount - 1)
            imgFind(mintPageImagesCount - 1).Picture = LoadPicture
        End If
        If mintPageDisplayed <> 1 Then
            Load chkSelected(mintPageDisplayed - 1)
            Load shpSelected(mintPageDisplayed - 1)
            Load linExclusive1(mintPageDisplayed - 1)
            Load linExclusive2(mintPageDisplayed - 1)
            Load cmdScreenRes(mintPageDisplayed - 1)
            cmdScreenRes(mintPageDisplayed - 1).ZOrder
            linExclusive1(mintPageDisplayed - 1).ZOrder
            linExclusive2(mintPageDisplayed - 1).ZOrder
            Load imgA4(mintPageDisplayed - 1)
            imgA4(mintPageDisplayed - 1).ZOrder
            mblnSelectText = False
            chkSelected(mintPageDisplayed - 1).Value = IIf(oDigitalImage.Selected, 1, 0)
            mblnSelectText = True
        End If
        '
        '   Load jpeg file into image box.
        '
        With imgFind(mintPageImagesCount - 1)
            .Picture = LoadPicture(oDigitalImage.FileLocation)
            .Tag = oDigitalImage.DigitalImageIndex
        End With
        '
        '   Is the Image Exclusive?
        '
        If oDigitalImage.Exclusive Then
            linExclusive1(mintPageDisplayed - 1).Visible = True
            linExclusive2(mintPageDisplayed - 1).Visible = True
        Else
            linExclusive1(mintPageDisplayed - 1).Visible = False
            linExclusive2(mintPageDisplayed - 1).Visible = False
        End If
        '
        '   Position Image Box.
        '
        PositionImageControl imgFind(mintPageImagesCount - 1), _
                             chkSelected(mintPageDisplayed - 1), _
                             shpSelected(mintPageDisplayed - 1), _
                             linExclusive1(mintPageDisplayed - 1), _
                             linExclusive2(mintPageDisplayed - 1), _
                             cmdScreenRes(mintPageDisplayed - 1), _
                             imgA4(mintPageDisplayed - 1), _
                             oDigitalImage
    Else
        oDigitalImage.TextControlIndex = mintPageTextCount - 1
        If mintPageTextCount <> 1 Then
            Load lblFind(mintPageTextCount - 1)
            lblFind(mintPageTextCount - 1).Caption = ""
        End If
        If mintPageDisplayed <> 1 Then
            Load chkSelected(mintPageDisplayed - 1)
            Load shpSelected(mintPageDisplayed - 1)
            Load linExclusive1(mintPageDisplayed - 1)
            Load linExclusive2(mintPageDisplayed - 1)
            Load cmdScreenRes(mintPageDisplayed - 1)
            cmdScreenRes(mintPageDisplayed - 1).ZOrder
            linExclusive1(mintPageDisplayed - 1).ZOrder
            linExclusive2(mintPageDisplayed - 1).ZOrder
            Load imgA4(mintPageDisplayed - 1)
            imgA4(mintPageDisplayed - 1).ZOrder
            mblnSelectText = False
            chkSelected(mintPageDisplayed - 1).Value = IIf(oDigitalImage.Selected, 1, 0)
            mblnSelectText = True
        End If
        '
        '   Load text into text box.
        '
        With lblFind(mintPageTextCount - 1)
            .Caption = oDigitalImage.Description
            .Tag = oDigitalImage.DigitalImageIndex
        End With
        '
        '   Is the Image Exclusive?
        '
        If oDigitalImage.Exclusive Then
            linExclusive1(mintPageDisplayed - 1).Visible = True
            linExclusive2(mintPageDisplayed - 1).Visible = True
        Else
            linExclusive1(mintPageDisplayed - 1).Visible = False
            linExclusive2(mintPageDisplayed - 1).Visible = False
        End If
        '
        '   Position text box.
        '
        PositionImageControl lblFind(mintPageTextCount - 1), _
                             chkSelected(mintPageDisplayed - 1), _
                             shpSelected(mintPageDisplayed - 1), _
                             linExclusive1(mintPageDisplayed - 1), _
                             linExclusive2(mintPageDisplayed - 1), _
                             cmdScreenRes(mintPageDisplayed - 1), _
                             imgA4(mintPageDisplayed - 1), _
                             oDigitalImage
    End If
    '
    '   Show whether image is selected or otherwise.
    '
    mblnSelectText = False
    If oDigitalImage.Selected Then
        chkSelected(oDigitalImage.DigitalImageIndex - (mintCurrentPage - 1) * goSystemConfig.ImagesPerPage).Value = 1
        ToggleDigitalImageSelection oDigitalImage.DigitalImageIndex - (mintCurrentPage - 1) * goSystemConfig.ImagesPerPage, False
    Else
        chkSelected(oDigitalImage.DigitalImageIndex - (mintCurrentPage - 1) * goSystemConfig.ImagesPerPage).Value = 0
        ToggleDigitalImageSelection oDigitalImage.DigitalImageIndex - (mintCurrentPage - 1) * goSystemConfig.ImagesPerPage, False
    End If
    mblnSelectText = True
'
'********* Code Ends Here **************
'
   Exit Sub
'
DisplayDigitalImage_Error:
    ErrorRaise "frm_search_photograph.DisplayDigitalImage"
End Sub

Private Sub PositionImageControl(ctlImage As Control, _
                                 ctlCheck As CheckBox, _
                                 shpSelected As Shape, _
                                 linExclusive1 As Line, _
                                 linExclusive2 As Line, _
                                 ctlScreenRes As CommandButton, _
                                 ctlA4 As Image, _
                                 oDigitalImage As DigitalImage)
'***************************************
' Module/Form Name   : frm_search_photograph
'
' Procedure Name     : PositionImageControl
'
' Purpose            :
'
' Date Created       : 27/12/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : ctlImage - Control
'                    : ctlCheck - CheckBox
'                    : shpSelected - Shape
'                    : ctlA4 - Image
'                    : oDigitalImage - DigitalImage
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 05/07/2005 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo PositionImageControl_Error
'
'******** Code Starts Here *************
'
    '
    '   If image, calculate the dimensions.
    '
    With oDigitalImage
        .MaxHeight = mintImageHeight
        .MaxWidth = mintImageWidth
        If TypeOf ctlImage Is Image Then
            Set .Picture = ctlImage.Picture
        End If
    End With
    '
    '   Now position the image.
    '
    ctlImage.Move (mintImageWidth + 100) * ((oDigitalImage.DigitalImageIndex - (mintCurrentPage - 1) * goSystemConfig.ImagesPerPage) Mod mintMaxImagesPerRow) + 100, _
                  CONST_IMAGE_TOP + ((mintImageHeight + ctlCheck.Height + 200) * (Int((oDigitalImage.DigitalImageIndex - (mintCurrentPage - 1) * goSystemConfig.ImagesPerPage) / mintMaxImagesPerRow))), _
                  oDigitalImage.Width, _
                  oDigitalImage.Height
    ctlImage.Visible = True
    '
    '   Position Screen Res command button.
    '
    With ctlScreenRes
        .Move ctlImage.Left + mintImageWidth - ctlScreenRes.Width, _
              ctlImage.Top + mintImageHeight + 100
        .Visible = IIf(oDigitalImage.FileLocation("SC") = "", False, True)
    End With
    '
    '   Position High Res Indicator Image.
    '
    With ctlA4
        .Move ctlScreenRes.Left - 100 - ctlA4.Width, _
              ctlScreenRes.Top
        .Visible = IIf(oDigitalImage.FileLocation("A4") = "", False, True)
        
    End With
    '
    '   Position checkbox.
    '
    ctlCheck.Move ctlImage.Left, _
                  ctlImage.Top + mintImageHeight + 100, _
                  ctlA4.Left - ctlImage.Left, _
                  ctlCheck.Height
    ctlCheck.Caption = CStr(oDigitalImage.BatchNo)
    ctlCheck.Visible = True
    '
    '   Position selection shape.
    '
    With shpSelected
        .Move ctlImage.Left - 30, _
              ctlImage.Top - 30, _
              ctlImage.Width + 80, _
              ctlImage.Height + 90
    End With
    '
    '   Position Exclusive Lines.
    '
    With linExclusive1
        .X1 = ctlImage.Left
        .X2 = ctlImage.Left + ctlImage.Width
        .Y1 = ctlImage.Top
        .Y2 = ctlImage.Top + ctlImage.Height
    End With
    With linExclusive2
        .X1 = ctlImage.Left + ctlImage.Width
        .X2 = ctlImage.Left
        .Y1 = ctlImage.Top
        .Y2 = ctlImage.Top + ctlImage.Height
    End With
'
'********* Code Ends Here **************
'
   Exit Sub
'
PositionImageControl_Error:
    ErrorRaise "frm_search_photograph.PositionImageControl"
End Sub

Private Sub lblfind_Click(Index As Integer)
'***************************************
' Module/Form Name   : frm_search_photograph
'
' Procedure Name     : lblfind_Click
'
' Purpose            :
'
' Date Created       : 27/12/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : Index - Integer
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo lblfind_Click_Error
'
'******** Code Starts Here *************
'
    blnThumbClicked = True
    If chkSelected(CInt(lblFind(Index).Tag - (mintCurrentPage - 1) * goSystemConfig.ImagesPerPage)).Value = 0 Then
        chkSelected(CInt(lblFind(Index).Tag - (mintCurrentPage - 1) * goSystemConfig.ImagesPerPage)).Value = 1
    Else
        chkSelected(CInt(lblFind(Index).Tag - (mintCurrentPage - 1) * goSystemConfig.ImagesPerPage)).Value = 0
    End If
'
'********* Code Ends Here **************
'
   Exit Sub
'
lblfind_Click_Error:
    DisplayError , "frm_search_photograph.lblfind_Click", vbExclamation
End Sub

Private Sub tim_Tooltip_Timer()
'***************************************
' Module/Form Name   : frm_search_photograph
'
' Procedure Name     : tim_Tooltip_Timer
'
' Purpose            :
'
' Date Created       : 27/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo tim_Tooltip_Timer_Error
'
'******** Code Starts Here *************
'
    Dim current_Image As Long
    Static previous_Image As Long
    Dim cursor_location As POINTAPI
    Dim blnImage As Boolean

    ' Get the current position of the cursor.
    Call GetCursorPos(cursor_location)

    ' Determine the hWnd for the control under the cursor.
    current_Image = DetermineImageFromPoint(cursor_location.X, cursor_location.Y, blnImage)

    ' We have two flags, current_ & previous, which monitor changes in the current hWnd.  Else display tooltip if requested.
    If (current_Image = previous_Image) Then
        If current_Image = -1 And lngTimeTipHidden <> 0 Then
            Hide_ToolTip
        Else
            Exit Sub
        End If
    Else
        previous_Image = current_Image
    End If
    ' Identify the control currently under the cursor.
    If current_Image = -1 Then
        Hide_ToolTip
    Else
        If blnImage Then
            Display_ToolTip imgFind(moDigitalImages(CStr(current_Image)).ImageControlIndex)
        Else
            Display_ToolTip lblFind(moDigitalImages(CStr(current_Image)).TextControlIndex)
        End If
    End If
'
'********* Code Ends Here **************
'
   Exit Sub
'
tim_Tooltip_Timer_Error:
    DisplayError , "frm_search_photograph.tim_Tooltip_Timer", vbExclamation
End Sub

Private Sub txt_description_GotFocus()
    cmdFind.Default = True
End Sub

Private Sub vscFind_Change()
'***************************************
' Module/Form Name   : frm_search_photograph
'
' Procedure Name     : vscFind_Change
'
' Purpose            :
'
' Date Created       : 27/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo vscFind_Change_Error
'
'******** Code Starts Here *************
'
    ScrollImages
'
'********* Code Ends Here **************
'
   Exit Sub
'
vscFind_Change_Error:
    DisplayError , "frm_search_photograph.vscFind_Change", vbExclamation
End Sub

Private Sub ConfigureScrollBar()
'***************************************
' Module/Form Name   : frm_search_photograph
'
' Procedure Name     : ConfigureScrollBar
'
' Purpose            :
'
' Date Created       : 27/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo ConfigureScrollBar_Error
'
'******** Code Starts Here *************
'
    Dim dblRatio As Double

    '
    '   Determine the current ratio.
    '
    dblRatio = vscFind.Value / vscFind.Max
    '
    '   Calculate the number of rows of Images in the tab.
    '
    'mdblRowsOfImages = Int((IIf(mintPicturesDisplayed < goSystemConfig.ImagesPerPage, mintPicturesDisplayed, goSystemConfig.ImagesPerPage) / mintMaxImagesPerRow) * -1) * -1
    mdblRowsOfImages = Int((IIf(mintPageDisplayed < goSystemConfig.ImagesPerPage, mintPageDisplayed, goSystemConfig.ImagesPerPage) / mintMaxImagesPerRow) * -1) * -1
    '
    '   Resize the large picture.
    '
    With picFindLarge
        .Left = picFind.Left - 100
        .Width = picFind.Width - vscFind.Width - CONST_SCROLL_BAR_LEFT_OFFSET
        .Height = mdblRowsOfImages * (mintImageHeight + chkSelected(0).Height + 200) + CONST_IMAGE_TOP
        '.Height = picFind.Height + 2 * (mintImageHeight + chkSelected(0).Height + 200)
        mLogicalImagesHeight = mdblRowsOfImages * (mintImageHeight + chkSelected(0).Height + 200) + CONST_IMAGE_TOP
    End With
    '
    '   Calculate the number of Images visibly displayed.
    '
    mdblRowsDisplayed = (picFind.Height - CONST_IMAGE_TOP) / (mintImageHeight + chkSelected(0).Height + 200)
    With vscFind
        .Left = picFind.Width - .Width - CONST_SCROLL_BAR_LEFT_OFFSET
        .Top = 0
        .Height = picFind.Height - CONST_SCROLL_BAR_HEIGHT_OFFSET
        If (mdblRowsOfImages - mdblRowsDisplayed) < 0 Then
            .Enabled = False
            vscFind.Value = 0
        Else
            .Enabled = True
            .SmallChange = 2
            .Max = mdblRowsOfImages * 100
            If mdblRowsDisplayed < 2 Then
                .LargeChange = 100
            Else
                .LargeChange = (mdblRowsDisplayed - 1) * 100
            End If
            .Min = 0
            '
            '   Apply the old ratio.
            '
            vscFind.Value = vscFind.Max * dblRatio
        End If
    End With
'
'********* Code Ends Here **************
'
   Exit Sub
'
ConfigureScrollBar_Error:
    ErrorRaise "frm_search_photograph.ConfigureScrollBar"
End Sub

Private Sub ToggleDigitalImageSelection(intIndex As Integer, Optional blnUpdateStatusBar As Boolean = True)
'***************************************
' Module/Form Name   : frm_search_photograph
'
' Procedure Name     : ToggleDigitalImageSelection
'
' Purpose            :
'
' Date Created       : 27/12/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : intIndex - Integer
'                    : blnUpdateStatusBar - Boolean
'                    : = -
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo ToggleDigitalImageSelection_Error
'
'******** Code Starts Here *************
'
    shpSelected(moDigitalImages(CStr(intIndex)).DigitalImageIndex).Visible = CBool(chkSelected(intIndex))
    moDigitalImages(CStr(intIndex + (mintCurrentPage - 1) * goSystemConfig.ImagesPerPage)).Selected = CBool(chkSelected(intIndex))
    If blnUpdateStatusBar Then
        UpdateStatusBar
    End If
'
'********* Code Ends Here **************
'
   Exit Sub
'
ToggleDigitalImageSelection_Error:
    ErrorRaise "frm_search_photograph.ToggleDigitalImageSelection"
End Sub

Private Sub UpdateStatusBar()
'***************************************
' Module/Form Name   : frm_search_photograph
'
' Procedure Name     : UpdateStatusBar
'
' Purpose            :
'
' Date Created       : 27/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo UpdateStatusBar_Error
'
'******** Code Starts Here *************
'
    Dim pnlX As Panel
    Dim oDigitalImage As DigitalImage
    Dim intSelected As Integer
    Dim intTotal As Integer

    intSelected = 0
    intTotal = 0
    For Each oDigitalImage In moDigitalImages
        If oDigitalImage.Selected Then
            intSelected = intSelected + 1
        End If
        intTotal = intTotal + 1
    Next oDigitalImage
    Set pnlX = stbSearchFind.Panels.Item(1)
    pnlX.Text = "Selected: " & CStr(intSelected)

    Set pnlX = stbSearchFind.Panels.Item(2)
    pnlX.Text = "Position: " & CStr(lis_photographs.ListIndex + 1)

    Set pnlX = stbSearchFind.Panels.Item(3)
    pnlX.Text = "Total: " + CStr(intTotal)

'
'********* Code Ends Here **************
'
   Exit Sub
'
UpdateStatusBar_Error:
    ErrorRaise "frm_search_photograph.UpdateStatusBar"
End Sub

Sub Display_ToolTip(Ctl As Control)
'***************************************
' Module/Form Name   : frm_search_photograph
'
' Procedure Name     : Display_ToolTip
'
' Purpose            :
'
' Date Created       : 27/12/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : Ctl - Control
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo Display_ToolTip_Error
'
'******** Code Starts Here *************
'
    Dim lngTag As Long
    '
    '   Initialise the label.
    '
    txtToolTip.Visible = False
    DoEvents
    With lblStretch
        .Height = 300
    End With
    '
    'lblStretch.Caption = moDigitalImages(CStr(moDigitalImages(Ctl.Tag).DigitalImageIndex)).Description & " - " & _
                         moDigitalImages(CStr(moDigitalImages(Ctl.Tag).DigitalImageIndex)).Credit
    lblStretch.Caption = moDigitalImages(Ctl.Tag).Description & " - " & _
                         moDigitalImages(Ctl.Tag).Credit
    lngTag = Ctl.Tag
    '
    '   Keep the ratio of the tooltip to 2w x 1h.
    '
    With lblStretch
        .AutoSize = True
        .Width = Sqr(.Width * .Height * 2)
        .AutoSize = True
    End With
    '
    txtToolTip.Height = lblStretch.Height + 100
    txtToolTip.Text = lblStretch.Caption
    txtToolTip.Width = lblStretch.Width
    If Ctl.Top + Ctl.Height + txtToolTip.Height + 80 + picFindLarge.Top > picFind.ScaleHeight Then
        txtToolTip.Top = Ctl.Top - txtToolTip.Height - 80 + picFindLarge.Top
    Else
        txtToolTip.Top = Ctl.Top + Ctl.Height + 80 + picFindLarge.Top
    End If
    If Ctl.Left + txtToolTip.Width < picFindLarge.ScaleWidth Then
        txtToolTip.Left = Ctl.Left
    Else
        txtToolTip.Left = Ctl.Left - txtToolTip.Width + Ctl.Width
    End If
    'txtToolTip.ZOrder 0
    txtToolTip.Visible = True

    ' Set the timer interval to check the mouse location more frequently so that we don't have a lag between when the mouse is
    ' moved and when the label changes or is hidden.
    tim_Tooltip.Interval = Time_When_ToolTip_Displayed
    lngTimeTipHidden = 0
'
'********* Code Ends Here **************
'
   Exit Sub
'
Display_ToolTip_Error:
    ErrorRaise "frm_search_photograph.Display_ToolTip"
End Sub

Sub Hide_ToolTip(Optional ByVal ChangeTimerInterval As Boolean = True)
'***************************************
' Module/Form Name   : frm_search_photograph
'
' Procedure Name     : Hide_ToolTip
'
' Purpose            :
'
' Date Created       : 27/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo Hide_ToolTip_Error
'
'******** Code Starts Here *************
'
    ' Hides the tooltip when the cursor is over an object which does not have a tip to display.
    txtToolTip.Visible = False
    DoEvents
    '
    '   If tip has been hidden then check to see how long ago it was hidden. If it is less than
    '   'Time_When_ToolTip_Displayed', display the tooltip otherwise, set the timer interval to
    '   'Time_When_ToolTip_Not_Displayed'.
    '
    If lngTimeTipHidden = 0 Then
        lngTimeTipHidden = Timer
    Else
        If (Timer - lngTimeTipHidden) * 1000 > Time_When_ToolTip_Not_Displayed Then
            tim_Tooltip.Interval = Time_When_ToolTip_Not_Displayed
        End If
    End If
'
'********* Code Ends Here **************
'
   Exit Sub
'
Hide_ToolTip_Error:
    ErrorRaise "frm_search_photograph.Hide_ToolTip"
End Sub

Private Function DetermineImageFromPoint(lngX As Long, lngY As Long, blnImage As Boolean) As Long
'***************************************
' Module/Form Name   : frm_search_photograph
'
' Procedure Name     : DetermineImageFromPoint
'
' Purpose            :
'
' Date Created       : 27/12/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : lngX - Long
'                    : lngY - Long
'                    : blnImage - Boolean
'
' Returns            : Long
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo DetermineImageFromPoint_Error
'
'******** Code Starts Here *************
'
    Dim ctlImage As Control
    Dim lngRelX As Long
    Dim lngRelY As Long

    DetermineImageFromPoint = -1
    ConvertPIXELSToTWIPS lngX, lngY
    If lngY < Me.Top + goSystemConfig.TitleBarHeight + sstFind.Top + picFind.Top + 90 Then
        Exit Function
    End If
    If lngY > Me.Top + goSystemConfig.TitleBarHeight + sstFind.Top + picFind.Top + picFind.Height + 50 Then
        Exit Function
    End If
    lngRelX = lngX - (Me.Left + sstFind.Left + picFind.Left + 100)
    lngRelY = lngY - (Me.Top + sstFind.Top + picFind.Top + goSystemConfig.TitleBarHeight + 90)


    For Each ctlImage In Controls
        If ctlImage.Name = "imgFind" Or ctlImage.Name = "lblFind" Then
            If ctlImage.Visible = True Then
                If ctlImage.Left < lngRelX And _
                   ctlImage.Left + ctlImage.Width > lngRelX And _
                   ctlImage.Top + picFindLarge.Top < lngRelY And _
                   ctlImage.Top + picFindLarge.Top + ctlImage.Height > lngRelY Then
                    DetermineImageFromPoint = ctlImage.Tag
                    If ctlImage.Name = "imgFind" Then
                        blnImage = True
                    Else
                        blnImage = False
                    End If
                    Exit Function
                End If
            End If
        End If
    Next ctlImage

'
'********* Code Ends Here **************
'
   Exit Function
'
DetermineImageFromPoint_Error:
    ErrorRaise "frm_search_photograph.DetermineImageFromPoint"
End Function

Private Sub vscFind_Scroll()
'***************************************
' Module/Form Name   : frm_search_photograph
'
' Procedure Name     : vscFind_Scroll
'
' Purpose            :
'
' Date Created       : 27/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo vscFind_Scroll_Error
'
'******** Code Starts Here *************
'
    ScrollImages
'
'********* Code Ends Here **************
'
   Exit Sub
'
vscFind_Scroll_Error:
    DisplayError , "frm_search_photograph.vscFind_Scroll", vbExclamation
End Sub

Private Sub ScrollImages()
'***************************************
' Module/Form Name   : frm_search_photograph
'
' Procedure Name     : ScrollImages
'
' Purpose            :
'
' Date Created       : 27/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo ScrollImages_Error
'
'******** Code Starts Here *************
'
    '
    '   Which row is currently the top.
    '
    picFindLarge.Top = -1 * (vscFind.Value / vscFind.Max) * (picFindLarge.Height - picFind.Height)
    '
    '   Decide which check box to highlight after scrollng.
    '
    vscFind.Refresh
'
'********* Code Ends Here **************
'
   Exit Sub
'
ScrollImages_Error:
    ErrorRaise "frm_search_photograph.ScrollImages"
End Sub

Private Sub FixBadControls()
    Dim Ctl As Control
    On Error Resume Next
    Set mCtlCln = New Collection

    For Each Ctl In Controls
       If Ctl.Left < 0 Then mCtlCln.Add Ctl
    Next
    '
    '   Now add 75000 to all negative left properties.
    '
    For Each Ctl In Controls
       If Ctl.Left < 0 Then Ctl.Left = Ctl.Left + 75000
    Next
End Sub

Private Sub UnFixBadControls()
    Dim Ctl As Control
    '
    '   Put bad left property controls back again.
    '
    On Error Resume Next
    For Each Ctl In mCtlCln
        If Ctl.Left > 0 Then Ctl.Left = Ctl.Left - 75000
    Next
    While mCtlCln.Count <> 0
        mCtlCln.Remove 1
    Wend
    Set mCtlCln = Nothing
End Sub

Private Sub ClearDisplay()
'***************************************
' Module/Form Name   : frm_search_photograph
'
' Procedure Name     : ClearDisplay
'
' Purpose            :
'
' Date Created       : 27/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo ClearDisplay_Error
'
'******** Code Starts Here *************
'
    Set moDigitalImages = Nothing
    Set moDigitalImages = New DigitalImages
    '
    '   Remove all text from lis_photographs
    '
    lis_photographs.Clear
    '
    '   Unload all image controls from memory.
    '
    UnloadImageControls
'
'********* Code Ends Here **************
'
   Exit Sub
'
ClearDisplay_Error:
    ErrorRaise "frm_search_photograph.ClearDisplay"
End Sub

Private Sub UnloadImageControls()
'***************************************
' Module/Form Name   : frm_search_photograph
'
' Procedure Name     : UnloadImageControls
'
' Purpose            :
'
' Date Created       : 27/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo UnloadImageControls_Error
'
'******** Code Starts Here *************
'
    Dim intCount As Integer

    '
    For intCount = 1 To imgFind.Count
        If intCount = 1 Then
            imgFind(intCount - 1).Visible = False
        Else
            Unload imgFind(intCount - 1)
        End If
    Next intCount
    For intCount = 1 To lblFind.Count
        If intCount = 1 Then
            lblFind(intCount - 1).Visible = False
        Else
            Unload lblFind(intCount - 1)
        End If
    Next intCount
    For intCount = 1 To chkSelected.Count
        If intCount = 1 Then
            chkSelected(intCount - 1).Visible = False
        Else
            Unload chkSelected(intCount - 1)
        End If
    Next intCount
    For intCount = 1 To shpSelected.Count
        If intCount = 1 Then
            shpSelected(intCount - 1).Visible = False
        Else
            Unload shpSelected(intCount - 1)
        End If
    Next intCount
    For intCount = 1 To linExclusive1.Count
        If intCount = 1 Then
            linExclusive1(intCount - 1).Visible = False
        Else
            Unload linExclusive1(intCount - 1)
        End If
    Next intCount
    For intCount = 1 To linExclusive2.Count
        If intCount = 1 Then
            linExclusive2(intCount - 1).Visible = False
        Else
            Unload linExclusive2(intCount - 1)
        End If
    Next intCount
    For intCount = 1 To cmdScreenRes.Count
        If intCount = 1 Then
            cmdScreenRes(intCount - 1).Visible = False
        Else
            Unload cmdScreenRes(intCount - 1)
        End If
    Next intCount
    For intCount = 1 To imgA4.Count
        If intCount = 1 Then
            imgA4(intCount - 1).Visible = False
        Else
            Unload imgA4(intCount - 1)
        End If
    Next intCount
    '
    '   Initialise counts.
    '
    mintPageDisplayed = 0
    mintPageImagesCount = 0
    mintPageTextCount = 0

'
'********* Code Ends Here **************
'
   Exit Sub
'
UnloadImageControls_Error:
    ErrorRaise "frm_search_photograph.UnloadImageControls"
End Sub

Private Sub SetUpPage(intIncrement As Integer)
'***************************************
' Module/Form Name   : frm_search_photograph
'
' Procedure Name     : SetUpPage
'
' Purpose            :
'
' Date Created       : 27/12/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : intIncrement - Integer
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo SetUpPage_Error
'
'******** Code Starts Here *************
'
    UnloadImageControls
    picFindLarge.Move 0, 0, picFind.Width - CONST_SCROLL_BAR_LEFT_OFFSET - goSystemConfig.VScrollBarWidth, picFind.Height
    mintCurrentPage = mintCurrentPage + intIncrement
    EnableNextPrevious
'
'********* Code Ends Here **************
'
   Exit Sub
'
SetUpPage_Error:
    ErrorRaise "frm_search_photograph.SetUpPage"
End Sub

Private Sub EnableNextPrevious()
'***************************************
' Module/Form Name   : frm_search_photograph
'
' Procedure Name     : EnableNextPrevious
'
' Purpose            :
'
' Date Created       : 27/12/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo EnableNextPrevious_Error
'
'******** Code Starts Here *************
'
    If mintCurrentPage >= mintNoOfPages Then
        cmdNext.Enabled = False
    Else
        cmdNext.Enabled = True
    End If
    If mintCurrentPage = 1 Then
        cmdPrevious.Enabled = False
    Else
        cmdPrevious.Enabled = True
    End If

    lblCurrentPage.Caption = "Page " & CStr(mintCurrentPage) & " of " & CStr(mintNoOfPages)
'
'********* Code Ends Here **************
'
   Exit Sub
'
EnableNextPrevious_Error:
    ErrorRaise "frm_search_photograph.EnableNextPrevious"
End Sub

Public Sub MouseWheelRolled(ByVal MouseKeys As Long, _
                            ByVal Rotation As Long, _
                            ByVal Xpos As Long, _
                            ByVal Ypos As Long)
    Dim NewValue As Long
    
    With vscFind
        If Rotation > 0 Then
            NewValue = .Value - .LargeChange
            If NewValue < .Min Then
                NewValue = .Min
            End If
        Else
            NewValue = .Value + .LargeChange
            If NewValue > .Max Then
                NewValue = .Max
            End If
        End If
        .Value = NewValue
    End With
    
    ScrollImages
End Sub

