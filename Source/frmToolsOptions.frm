VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmToolsOptions 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   8115
   Icon            =   "frmToolsOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   8115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   5880
      TabIndex        =   76
      Top             =   6540
      Width           =   975
   End
   Begin TabDlg.SSTab sstBusinessTypes 
      Height          =   6480
      Left            =   120
      TabIndex        =   78
      Top             =   0
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   11430
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      TabsPerRow      =   7
      TabHeight       =   520
      TabCaption(0)   =   "Busines Types"
      TabPicture(0)   =   "frmToolsOptions.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "smgBusinessTypes"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdDelete"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdEdit"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdAdd"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Digital Images"
      TabPicture(1)   =   "frmToolsOptions.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblImagesReturned"
      Tab(1).Control(1)=   "blImageHeightWidth"
      Tab(1).Control(2)=   "lblCms"
      Tab(1).Control(3)=   "lblMaxImagesPerPage"
      Tab(1).Control(4)=   "txtImagesReturned"
      Tab(1).Control(5)=   "txtImageHeightWidth"
      Tab(1).Control(6)=   "txtMaxImagesPerPage"
      Tab(1).Control(7)=   "fraImageLocations"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Company Information"
      TabPicture(2)   =   "frmToolsOptions.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblSignatory"
      Tab(2).Control(1)=   "lblAddress"
      Tab(2).Control(2)=   "lblTown"
      Tab(2).Control(3)=   "lblCounty"
      Tab(2).Control(4)=   "lblCountry"
      Tab(2).Control(5)=   "lblPostCode"
      Tab(2).Control(6)=   "lblTelephone"
      Tab(2).Control(7)=   "lblFax"
      Tab(2).Control(8)=   "lblVATNo"
      Tab(2).Control(9)=   "lblEmail"
      Tab(2).Control(10)=   "lblWebSite"
      Tab(2).Control(11)=   "lblInfo2"
      Tab(2).Control(12)=   "lblInfo1"
      Tab(2).Control(13)=   "lblCompanyName"
      Tab(2).Control(14)=   "txtSignatory"
      Tab(2).Control(15)=   "txtAddress1"
      Tab(2).Control(16)=   "txtAddress2"
      Tab(2).Control(17)=   "txtAddress3"
      Tab(2).Control(18)=   "txtTown"
      Tab(2).Control(19)=   "txtCounty"
      Tab(2).Control(20)=   "txtCountry"
      Tab(2).Control(21)=   "txtPostCode"
      Tab(2).Control(22)=   "txtTelNo"
      Tab(2).Control(23)=   "txtFaxNo"
      Tab(2).Control(24)=   "txtVATNo"
      Tab(2).Control(25)=   "txtEmail"
      Tab(2).Control(26)=   "txtWebSite"
      Tab(2).Control(27)=   "txtInfo2"
      Tab(2).Control(28)=   "txtInfo1"
      Tab(2).Control(29)=   "txtCompanyName"
      Tab(2).ControlCount=   30
      TabCaption(3)   =   "Support"
      TabPicture(3)   =   "frmToolsOptions.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblWebSearchTestEmail"
      Tab(3).Control(1)=   "lblUpgradeFrequency"
      Tab(3).Control(2)=   "txtWebSearchTestEmail"
      Tab(3).Control(3)=   "txtUpgradeFrequency"
      Tab(3).Control(4)=   "fraLogFiles"
      Tab(3).Control(5)=   "fraEmails"
      Tab(3).ControlCount=   6
      TabCaption(4)   =   "Data Fixes"
      TabPicture(4)   =   "frmToolsOptions.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cmdFixIncorrectDescriptions"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "FTP Settings"
      TabPicture(5)   =   "frmToolsOptions.frx":0098
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "txtWebSearchesEmailFrom"
      Tab(5).Control(1)=   "txtWebSearchesEmailTo"
      Tab(5).Control(2)=   "txtPostWebAddress"
      Tab(5).Control(3)=   "txtFTPRetypePassword"
      Tab(5).Control(4)=   "txtFTPPassword"
      Tab(5).Control(5)=   "txtFTPUser"
      Tab(5).Control(6)=   "txtFTPServer"
      Tab(5).Control(7)=   "lblWebSearchesEmailFrom"
      Tab(5).Control(8)=   "lblWebSearchesEmailTo"
      Tab(5).Control(9)=   "lblPostWebAddress"
      Tab(5).Control(10)=   "lblFTPRetypePassword"
      Tab(5).Control(11)=   "lblFTPPassword"
      Tab(5).Control(12)=   "lblFTPUser"
      Tab(5).Control(13)=   "lblFTPServer"
      Tab(5).ControlCount=   14
      TabCaption(6)   =   "General"
      TabPicture(6)   =   "frmToolsOptions.frx":00B4
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "txtDatabase"
      Tab(6).Control(1)=   "txtServerLocation"
      Tab(6).Control(2)=   "chkBasicImageWhere"
      Tab(6).Control(3)=   "fraToolsOptionsGeneralPictureSearch"
      Tab(6).Control(4)=   "labDatabase"
      Tab(6).Control(5)=   "labServerLocation"
      Tab(6).ControlCount=   6
      Begin VB.TextBox txtDatabase 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   -72420
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   93
         Top             =   3600
         Width           =   4800
      End
      Begin VB.TextBox txtServerLocation 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   -72420
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   91
         Top             =   2760
         Width           =   4800
      End
      Begin VB.CheckBox chkBasicImageWhere 
         Alignment       =   1  'Right Justify
         Caption         =   "Basic Image Where"
         Height          =   255
         Left            =   -74520
         TabIndex        =   86
         Top             =   2400
         Width           =   2295
      End
      Begin VB.Frame fraToolsOptionsGeneralPictureSearch 
         Caption         =   "Picture Search"
         Height          =   1635
         Left            =   -74760
         TabIndex        =   80
         Top             =   540
         Width           =   7215
         Begin VB.TextBox txtTooltipDelay 
            Height          =   285
            Left            =   2340
            MaxLength       =   1
            TabIndex        =   85
            Text            =   "0"
            Top             =   1200
            Visible         =   0   'False
            Width           =   255
         End
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   285
            Left            =   2580
            TabIndex        =   84
            Top             =   1200
            Visible         =   0   'False
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtTooltipDelay"
            BuddyDispid     =   196613
            OrigLeft        =   2640
            OrigTop         =   1200
            OrigRight       =   2880
            OrigBottom      =   1485
            Max             =   4
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.CheckBox chkMouseWheel 
            Alignment       =   1  'Right Justify
            Caption         =   "Mouse Wheel Scroll"
            Height          =   255
            Left            =   240
            TabIndex        =   82
            Top             =   840
            Width           =   2295
         End
         Begin VB.CheckBox chkFuzzyKeywordSearch 
            Alignment       =   1  'Right Justify
            Caption         =   "Fuzzy Keyword Search"
            Height          =   255
            Left            =   240
            TabIndex        =   81
            Top             =   480
            Width           =   2295
         End
         Begin VB.Label lblTooltipDelay 
            Caption         =   "Tool Tip Delay"
            Height          =   255
            Left            =   240
            TabIndex        =   83
            Top             =   1200
            Visible         =   0   'False
            Width           =   1455
         End
      End
      Begin VB.TextBox txtWebSearchesEmailFrom 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   -72765
         MaxLength       =   100
         TabIndex        =   75
         Top             =   3420
         Width           =   5200
      End
      Begin VB.TextBox txtWebSearchesEmailTo 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   -72765
         MaxLength       =   100
         TabIndex        =   73
         Top             =   2940
         Width           =   5200
      End
      Begin VB.TextBox txtPostWebAddress 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   -72765
         MaxLength       =   100
         TabIndex        =   71
         Top             =   2460
         Width           =   5200
      End
      Begin VB.TextBox txtFTPRetypePassword 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   -72765
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   69
         Top             =   1980
         Width           =   2580
      End
      Begin VB.TextBox txtFTPPassword 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   -72765
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   67
         Top             =   1500
         Width           =   2580
      End
      Begin VB.TextBox txtFTPUser 
         Height          =   315
         Left            =   -72765
         MaxLength       =   100
         TabIndex        =   65
         Top             =   1020
         Width           =   5200
      End
      Begin VB.TextBox txtFTPServer 
         Height          =   315
         Left            =   -72765
         MaxLength       =   100
         TabIndex        =   63
         Top             =   540
         Width           =   5200
      End
      Begin VB.CommandButton cmdFixIncorrectDescriptions 
         Caption         =   "Fix Incorrect Descriptions"
         Height          =   375
         Left            =   -73080
         TabIndex        =   61
         Top             =   2760
         Width           =   3615
      End
      Begin VB.Frame fraImageLocations 
         Caption         =   "Image Locations"
         Height          =   3975
         Left            =   -74760
         TabIndex        =   9
         Top             =   1860
         Width           =   7335
         Begin VB.CommandButton cmdEditLocation 
            Caption         =   "&Edit"
            Height          =   375
            Left            =   5160
            TabIndex        =   11
            Top             =   3480
            Width           =   975
         End
         Begin TabDlg.SSTab sstImageLocations 
            Height          =   3060
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   7005
            _ExtentX        =   12356
            _ExtentY        =   5398
            _Version        =   393216
            Style           =   1
            Tabs            =   2
            TabHeight       =   520
            TabCaption(0)   =   "Default Locations"
            TabPicture(0)   =   "frmToolsOptions.frx":00D0
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "smgDefaultImageLocations"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Machine Specific Locations"
            TabPicture(1)   =   "frmToolsOptions.frx":00EC
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "smgMachineImageLocations"
            Tab(1).ControlCount=   1
            Begin ImageWhere.SimpleGrid smgDefaultImageLocations 
               Height          =   2475
               Left            =   120
               TabIndex        =   89
               Top             =   360
               Width           =   6735
               _ExtentX        =   11880
               _ExtentY        =   4366
               Columns         =   1
               KeyCol          =   0
            End
            Begin ImageWhere.SimpleGrid smgMachineImageLocations 
               Height          =   2475
               Left            =   -74880
               TabIndex        =   90
               Top             =   360
               Width           =   6735
               _ExtentX        =   11880
               _ExtentY        =   4366
               Columns         =   1
               KeyCol          =   0
            End
         End
         Begin VB.CommandButton cmdDeleteLocation 
            Caption         =   "&Delete"
            Enabled         =   0   'False
            Height          =   375
            Left            =   6240
            TabIndex        =   12
            Top             =   3480
            Width           =   975
         End
      End
      Begin VB.Frame fraEmails 
         Caption         =   "Email Templates"
         Height          =   1695
         Left            =   -74880
         TabIndex        =   51
         Top             =   2940
         Width           =   7575
         Begin VB.TextBox txtEmailsLastUpgraded 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   3240
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   55
            Top             =   720
            Width           =   4200
         End
         Begin VB.TextBox txtDateEmailLastPosted 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   3240
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   53
            Top             =   360
            Width           =   4200
         End
         Begin VB.CommandButton cmdUpgradeTemplates 
            Caption         =   "&Upgrade"
            Height          =   375
            Left            =   6480
            TabIndex        =   56
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label lblEmailsLastUpgraded 
            Caption         =   "Last Upgraded:"
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   720
            Width           =   1755
         End
         Begin VB.Label lblEmailsLastPosted 
            Caption         =   "Last Posted:"
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   360
            Width           =   1755
         End
      End
      Begin VB.Frame fraLogFiles 
         Caption         =   "Log Files"
         Height          =   2295
         Left            =   -74880
         TabIndex        =   43
         Top             =   540
         Width           =   7575
         Begin VB.CommandButton cmdPost 
            Caption         =   "&Post"
            Height          =   375
            Left            =   6480
            TabIndex        =   50
            Top             =   1800
            Width           =   975
         End
         Begin VB.TextBox txtDateLogFilePosted 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   3240
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   45
            Top             =   360
            Width           =   4200
         End
         Begin VB.TextBox txtLogFilePostingFrequency 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3240
            MaxLength       =   2
            TabIndex        =   47
            Top             =   720
            Width           =   600
         End
         Begin VB.TextBox txtSupportEmail 
            Height          =   555
            Left            =   3240
            MaxLength       =   50
            MultiLine       =   -1  'True
            TabIndex        =   49
            Top             =   1080
            Width           =   4200
         End
         Begin VB.Label lblDateLogFilePosted 
            Caption         =   "Last Posted:"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   360
            Width           =   1755
         End
         Begin VB.Label lblLogFilePostingFrequency 
            Caption         =   "Posting Frequency (Days):"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   720
            Width           =   2955
         End
         Begin VB.Label lblSupportEmail 
            Caption         =   "Support Email Address:"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   1080
            Width           =   2955
         End
      End
      Begin VB.TextBox txtUpgradeFrequency 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   -71640
         MaxLength       =   4
         TabIndex        =   60
         Top             =   5340
         Width           =   600
      End
      Begin VB.TextBox txtCompanyName 
         Height          =   315
         Left            =   -71805
         MaxLength       =   50
         TabIndex        =   14
         Top             =   540
         Width           =   4000
      End
      Begin VB.TextBox txtWebSearchTestEmail 
         Height          =   555
         Left            =   -71640
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   58
         Top             =   4740
         Width           =   4200
      End
      Begin VB.TextBox txtInfo1 
         Height          =   315
         Left            =   -71805
         TabIndex        =   40
         Text            =   "Partners Clive Nichols & Jane Nichols"
         Top             =   5580
         Width           =   4000
      End
      Begin VB.TextBox txtInfo2 
         Height          =   315
         Left            =   -71805
         MaxLength       =   30
         TabIndex        =   42
         Text            =   "Member of BAPLA"
         Top             =   5940
         Width           =   4000
      End
      Begin VB.TextBox txtWebSite 
         Height          =   315
         Left            =   -71805
         MaxLength       =   30
         TabIndex        =   38
         Text            =   "www.clivenichols.com"
         Top             =   5220
         Width           =   4000
      End
      Begin VB.TextBox txtEmail 
         Height          =   315
         Left            =   -71805
         MaxLength       =   30
         TabIndex        =   36
         Text            =   "enquiries@clivenichols.com"
         Top             =   4860
         Width           =   4000
      End
      Begin VB.TextBox txtVATNo 
         Height          =   315
         Left            =   -71805
         MaxLength       =   30
         TabIndex        =   34
         Text            =   "569 7570 79"
         Top             =   4500
         Width           =   4000
      End
      Begin VB.TextBox txtFaxNo 
         Height          =   315
         Left            =   -71805
         MaxLength       =   30
         TabIndex        =   32
         Text            =   "+44 (0)1295 713672"
         Top             =   4140
         Width           =   4000
      End
      Begin VB.TextBox txtTelNo 
         Height          =   315
         Left            =   -71805
         TabIndex        =   30
         Text            =   "+44 (0)1295 712288"
         Top             =   3780
         Width           =   4000
      End
      Begin VB.TextBox txtPostCode 
         Height          =   315
         Left            =   -71805
         MaxLength       =   30
         TabIndex        =   28
         Text            =   "OX17 2EN"
         Top             =   3420
         Width           =   4000
      End
      Begin VB.TextBox txtCountry 
         Height          =   315
         Left            =   -71805
         MaxLength       =   30
         TabIndex        =   26
         Text            =   "England"
         Top             =   3060
         Width           =   4000
      End
      Begin VB.TextBox txtCounty 
         Height          =   315
         Left            =   -71805
         MaxLength       =   30
         TabIndex        =   24
         Text            =   "Oxon"
         Top             =   2700
         Width           =   4000
      End
      Begin VB.TextBox txtTown 
         Height          =   315
         Left            =   -71805
         MaxLength       =   30
         TabIndex        =   22
         Text            =   "Banbury"
         Top             =   2340
         Width           =   4000
      End
      Begin VB.TextBox txtAddress3 
         Height          =   315
         Left            =   -71805
         MaxLength       =   30
         TabIndex        =   20
         Text            =   "Chacombe"
         Top             =   1980
         Width           =   4000
      End
      Begin VB.TextBox txtAddress2 
         Height          =   315
         Left            =   -71805
         MaxLength       =   30
         TabIndex        =   19
         Text            =   "Castle Farm"
         Top             =   1620
         Width           =   4000
      End
      Begin VB.TextBox txtAddress1 
         Height          =   315
         Left            =   -71805
         MaxLength       =   30
         TabIndex        =   18
         Text            =   "Rickyard Barn"
         Top             =   1260
         Width           =   4000
      End
      Begin VB.TextBox txtMaxImagesPerPage 
         Height          =   315
         Left            =   -71805
         MaxLength       =   6
         TabIndex        =   8
         Text            =   "5000"
         Top             =   1380
         Width           =   1035
      End
      Begin VB.TextBox txtImageHeightWidth 
         Height          =   315
         Left            =   -71805
         MaxLength       =   6
         TabIndex        =   6
         Text            =   "5000"
         Top             =   960
         Width           =   1035
      End
      Begin VB.TextBox txtSignatory 
         Height          =   315
         Left            =   -71805
         MaxLength       =   30
         TabIndex        =   16
         Text            =   "ALI PHILPOTTS"
         Top             =   900
         Width           =   4000
      End
      Begin VB.TextBox txtImagesReturned 
         Height          =   315
         Left            =   -71805
         MaxLength       =   6
         TabIndex        =   4
         Text            =   "5000"
         Top             =   540
         Width           =   1035
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   4560
         TabIndex        =   0
         Top             =   5940
         Width           =   975
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   5640
         TabIndex        =   1
         Top             =   5940
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   6720
         TabIndex        =   2
         Top             =   5940
         Width           =   975
      End
      Begin ImageWhere.SimpleGrid smgBusinessTypes 
         Height          =   5355
         Left            =   240
         TabIndex        =   88
         Top             =   540
         Width           =   7365
         _ExtentX        =   12991
         _ExtentY        =   9446
         Columns         =   1
         KeyCol          =   0
      End
      Begin VB.Label labDatabase 
         Caption         =   "Database:"
         Height          =   255
         Left            =   -74520
         TabIndex        =   92
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label labServerLocation 
         Caption         =   "Server Location:"
         Height          =   255
         Left            =   -74520
         TabIndex        =   87
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label lblWebSearchesEmailFrom 
         Caption         =   "Web Searches Email From:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   74
         Top             =   3420
         Width           =   1935
      End
      Begin VB.Label lblWebSearchesEmailTo 
         Caption         =   "Web Searches Email To:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   72
         Top             =   2940
         Width           =   1815
      End
      Begin VB.Label lblPostWebAddress 
         Caption         =   "Web Address:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   70
         Top             =   2460
         Width           =   1095
      End
      Begin VB.Label lblFTPRetypePassword 
         Caption         =   "Retype Password:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   68
         Top             =   1980
         Width           =   1335
      End
      Begin VB.Label lblFTPPassword 
         Caption         =   "Password:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   66
         Top             =   1500
         Width           =   1095
      End
      Begin VB.Label lblFTPUser 
         Caption         =   "User:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   64
         Top             =   1020
         Width           =   1095
      End
      Begin VB.Label lblFTPServer 
         Caption         =   "Server:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   62
         Top             =   540
         Width           =   1095
      End
      Begin VB.Label lblUpgradeFrequency 
         Caption         =   "Application Upgrade Check Frequency (Minutes):"
         Height          =   495
         Left            =   -74760
         TabIndex        =   59
         Top             =   5340
         Width           =   2955
      End
      Begin VB.Label lblCompanyName 
         Caption         =   "Company Name:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   13
         Top             =   540
         Width           =   1755
      End
      Begin VB.Label lblWebSearchTestEmail 
         Caption         =   "Web Search Test Email Address:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   57
         Top             =   4740
         Width           =   2955
      End
      Begin VB.Label lblInfo1 
         Caption         =   "Information 1:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   39
         Top             =   5580
         Width           =   1755
      End
      Begin VB.Label lblInfo2 
         Caption         =   "Information 2:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   41
         Top             =   5940
         Width           =   1755
      End
      Begin VB.Label lblWebSite 
         Caption         =   "Web Site:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   37
         Top             =   5220
         Width           =   1755
      End
      Begin VB.Label lblEmail 
         Caption         =   "Email:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   35
         Top             =   4860
         Width           =   1755
      End
      Begin VB.Label lblVATNo 
         Caption         =   "VAT Number:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   33
         Top             =   4500
         Width           =   1755
      End
      Begin VB.Label lblFax 
         Caption         =   "Fax:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   31
         Top             =   4140
         Width           =   1755
      End
      Begin VB.Label lblTelephone 
         Caption         =   "Telephone:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   29
         Top             =   3780
         Width           =   1755
      End
      Begin VB.Label lblPostCode 
         Caption         =   "Post Code:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   27
         Top             =   3420
         Width           =   1755
      End
      Begin VB.Label lblCountry 
         Caption         =   "Country:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   25
         Top             =   3060
         Width           =   1755
      End
      Begin VB.Label lblCounty 
         Caption         =   "County:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   23
         Top             =   2700
         Width           =   1755
      End
      Begin VB.Label lblTown 
         Caption         =   "Town:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   21
         Top             =   2340
         Width           =   1755
      End
      Begin VB.Label lblAddress 
         Caption         =   "Address:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   17
         Top             =   1260
         Width           =   1755
      End
      Begin VB.Label lblMaxImagesPerPage 
         Caption         =   "Maximum Images Shown on a Page:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   7
         Top             =   1380
         Width           =   2775
      End
      Begin VB.Label lblCms 
         Caption         =   "cms."
         Height          =   255
         Left            =   -70620
         TabIndex        =   79
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label blImageHeightWidth 
         Caption         =   "Height/Width of Images Shown:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   5
         Top             =   960
         Width           =   2775
      End
      Begin VB.Label lblSignatory 
         Caption         =   "Chaser Letter Signatory:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   15
         Top             =   900
         Width           =   1755
      End
      Begin VB.Label lblImagesReturned 
         Caption         =   "Maximum Images Returned by Search:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   3
         Top             =   540
         Width           =   2775
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6960
      TabIndex        =   77
      Top             =   6540
      Width           =   975
   End
   Begin VB.Menu mnuMaintain 
      Caption         =   "Maintain"
      Begin VB.Menu mnuMaintainAdd 
         Caption         =   "&Add"
      End
      Begin VB.Menu mnuMaintainEdit 
         Caption         =   "&Edit"
      End
      Begin VB.Menu mnuMaintainDelete 
         Caption         =   "&Delete"
      End
   End
End
Attribute VB_Name = "frmToolsOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mobjBusinessTypes               As BusinessTypes
Private WithEvents fBusinessTypeMaint   As frmBusinessTypeMaint
Attribute fBusinessTypeMaint.VB_VarHelpID = -1
Private mstrCurrentImageType            As String
Private WithEvents fLocationEdit        As frmLocationEdit
Attribute fLocationEdit.VB_VarHelpID = -1
Private mvarBasicImageWhere             As Boolean

Private Sub AddBusinessType()
'***************************************
' Module/Form Name   : frmToolsOptions
'
' Procedure Name     : AddBusinessType
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
On Error GoTo AddBusinessType_Error
'
'******** Code Starts Here *************
'
    Set fBusinessTypeMaint = New frmBusinessTypeMaint
    fBusinessTypeMaint.Display Add, mobjBusinessTypes(smgBusinessTypes.Column(smgBusinessTypes.KeyCol).Value)
    Set fBusinessTypeMaint = Nothing
    cmdCancel.Caption = "&Close"
    '
    '********* Code Ends Here **************
    '
    Exit Sub
    '
AddBusinessType_Error:
    ErrorRaise "frmToolsOptions.AddBusinessType"
End Sub

Private Sub cmdAdd_Click()
'***************************************
' Module/Form Name   : frmToolsOptions
'
' Procedure Name     : cmdAdd_Click
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
On Error GoTo cmdAdd_Click_Error
'
'******** Code Starts Here *************
'
    AddBusinessType
'
'********* Code Ends Here **************
'
   Exit Sub
'
cmdAdd_Click_Error:
    DisplayError , "frmToolsOptions.cmdAdd_Click", vbExclamation
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
'***************************************
' Module/Form Name   : frmToolsOptions
'
' Procedure Name     : cmdDelete_Click
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
On Error GoTo cmdDelete_Click_Error
'
'******** Code Starts Here *************
'
    DeleteBusinessType
'
'********* Code Ends Here **************
'
   Exit Sub
'
cmdDelete_Click_Error:
    DisplayError , "frmToolsOptions.cmdDelete_Click", vbExclamation
End Sub

Private Sub cmdEdit_Click()
'***************************************
' Module/Form Name   : frmToolsOptions
'
' Procedure Name     : cmdEdit_Click
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
On Error GoTo cmdEdit_Click_Error
'
'******** Code Starts Here *************
'
    EditBusinessType
'
'********* Code Ends Here **************
'
   Exit Sub
'
cmdEdit_Click_Error:
    DisplayError , "frmToolsOptions.cmdEdit_Click", vbExclamation
End Sub

Private Sub cmdEditLocation_Click()
'***************************************
' Module/Form Name   : frmToolsOptions
'
' Procedure Name     : cmdEditLocation_Click
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
On Error GoTo cmdEditLocation_Click_Error
'
'******** Code Starts Here *************
'
    EditLocation
'
'********* Code Ends Here **************
'
   Exit Sub
'
cmdEditLocation_Click_Error:
    DisplayError , "frmToolsOptions.cmdEditLocation_Click", vbExclamation
End Sub

Private Sub cmdFixIncorrectDescriptions_Click()
'***************************************
' Module/Form Name   : frmToolsOptions
'
' Procedure Name     : cmdFixIncorrectDescriptions_Click
'
' Purpose            :
'
' Date Created       : 04/10/2004
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo cmdFixIncorrectDescriptions_Click_Error
'
'******** Code Starts Here *************
'
    Dim strSQL      As String
    
    If MsgBox("Please ensure you have taken a backup of your database." & vbCrLf & "Do you wish to continue?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
        Exit Sub
    End If
    
    strSQL = "UPDATE search_result INNER JOIN Photograph ON (search_result.photograph_no = Photograph.photograph_no) AND (search_result.batch_no = Photograph.batch_no) "
    strSQL = strSQL & "Set search_result.photograph_key = [photograph].[photograph_key] "
    strSQL = strSQL & "WHERE (((search_result.photograph_key)=1) AND ((search_result.batch_no)<>1))"
    db.Execute strSQL, dbSeeChanges + dbFailOnError
    '
    SaveSetting App.Title, "ToolsOptions", "DataFix", "N"
    sstBusinessTypes.TabVisible(4) = False
    '
    MsgBox "Your database has been fixed. This tab will no longer be available.", vbInformation
'
'********* Code Ends Here **************
'
   Exit Sub
'
cmdFixIncorrectDescriptions_Click_Error:
    DisplayError , "frmToolsOptions.cmdFixIncorrectDescriptions_Click", vbExclamation
End Sub

Private Sub cmdOK_Click()
'***************************************
' Module/Form Name   : frmToolsOptions
'
' Procedure Name     : cmdOK_Click
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
On Error GoTo cmdOK_Click_Error
'
'******** Code Starts Here *************
'
    If Not ValidEntry Then
        Exit Sub
    End If
    '
    '   Update the Company Info object.
    '
    With goCompanyInfo
        .CompanyName = txtCompanyName.Text
        .Signatory = txtSignatory.Text
        .Address1 = txtAddress1.Text
        .Address2 = txtAddress2.Text
        .Address3 = txtAddress3.Text
        .Town = txtTown.Text
        .County = txtCounty.Text
        .Country = txtCountry.Text
        .PostCode = txtPostCode.Text
        .TelNo = txtTelNo.Text
        .FaxNo = txtFaxNo.Text
        .VATNo = txtVATNo.Text
        .Email = txtEmail.Text
        .WebSite = txtWebSite.Text
        .Info1 = txtInfo1.Text
        .Info2 = txtInfo2.Text
        .update
    End With
    '
    '   Update the System Config object.
    '
    With goSystemConfig
        .ImagesReturned = CLng(txtImagesReturned.Text)
        .ImageHeightWidthCms = CDbl(txtImageHeightWidth.Text)
        .ImagesPerPage = CLng(txtMaxImagesPerPage.Text)
        If goSystemConfig.SupportUser Then
            .LogFilePostingFrequency = CInt(txtLogFilePostingFrequency.Text)
            .SupportEmail = txtSupportEmail.Text
            .WebSearchTestEmail = txtWebSearchTestEmail.Text
            .UpgradeCheckFrequency = CInt(txtUpgradeFrequency.Text)
            .FTPServer = Trim(txtFTPServer.Text)
            .FTPUser = Trim(txtFTPUser.Text)
            .FTPPassword = Trim(txtFTPPassword.Text)
            .PostWebAddress = Trim(txtPostWebAddress.Text)
            .WebSearchesEmailTo = Trim(txtWebSearchesEmailTo.Text)
            .WebSearchesEmailFrom = Trim(txtWebSearchesEmailFrom.Text)
        End If
        .FuzzyKeywordSearch = (chkFuzzyKeywordSearch.Value = vbChecked)
        .MouseWheel = (chkMouseWheel.Value = vbChecked)
        .TooltipDelay = CInt(txtTooltipDelay.Text)
        .BasicImageWhere = (chkBasicImageWhere.Value = vbChecked)
        .update
    End With
    
    mdi_npls.RefreshCaption

    If mvarBasicImageWhere <> goSystemConfig.BasicImageWhere Then
        MsgBox "Note that Image Where needs to be restarted for the new 'Basic Image Where' setting to take effect", vbInformation, App.Title
    End If

    Unload Me
'
'********* Code Ends Here **************
'
   Exit Sub
'
cmdOK_Click_Error:
    DisplayError , "frmToolsOptions.cmdOK_Click", vbExclamation
End Sub

Private Function ValidEntry() As Boolean
'***************************************
' Module/Form Name   : frmToolsOptions
'
' Procedure Name     : ValidEntry
'
' Purpose            :
'
' Date Created       : 14/02/2005
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
On Error GoTo ValidEntry_Error
'
'******** Code Starts Here *************
'
    ValidEntry = False
    '
    '   Validate the input.
    '
    If Not IsNumeric(txtImagesReturned.Text) Then
        MsgBox "Enter a valid number.", vbExclamation
        sstBusinessTypes.Tab = 1
        txtImagesReturned.SetFocus
        Exit Function
    End If

    If CLng(txtImagesReturned.Text) > goSystemConfig.MaxImagesPerSearch Or _
       CLng(txtImagesReturned.Text) < 1 Then
        MsgBox "Images Returned must be between 1 and " & CStr(goSystemConfig.MaxImagesPerSearch) & " inclusive.", vbExclamation
        sstBusinessTypes.Tab = 1
        txtImagesReturned.SetFocus
        Exit Function
    End If

    If Not IsNumeric(txtImageHeightWidth.Text) Then
        MsgBox "Enter a valid number.", vbExclamation
        sstBusinessTypes.Tab = 1
        txtImageHeightWidth.SetFocus
        Exit Function
    End If

    If CDbl(txtImageHeightWidth.Text) > goSystemConfig.MaxImageHeightWidth Or _
       CDbl(txtImageHeightWidth.Text) < goSystemConfig.MinImageHeightWidth Then
        MsgBox "Images Height/Width must be between " & Format(goSystemConfig.MinImageHeightWidth, "##0.00") & " and " & Format(goSystemConfig.MaxImageHeightWidth, "##0.00") & " inclusive.", vbExclamation
        sstBusinessTypes.Tab = 1
        txtImageHeightWidth.SetFocus
        Exit Function
    End If

    If Not IsNumeric(txtMaxImagesPerPage.Text) Then
        MsgBox "Enter a valid number.", vbExclamation
        sstBusinessTypes.Tab = 1
        txtMaxImagesPerPage.SetFocus
        Exit Function
    End If

    If CLng(txtMaxImagesPerPage.Text) > goSystemConfig.MaxImagesPerPage Or _
       CLng(txtMaxImagesPerPage.Text) < 1 Then
        MsgBox "Images Per Page must be between 1 and " & CStr(goSystemConfig.MaxImagesPerPage) & " inclusive.", vbExclamation
        sstBusinessTypes.Tab = 1
        txtMaxImagesPerPage.SetFocus
        Exit Function
    End If
    '
    If goSystemConfig.SupportUser Then
        If Trim(txtLogFilePostingFrequency.Text) = "" Then
            MsgBox "Please enter a Posting Frequency for the Support Log Files." & vbCrLf & "Enter 0 to stop Posting.", vbExclamation
            sstBusinessTypes.Tab = 3
            txtLogFilePostingFrequency.SetFocus
            Exit Function
        End If
        '
        If Trim(txtUpgradeFrequency.Text) = "" Then
            MsgBox "Please enter a Frequency for Checking for Upgrades.", vbExclamation
            sstBusinessTypes.Tab = 3
            txtUpgradeFrequency.SetFocus
            Exit Function
        End If
        '
        If Trim(txtFTPPassword.Text) = "" Then
            MsgBox "Please enter a Password.", vbExclamation
            sstBusinessTypes.Tab = 5
            txtFTPPassword.SetFocus
            Exit Function
        End If
        '
        If Trim(txtFTPRetypePassword.Text) = "" Then
            MsgBox "Please retype your Password.", vbExclamation
            sstBusinessTypes.Tab = 5
            txtFTPRetypePassword.SetFocus
            Exit Function
        End If
        '
        If txtFTPPassword.Text <> txtFTPRetypePassword.Text Then
            MsgBox "Passwords Differ. Please re-enter.", vbExclamation
            sstBusinessTypes.Tab = 5
            txtFTPPassword.SetFocus
            Exit Function
        End If
    End If
    '
    If Trim(txtTooltipDelay.Text) = "" Then
        MsgBox "Please enter a value for the Tool Tip Delay.", vbExclamation
        sstBusinessTypes.Tab = 6
        txtTooltipDelay.SetFocus
        Exit Function
    End If
    
    ValidEntry = True
'
'********* Code Ends Here **************
'
   Exit Function
'
ValidEntry_Error:
    ErrorRaise "frmToolsOptions.ValidEntry"
End Function

Private Sub cmdPost_Click()
'***************************************
' Module/Form Name   : frmToolsOptions
'
' Procedure Name     : cmdPost_Click
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
On Error GoTo cmdPost_Click_Error
'
'******** Code Starts Here *************
'
    If MsgBox("Are you sure you wish to Post the log files to Image Where Support immediately?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
        Exit Sub
    End If
    '
    '   Post the log files.
    '
    goLog.PostLogFile
    RefreshSupportAndFTPTabs
    '
'
'********* Code Ends Here **************
'
   Exit Sub
'
cmdPost_Click_Error:
    DisplayError , "frmToolsOptions.cmdPost_Click", vbExclamation
End Sub

Private Sub DeleteBusinessType()
'***************************************
' Module/Form Name   : frmToolsOptions
'
' Procedure Name     : DeleteBusinessType
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
On Error GoTo DeleteBusinessType_Error
'
'******** Code Starts Here *************
'
    Dim intTopRow As Integer

    If MsgBox("Are you sure you wish to delete business type: '" & smgBusinessTypes.Column(smgBusinessTypes.KeyCol).Value & "'", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    With mobjBusinessTypes
        On Error Resume Next
        intTopRow = smgBusinessTypes.TopRow
        .Item(smgBusinessTypes.Column(smgBusinessTypes.KeyCol).Value).Delete
        If Err.Number - vbObjectError = 4 Then
            MsgBox "Business Type '" & smgBusinessTypes.Column(smgBusinessTypes.KeyCol).Value & "' is used by one or more customers and cannot be deleted.", vbExclamation
            Screen.MousePointer = vbDefault
            Exit Sub
        ElseIf Err.Number = 0 Then
            On Error GoTo DeleteBusinessType_Error
        Else
            GoTo DeleteBusinessType_Error
        End If
        .Refresh
        RefreshBusinessTypes
        smgBusinessTypes.GetKeyRow mobjBusinessTypes.CurrentName
        smgBusinessTypes.TopRow = intTopRow
    End With
    cmdCancel.Caption = "&Close"
    Screen.MousePointer = vbDefault
    '
    '********* Code Ends Here **************
    '
    Exit Sub
    '
DeleteBusinessType_Error:
    ErrorRaise "frmToolsOptions.DeleteBusinessType"
End Sub

Private Sub EditLocation()
    Set fLocationEdit = New frmLocationEdit
    fLocationEdit.Display goSystemConfig.Locations(smgDefaultImageLocations.Column(smgDefaultImageLocations.KeyCol).Value)
    
    Set fLocationEdit = Nothing

End Sub

Private Sub EditBusinessType()
'***************************************
' Module/Form Name   : frmToolsOptions
'
' Procedure Name     : EditBusinessType
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
On Error GoTo EditBusinessType_Error
'
'******** Code Starts Here *************
'
    Set fBusinessTypeMaint = New frmBusinessTypeMaint
    fBusinessTypeMaint.Display Edit, mobjBusinessTypes(smgBusinessTypes.Column(smgBusinessTypes.KeyCol).Value)
    Set fBusinessTypeMaint = Nothing
    cmdCancel.Caption = "&Close"
'
'********* Code Ends Here **************
'
   Exit Sub
'
EditBusinessType_Error:
    ErrorRaise "frmToolsOptions.EditBusinessType"
End Sub

Private Sub cmdUpgradeTemplates_Click()
'***************************************
' Module/Form Name   : frmToolsOptions
'
' Procedure Name     : cmdUpgradeTemplates_Click
'
' Purpose            :
'
' Date Created       : 03/01/2003
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo cmdUpgradeTemplates_Click_Error
'
'******** Code Starts Here *************
'
    Dim strPostTemplate     As String
    Dim strConfirmTemplate  As String
    Dim strCancelTemplate   As String
    Dim strBundledTemplate  As String
    
    Screen.MousePointer = vbHourglass
    '
    '   Make a call to the CNGP web site using XML to pull back the templates.
    '
    If RetrieveEmailTemplates(strPostTemplate, strConfirmTemplate, strCancelTemplate, strBundledTemplate) Then
        '
        '   Write each template away to the database.
        '
        With goSystemConfig
            .PostedHTMLEmail = Replace(strPostTemplate, Chr(10), Chr(13))
            .ConfirmationHTMLEmail = Replace(strConfirmTemplate, Chr(10), Chr(13))
            .CancellationHTMLEmail = Replace(strCancelTemplate, Chr(10), Chr(13))
            .BundledHTMLEmail = Replace(strBundledTemplate, Chr(10), Chr(13))
            .DateHTMLEmailUpgraded = Now
            .update
        End With
        '
        RefreshSupportAndFTPTabs
        MsgBox "Email Templates have been upgraded to the latest.", vbInformation, App.Title
    Else
        MsgBox "Email Templates have failed to upgrade to the latest. Please contact Image Where support.", vbExclamation, App.Title
    End If
    '
    Screen.MousePointer = vbDefault
'
'********* Code Ends Here **************
'
   Exit Sub
'
cmdUpgradeTemplates_Click_Error:
    DisplayError , "frmToolsOptions.cmdUpgradeTemplates_Click", vbExclamation
End Sub

Private Function RetrieveEmailTemplates(ByRef strPostTemplate As String, _
                                        ByRef strConfirmTemplate As String, _
                                        ByRef strCancelTemplate As String, _
                                        ByRef strBundledTemplate As String) As Boolean
'***************************************
' Module/Form Name   : frmToolsOptions
'
' Procedure Name     : RetrieveEmailTemplates
'
' Purpose            :
'
' Date Created       : 03/01/2003
'
' Author             : GARETH SAUNDERS
'
' Parameters         : strPostTemplate - String
'                    : strConfirmTemplate - String
'                    : strCancelTemplate - String
'
' Returns            : Boolean
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo RetrieveEmailTemplates_Error
'
'******** Code Starts Here *************
'
    Dim objHTTP As New MSXML2.XMLHTTP
    Dim strEnvelope As String
    Dim strReturn As String
    Dim objReturn As New MSXML2.DOMDocument30
    Dim objNode As MSXML2.IXMLDOMNode
    Dim strQuery As String
    Dim strSuccess As String
    Dim fError As frmHTMLError
    Dim strError As String
    
    RetrieveEmailTemplates = False
    
    strEnvelope = "<SOAP:Envelope xmlns:SOAP=""urn:schemas-xmlsoap-org:soap.v1"">" & _
                  "<SOAP:Body>" & _
                  "<m:IWExportEmails xmlns:m=""urn:imagewhere/soap:ExportEmails"">" & _
                  "<Version></Version>" & _
                  "</m:IWExportEmails>" & _
                  "</SOAP:Body>" & _
                  "</SOAP:Envelope>"

    With objHTTP
        .Open "post", "http://www.clivenichols.com/websearches/exportemails.asp", False
        .setRequestHeader "Content-Type", "text/xml"
        .setRequestHeader "SOAPMethodName", "urn:imagewhere/soap:ExportEmails#IWExportEmails"
        .Send strEnvelope
        strReturn = .responseText
    End With

    objReturn.loadXML strReturn

    strQuery = "SOAP:Envelope/SOAP:Body/m:IWExportEmails/Result"

    strSuccess = "Fail"
    On Error Resume Next
    strSuccess = objReturn.selectSingleNode(strQuery).Text
    On Error GoTo RetrieveEmailTemplates_Error
    If strSuccess <> "Success" Then
        Set fError = New frmHTMLError
        fError.Display strReturn
        strError = fError.ErrorText
        Set fError = Nothing
        Exit Function
    End If
    '
    '   Return the Post Email Template.
    '
    strQuery = "SOAP:Envelope/SOAP:Body/m:IWExportEmails/PostedEmail"
    strPostTemplate = objReturn.selectSingleNode(strQuery).Text
    strPostTemplate = ConvertIllegalCharsToTags(strPostTemplate)
    '
    '   Return the Confirm Email Template.
    '
    strQuery = "SOAP:Envelope/SOAP:Body/m:IWExportEmails/ConfirmEmail"
    strConfirmTemplate = objReturn.selectSingleNode(strQuery).Text
    strConfirmTemplate = ConvertIllegalCharsToTags(strConfirmTemplate)
    '
    '   Return the Cancel Email Template.
    '
    strQuery = "SOAP:Envelope/SOAP:Body/m:IWExportEmails/CancelEmail"
    strCancelTemplate = objReturn.selectSingleNode(strQuery).Text
    strCancelTemplate = ConvertIllegalCharsToTags(strCancelTemplate)
    '
    '   Return the Bundled Email Template.
    '
    strQuery = "SOAP:Envelope/SOAP:Body/m:IWExportEmails/BundledEmail"
    strBundledTemplate = objReturn.selectSingleNode(strQuery).Text
    strBundledTemplate = ConvertIllegalCharsToTags(strBundledTemplate)
    '
    RetrieveEmailTemplates = True
'
'********* Code Ends Here **************
'
   Exit Function
'
RetrieveEmailTemplates_Error:
    ErrorRaise "frmToolsOptions.RetrieveEmailTemplates"
End Function

Private Function ConvertIllegalCharsToTags(strReplace) As String
    Dim strTemp As String
    
    strTemp = Replace(strReplace, "&lt;", "<")
    strTemp = Replace(strTemp, "&gt;", ">")
    strTemp = Replace(strTemp, "&amp;", "&")
    strTemp = Replace(strTemp, "&apos;", "'")
    strTemp = Replace(strTemp, "&quot;", """")
    ConvertIllegalCharsToTags = strTemp

End Function

Private Sub fBusinessTypeMaint_BusinessTypeAdded(Name As String)
'***************************************
' Module/Form Name   : frmToolsOptions
'
' Procedure Name     : fBusinessTypeMaint_BusinessTypeAdded
'
' Purpose            :
'
' Date Created       : 31/12/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : Name - String
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo fBusinessTypeMaint_BusinessTypeAdded_Error
'
'******** Code Starts Here *************
'
    mobjBusinessTypes.Refresh
    RefreshBusinessTypes
    smgBusinessTypes.GetKeyRow Name
'
'********* Code Ends Here **************
'
   Exit Sub
'
fBusinessTypeMaint_BusinessTypeAdded_Error:
    DisplayError , "frmToolsOptions.fBusinessTypeMaint_BusinessTypeAdded", vbExclamation
End Sub

Private Sub fBusinessTypeMaint_BusinessTypeUpdated(Name As String)
'***************************************
' Module/Form Name   : frmToolsOptions
'
' Procedure Name     : fBusinessTypeMaint_BusinessTypeUpdated
'
' Purpose            :
'
' Date Created       : 31/12/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : Name - String
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'
'***************************************
'
On Error GoTo fBusinessTypeMaint_BusinessTypeUpdated_Error
'
'******** Code Starts Here *************
'
    Dim intTopRow As Integer
     
    intTopRow = smgBusinessTypes.TopRow
    RefreshBusinessTypes
    With smgBusinessTypes
        .GetKeyRow Name
        .TopRow = intTopRow
    End With
'
'********* Code Ends Here **************
'
   Exit Sub
'
fBusinessTypeMaint_BusinessTypeUpdated_Error:
    DisplayError , "frmToolsOptions.fBusinessTypeMaint_BusinessTypeUpdated", vbExclamation
End Sub

Private Sub fLocationEdit_LocationUpdated(Key As String)
    RefreshImageLocations
    smgDefaultImageLocations.GetKeyRow Key
End Sub

Private Sub Form_Load()
'***************************************
' Module/Form Name   : frmToolsOptions
'
' Procedure Name     : Form_Load
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
On Error GoTo Form_Load_Error
'
'******** Code Starts Here *************
'
    Dim strImagesReturned As String
    '
    '   Position controls.
    '
    cmdCancel.Top = Me.Height - cmdCancel.Height - 200 - goSystemConfig.TitleBarHeight
    cmdOK.Top = cmdCancel.Top
    sstBusinessTypes.Height = cmdCancel.Top - sstBusinessTypes.Top - 100
    cmdAdd.Top = sstBusinessTypes.Height - cmdAdd.Height - 100
    cmdEdit.Top = cmdAdd.Top
    cmdDelete.Top = cmdAdd.Top
    smgBusinessTypes.Height = cmdAdd.Top - smgBusinessTypes.Top - 100
    '
    '   Popup menus are invisible.
    '
    mnuMaintain.Visible = False
    '
    With smgBusinessTypes
        .Columns = 7
        .Column(1).Header = "Name"
        .Column(2).Header = "Initial"
        .Column(2).Align = flexAlignRightCenter
        .Column(3).Header = "SL1"
        .Column(3).Align = flexAlignRightCenter
        .Column(4).Header = "SL2"
        .Column(4).Align = flexAlignRightCenter
        .Column(5).Header = "Phone1"
        .Column(5).Align = flexAlignRightCenter
        .Column(6).Header = "Phone2"
        .Column(6).Align = flexAlignRightCenter
        .Column(7).Header = "Loss Fee"
        .Column(7).Align = flexAlignRightCenter
        .KeyCol = 1
    End With
    '
    '   Image Locations.
    '
    With smgDefaultImageLocations
        .Columns = 4
        .Column(1).Header = "ID"
        .Column(1).Align = flexAlignLeftTop
        .Column(2).Header = "Name"
        .Column(2).Align = flexAlignLeftTop
        .Column(3).Header = "Prefix"
        .Column(3).Align = flexAlignLeftTop
        .Column(4).Header = "Suffix"
        .Column(4).Align = flexAlignLeftTop
        .KeyCol = 1
    End With
    '
    With smgMachineImageLocations
        .Columns = 4
        .Column(1).Header = "ID"
        .Column(1).Align = flexAlignLeftTop
        .Column(2).Header = "Name"
        .Column(2).Align = flexAlignLeftTop
        .Column(3).Header = "Prefix"
        .Column(3).Align = flexAlignLeftTop
        .Column(4).Header = "Suffix"
        .Column(4).Align = flexAlignLeftTop
        .KeyCol = 1
    End With
    '
    Set mobjBusinessTypes = New BusinessTypes
    RefreshBusinessTypes
    '
    RefreshImageLocations
    '
    '   Display the Signatory and Images Returned.
    '
    goSystemConfig.Refresh
    With goCompanyInfo
        .Refresh
        txtCompanyName.Text = .CompanyName
        txtSignatory.Text = .Signatory
        txtAddress1.Text = .Address1
        txtAddress2.Text = .Address2
        txtAddress3.Text = .Address3
        txtTown.Text = .Town
        txtCounty.Text = .County
        txtCountry.Text = .Country
        txtPostCode.Text = .PostCode
        txtTelNo.Text = .TelNo
        txtFaxNo.Text = .FaxNo
        txtVATNo.Text = .VATNo
        txtEmail.Text = .Email
        txtWebSite.Text = .WebSite
        txtInfo1.Text = .Info1
        txtInfo2.Text = .Info2
    End With
    txtImagesReturned.Text = CStr(goSystemConfig.ImagesReturned)
    txtMaxImagesPerPage.Text = CStr(goSystemConfig.ImagesPerPage)
    txtImageHeightWidth.Text = Format(goSystemConfig.ImageHeightWidthCms, "##0.00")
    '
    '   Support Tab
    '
    If goSystemConfig.SupportUser Then
        RefreshSupportAndFTPTabs
    Else
        sstBusinessTypes.TabVisible(3) = False
        sstBusinessTypes.TabVisible(4) = False
        sstBusinessTypes.TabVisible(5) = False
    End If
    '
    '   General Tab
    '
    chkFuzzyKeywordSearch.Value = IIf(goSystemConfig.FuzzyKeywordSearch, vbChecked, vbUnchecked)
    chkMouseWheel.Value = IIf(goSystemConfig.MouseWheel, vbChecked, vbUnchecked)
    txtTooltipDelay.Text = CStr(goSystemConfig.TooltipDelay)
    chkBasicImageWhere.Value = IIf(goSystemConfig.BasicImageWhere, vbChecked, vbUnchecked)
    txtServerLocation.Text = goSystemConfig.ServerLocation
    txtDatabase = glo_dbname
    '
    '   Remember the setting in case it has changed and we need to let the user know that Image Where needs to be restarted.
    '
    mvarBasicImageWhere = chkBasicImageWhere.Value
    '
    '   Set the initial Tab.
    '
    sstBusinessTypes_Click 0
'
'********* Code Ends Here **************
'
    Exit Sub
    '
Form_Load_Error:
    DisplayError , "frmToolsOptions.Form_Load", vbExclamation
End Sub

Private Sub Form_Resize()
'***************************************
' Module/Form Name   : frmToolsOptions
'
' Procedure Name     : Form_Resize
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
On Error GoTo Form_Resize_Error
'
'******** Code Starts Here *************
'
    With smgBusinessTypes
        .Column(1).Width = 2000
''        .Column(2).Width = (.Width - .Column(1).Width - goSystemConfig.VScrollBarWidth - 100) / 6
        .Column(2).Width = (.Width - .Column(1).Width - 375) / 6
        .Column(3).Width = .Column(2).Width
        .Column(4).Width = .Column(2).Width
        .Column(5).Width = .Column(2).Width
        .Column(6).Width = .Column(2).Width
        .Column(7).Width = .Column(2).Width
    End With
'
    With smgDefaultImageLocations
        .Column(1).Width = 0
        .Column(2).Width = 2000
        .Column(4).Width = 1000
        .Column(3).Width = (.Width - .Column(2).Width - .Column(4).Width - goSystemConfig.VScrollBarWidth - 100)
        .ResizeRows
    End With
'
    With smgMachineImageLocations
        .Column(1).Width = 0
        .Column(2).Width = 2000
        .Column(4).Width = 1000
        .Column(3).Width = (.Width - .Column(2).Width - .Column(4).Width - goSystemConfig.VScrollBarWidth - 100)
        .ResizeRows
    End With
'
'********* Code Ends Here **************
'
    Exit Sub
    '
Form_Resize_Error:
    DisplayError , "frmToolsOptions.Form_Resize", vbExclamation
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjBusinessTypes = Nothing
End Sub

Private Sub mnuMaintainAdd_Click()
    AddBusinessType
End Sub

Private Sub mnuMaintainDelete_Click()
    DeleteBusinessType
End Sub

Private Sub mnuMaintainEdit_Click()
    EditBusinessType
End Sub

Private Sub RefreshImageLocations()
'***************************************
' Module/Form Name   : frmToolsOptions
'
' Procedure Name     : RefreshImageLocations
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
On Error GoTo RefreshImageLocations_Error
'
'******** Code Starts Here *************
'
    Dim oLocation       As Location
    
    With smgDefaultImageLocations
        .Redraw = False
        sstBusinessTypes.Visible = False
        .Clear
        For Each oLocation In goSystemConfig.Locations
            .AddRow False, _
                    oLocation.Key, _
                    oLocation.Description, _
                    oLocation.Prefix, _
                    oLocation.Suffix
        Next oLocation
        '
        .ResizeRows
        .Column(1).Sorted = smgAscending
        .GetKeyRow (mstrCurrentImageType)
        .Redraw = True
        sstBusinessTypes.Visible = True
    End With
'
'********* Code Ends Here **************
'
   Exit Sub
'
RefreshImageLocations_Error:
    ErrorRaise "frmToolsOptions.RefreshImageLocations"
End Sub

Private Sub RefreshBusinessTypes()
'***************************************
' Module/Form Name   : frmToolsOptions
'
' Procedure Name     : RefreshBusinessTypes
'
' Purpose            :
'
' Date Created       : 21/04/2002
'
' Author             : GARETH SAUNDERS
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 21/04/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo RefreshBusinessTypes_Error
'
'******** Code Starts Here *************
'
    Dim objBusinessType As BusinessType

    DoEvents
    With smgBusinessTypes
        .Redraw = False
        sstBusinessTypes.Visible = False
        .Clear
        For Each objBusinessType In mobjBusinessTypes
            .AddRow False, _
                    objBusinessType.Name, _
                    objBusinessType.InitialReturnPeriod, _
                    objBusinessType.SL1ReturnPeriod, _
                    objBusinessType.SL2ReturnPeriod, _
                    objBusinessType.Phone1ReturnPeriod, _
                    objBusinessType.Phone2ReturnPeriod, _
                    objBusinessType.LossFeeReturnPeriod
        Next objBusinessType
        .ResizeRows
        .Column(1).Sorted = smgAscending
        .GetKeyRow (mobjBusinessTypes.CurrentName)
        .Redraw = True
        sstBusinessTypes.Visible = True
    End With
    '
    '********* Code Ends Here **************
    '
    Exit Sub
    '
RefreshBusinessTypes_Error:
    sstBusinessTypes.Visible = True
    smgBusinessTypes.Redraw = True
    ErrorRaise "frmToolsOptions.RefreshBusinessTypes"
End Sub

Private Sub RefreshSupportAndFTPTabs()
'***************************************
' Module/Form Name   : frmToolsOptions
'
' Procedure Name     : RefreshSupportAndFTPTabs
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
On Error GoTo RefreshSupportTab_Error
'
'******** Code Starts Here *************
'
    With goSystemConfig
        txtDateLogFilePosted.Text = IIf(.DateLogFilePosted = 0, "", Format(.DateLogFilePosted, "dd/mm/yyyy hh:mm:ss"))
        txtLogFilePostingFrequency.Text = CStr(.LogFilePostingFrequency)
        txtSupportEmail.Text = .SupportEmail
        txtWebSearchTestEmail.Text = .WebSearchTestEmail
        txtUpgradeFrequency.Text = .UpgradeCheckFrequency
        txtDateEmailLastPosted.Text = IIf(.DateHTMLEmailPosted = 0, "", Format(.DateHTMLEmailPosted, "dd/mm/yyyy hh:mm:ss"))
        txtEmailsLastUpgraded.Text = IIf(.DateHTMLEmailUpgraded = 0, "", Format(.DateHTMLEmailUpgraded, "dd/mm/yyyy hh:mm:ss"))
    End With
    '
    '   FTP Settings.
    '
    With goSystemConfig
        txtFTPServer.Text = .FTPServer
        txtFTPUser.Text = .FTPUser
        txtFTPPassword.Text = .FTPPassword
        txtFTPRetypePassword.Text = .FTPPassword
        txtPostWebAddress.Text = .PostWebAddress
        txtWebSearchesEmailTo.Text = .WebSearchesEmailTo
        txtWebSearchesEmailFrom.Text = .WebSearchesEmailFrom
    End With
    '
    '   Data Fix Tab.
    '
    sstBusinessTypes.TabVisible(4) = GetSetting(App.Title, "ToolsOptions", "DataFix", "Y") = "Y"
'
'********* Code Ends Here **************
'
   Exit Sub
'
RefreshSupportTab_Error:
    ErrorRaise "frmToolsOptions.RefreshSupportTab"
End Sub

Private Sub smgBusinessTypes_DblClick()
    EditBusinessType
End Sub

Private Sub smgBusinessTypes_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'***************************************
' Module/Form Name   : frmToolsOptions
'
' Procedure Name     : smgBusinessTypes_MouseUp
'
' Purpose            :
'
' Date Created       : 02/06/2002
'
' Author             : GARETH SAUNDERS
'
' Parameters         : Button - Integer
'                    : Shift - Integer
'                    : X - Single
'                    : Y - Single
'
' Amendment History  : Date       Author    Description
'                    : --------------------------------
'                    : 02/06/2002 GARETH SAUNDERS
'
'***************************************
'
On Error GoTo smgBusinessTypes_MouseUp_Error
'
'******** Code Starts Here *************
'
    Dim i As Integer
    '
    '   If right mouse button clicked and an item selected, display pop up menu.
    '
    If Button <> vbRightButton Then
        Exit Sub
    End If
    '
    '   Display pop up menu.
    '
    PopupMenu mnuMaintain, vbPopupMenuRightButton, , , mnuMaintainEdit
'
'********* Code Ends Here **************
'
    Exit Sub
    '
smgBusinessTypes_MouseUp_Error:
    DisplayError , "frmToolsOptions.smgBusinessTypes_MouseUp", vbExclamation
End Sub

Private Sub smgBusinessTypes_RowChanged(CurrentRow As String)
    mobjBusinessTypes.CurrentName = smgBusinessTypes.Column(smgBusinessTypes.KeyCol).Value
End Sub

Private Sub smgDefaultImageLocations_DblClick()
'***************************************
' Module/Form Name   : frmToolsOptions
'
' Procedure Name     : smgDefaultImageLocations_DblClick
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
On Error GoTo smgDefaultImageLocations_DblClick_Error
'
'******** Code Starts Here *************
'
    EditLocation
'
'********* Code Ends Here **************
'
   Exit Sub
'
smgDefaultImageLocations_DblClick_Error:
    DisplayError , "frmToolsOptions.smgDefaultImageLocations_DblClick", vbExclamation
End Sub

Private Sub sstBusinessTypes_Click(PreviousTab As Integer)
    Dim Ctl         As Control
    Dim ctlTop      As Control
    
    If sstBusinessTypes.Tab = 0 Then
        smgBusinessTypes.Visible = True
        cmdEdit.Enabled = True
        cmdDelete.Enabled = True
        cmdEditLocation.Enabled = False
        cmdDeleteLocation.Enabled = False
    Else
        smgBusinessTypes.Visible = False
        cmdEdit.Enabled = False
        cmdDelete.Enabled = False
        cmdEditLocation.Enabled = True
        cmdDeleteLocation.Enabled = True
    End If
    '
    '   Disable controls not on the current tab to stop them getting focus.
    '
    On Error Resume Next
    For Each Ctl In Controls
        Set ctlTop = TopLevelControl(Me, Ctl)
        If TypeOf ctlTop.Container Is SSTab Then
            'Not all controls have the TabStop property
            Ctl.TabStop = ctlTop.Left > 0
        End If
    Next Ctl
End Sub

Private Sub txtAddress1_GotFocus()
    HighLightText txtAddress1
End Sub

Private Sub txtAddress2_GotFocus()
    HighLightText txtAddress2
End Sub

Private Sub txtAddress3_GotFocus()
    HighLightText txtAddress3
End Sub

Private Sub txtCountry_GotFocus()
    HighLightText txtCountry
End Sub

Private Sub txtCounty_GotFocus()
    HighLightText txtCounty
End Sub

Private Sub txtEmail_GotFocus()
    HighLightText txtEmail
End Sub

Private Sub txtFaxNo_GotFocus()
    HighLightText txtFaxNo
End Sub

Private Sub txtFTPPassword_GotFocus()
    HighLightText txtFTPPassword
End Sub

Private Sub txtFTPRetypePassword_GotFocus()
    HighLightText txtFTPRetypePassword
End Sub

Private Sub txtFTPServer_GotFocus()
    HighLightText txtFTPServer
End Sub

Private Sub txtFTPUser_GotFocus()
    HighLightText txtFTPUser
End Sub

Private Sub txtImageHeightWidth_GotFocus()
    HighLightText txtImageHeightWidth
End Sub

Private Sub txtImagesReturned_GotFocus()
    HighLightText txtImagesReturned
End Sub

Private Sub txtImagesReturned_KeyPress(KeyAscii As Integer)
    KeyAscii = allow_numeric_only(KeyAscii)
End Sub

Private Sub txtInfo1_GotFocus()
    HighLightText txtInfo1
End Sub

Private Sub txtInfo2_GotFocus()
    HighLightText txtInfo2
End Sub

Private Sub txtLogFilePostingFrequency_KeyPress(KeyAscii As Integer)
    KeyAscii = allow_numeric_only(KeyAscii)
End Sub

Private Sub txtMaxImagesPerPage_GotFocus()
    HighLightText txtMaxImagesPerPage
End Sub

Private Sub txtMaxImagesPerPage_KeyPress(KeyAscii As Integer)
    KeyAscii = allow_numeric_only(KeyAscii)
End Sub

Private Sub txtPostCode_GotFocus()
    HighLightText txtPostCode
End Sub

Private Sub txtSignatory_GotFocus()
    HighLightText txtSignatory
End Sub

Private Sub txtTelNo_GotFocus()
    HighLightText txtTelNo
End Sub

Private Sub txtTooltipDelay_KeyPress(KeyAscii As Integer)
    KeyAscii = allow_numeric_only(KeyAscii)
    If KeyAscii > Asc("4") Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtTown_GotFocus()
    HighLightText txtTown
End Sub

Private Sub txtUpgradeFrequency_KeyPress(KeyAscii As Integer)
    KeyAscii = allow_numeric_only(KeyAscii)
End Sub

Private Sub txtVATNo_GotFocus()
    HighLightText txtVATNo
End Sub

Private Sub txtPostWebAddress_GotFocus()
    HighLightText txtPostWebAddress
End Sub

Private Sub txtWebSearchesEmailFrom_GotFocus()
    HighLightText txtWebSearchesEmailFrom
End Sub

Private Sub txtWebSearchesEmailTo_GotFocus()
    HighLightText txtWebSearchesEmailTo
End Sub

Private Sub txtWebSite_GotFocus()
    HighLightText txtWebSite
End Sub

