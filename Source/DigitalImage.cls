VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DigitalImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Option Explicit

Private mvarBatchNo                 As Long
Private mvarDescription             As String
Private mvarCredit                  As String
Private mvarRelativeFileLocation    As String
Private mvarNewFileLocation         As String
Private mvarSelected                As Boolean
Private mvarTextControlIndex        As Integer
Private mvarImageControlIndex       As Variant
Private mvarDigitalImageIndex       As Integer
Private mvarPicture                 As Picture
Private mvarMaxHeight               As Integer
Private mvarMaxWidth                As Integer
Private mvarImageAvailable          As Boolean
Private mvarPending                 As enmPending
Private mvarPhotoTypeRequired       As String
Private mvarSearchResultKey         As Long
Private mvarExclusive               As Boolean
'
'   For use in view pictuire.
'
Private mvarDeliveryNoteNo          As Long
Private mvarCustomerNo              As Integer
Private mvarCustomerName            As String
Private mvarAddressLine1            As String
Private mvarDigitalStatus           As String

Private oFSO As Scripting.FileSystemObject

Public Property Let DigitalImageIndex(ByVal vData As Integer)
    mvarDigitalImageIndex = vData
End Property

Public Property Get DigitalImageIndex() As Integer
    DigitalImageIndex = mvarDigitalImageIndex
End Property

Public Property Let Exclusive(ByVal vData As Boolean)
    mvarExclusive = vData
End Property

Public Property Get Exclusive() As Boolean
    Exclusive = mvarExclusive
End Property

Public Property Get ImageAvailable() As Boolean
    ImageAvailable = mvarImageAvailable
End Property

Public Property Let ImageControlIndex(ByVal vData As Variant)
    mvarImageControlIndex = vData
End Property

Public Property Set ImageControlIndex(ByVal vData As Object)
    Set mvarImageControlIndex = vData
End Property

Public Property Get ImageControlIndex() As Variant
    If IsObject(mvarImageControlIndex) Then
        Set ImageControlIndex = mvarImageControlIndex
    Else
        ImageControlIndex = mvarImageControlIndex
    End If
End Property

Public Property Let TextControlIndex(ByVal vData As Integer)
    mvarTextControlIndex = vData
End Property

Public Property Get TextControlIndex() As Integer
    TextControlIndex = mvarTextControlIndex
End Property

Public Property Let Selected(ByVal vData As Boolean)
    mvarSelected = vData
End Property

Public Property Get Selected() As Boolean
    Selected = mvarSelected
End Property

Public Property Get FileLocation(Optional ByVal pstrResolution As String = "TH") As String
    Dim strDefaultFileLocation As String
    
    On Error GoTo ErrorProc
    strDefaultFileLocation = DefaultFileLocation(pstrResolution)
    
    If oFSO.FileExists(strDefaultFileLocation) Then
        FileLocation = strDefaultFileLocation
    Else
        FileLocation = ""
    End If
    
    Exit Property
ErrorProc:
    ErrorRaise "DigitalImage: Get FileLocation"
''    Err.Raise Err.Number, "DigitalImage: Get FileLocation", Err.Description
End Property

Public Property Get DefaultFileLocation(Optional ByVal pstrResolution As String = "TH") As String
    Dim strDefaultFileLocation As String
    Dim strRelativeFileLocation As String
    
    On Error GoTo ErrorProc
    strRelativeFileLocation = oFSO.BuildPath(Folder(pstrResolution), CStr(mvarBatchNo) & ".jpg")
    strDefaultFileLocation = oFSO.BuildPath(goSystemConfig.Locations("IMGE" & pstrResolution).Prefix, strRelativeFileLocation)
    DefaultFileLocation = strDefaultFileLocation
    
    Exit Property
ErrorProc:
    ErrorRaise "DigitalImage: Get FileLocation"
''    Err.Raise Err.Number, "DigitalImage: Get FileLocation", Err.Description
End Property

Public Property Let Description(ByVal vData As String)
    mvarDescription = vData
End Property

Public Property Get Description() As String
    Description = mvarDescription
End Property

Public Property Let Credit(ByVal vData As String)
    mvarCredit = vData
End Property

Public Property Get Credit() As String
    Credit = mvarCredit
End Property

Public Property Set Picture(vData As Picture)
    Set mvarPicture = vData
End Property

Public Property Get Height() As Integer
    If mvarPicture Is Nothing Then
        Height = mvarMaxHeight
        Exit Sub
    ElseIf mvarPicture = 0 Then
        Height = mvarMaxHeight
        Exit Sub
    ElseIf mvarMaxHeight = 0 Then
        Height = 0
        Exit Sub
    End If
    
    If mvarMaxWidth / mvarMaxHeight < mvarPicture.Width / mvarPicture.Height Then
        Height = mvarPicture.Height / mvarPicture.Width * mvarMaxWidth
    Else
        Height = mvarMaxHeight
    End If
End Property

Public Property Get Width() As Integer
    If mvarPicture Is Nothing Then
        Width = mvarMaxWidth
        Exit Sub
    ElseIf mvarPicture = 0 Then
        Width = mvarMaxWidth
        Exit Sub
    ElseIf mvarMaxWidth = 0 Then
        Width = 0
        Exit Sub
    End If
    
    If mvarMaxWidth / mvarMaxHeight < mvarPicture.Width / mvarPicture.Height Then
        Width = mvarMaxWidth
    Else
        Width = mvarPicture.Width / mvarPicture.Height * mvarMaxHeight
    End If
End Property

Public Property Let MaxHeight(vData As Integer)
    mvarMaxHeight = vData
End Property

Public Property Let MaxWidth(vData As Integer)
    mvarMaxWidth = vData
End Property


Private Sub Class_Initialize()
    mvarTextControlIndex = -1
    mvarImageControlIndex = -1
    mvarDigitalImageIndex = -1
    mvarBatchNo = 0
    mvarRelativeFileLocation = ""
    mvarNewFileLocation = ""
    mvarMaxHeight = 0
    mvarMaxWidth = 0
    mvarPending = pdgClean
    mvarPhotoTypeRequired = "T"
    mvarSearchResultKey = 0
    Set oFSO = New Scripting.FileSystemObject
End Sub

Public Property Get Folder(Optional ByRef pstrResolution As String = "TH") As String
    Folder = Format(Int(mvarBatchNo / 1000), "00") & "K\" & goSystemConfig.Locations("IMGE" & pstrResolution).Suffix
End Property

Public Sub Associate(Optional Force As Boolean = False)
    
    On Error GoTo ErrorProc
    '
    '   If there is no change don't do anything.
    '
    If mvarNewFileLocation = "" Then
        Exit Sub
    End If
    If LCase(mvarNewFileLocation) = LCase(DefaultFileLocation) Then
        Exit Sub
    End If
    If Not oFSO.FolderExists(StripFile(DefaultFileLocation)) Then
        Err.Raise vbObjectError + 1, , "Folder '" & StripFile(DefaultFileLocation) & "' does not exist." & vbCrLf & "Please Create."
    End If
    
    If oFSO.FileExists(DefaultFileLocation) Then
        If Force Then
            '
            '   Move the file to the Holding Area giving it the same name with the date
            '   and time in the format 'x(dd-mm-yyyy hh-mm-ss).jpg'.
            '
            'oFSO.MoveFile DefaultFileLocation, oFSO.BuildPath(App.Path, goSystemConfig.ImageHoldingArea) & "\" & CStr(mvarBatchNo) & "(" & Format(Now, "dd-mm-yyyy hh-mm-ss") & ").jpg"
            oFSO.MoveFile DefaultFileLocation, oFSO.BuildPath(goSystemConfig.ServerLocation, goSystemConfig.ImageHoldingArea) & "\" & CStr(mvarBatchNo) & "(" & Format(Now, "dd-mm-yyyy hh-mm-ss") & ").jpg"
            'oFSO.DeleteFile DefaultFileLocation
        Else
            Err.Raise vbObjectError + 2, , "Digital Image already exists in '" & DefaultFileLocation & "'"
        End If
    End If
    '
    '   If the New File Location is in the Holding Area then Move the Image. Otherwise Copy.
    '
    'If LCase(StripFile(mvarNewFileLocation)) = LCase(oFSO.BuildPath(App.Path, goSystemConfig.ImageHoldingArea)) Then
    If LCase(StripFile(mvarNewFileLocation)) = LCase(oFSO.BuildPath(goSystemConfig.ServerLocation, goSystemConfig.ImageHoldingArea)) Then
        oFSO.MoveFile mvarNewFileLocation, DefaultFileLocation
    Else
        oFSO.CopyFile mvarNewFileLocation, DefaultFileLocation
    End If
    
    Exit Sub
ErrorProc:
    Err.Raise Err.Number, "DigitalImage: Associate", Err.Description
End Sub

Public Property Let BatchNo(vData As Long)
    On Error GoTo ErrorProc
    mvarBatchNo = vData
    If FileLocation = "" Then
        mvarImageAvailable = False
    Else
        mvarImageAvailable = True
    End If
    
    Exit Property
ErrorProc:
    ErrorRaise "DigitalImage: Let BatchNo"
''    Err.Raise Err.Number, "DigitalImage: Let BatchNo", Err.Description
End Property

Public Property Get BatchNo() As Long
    BatchNo = mvarBatchNo
End Property

Public Property Let NewFileLocation(vData As String)
    mvarNewFileLocation = vData
End Property

Public Property Get NewFileLocation() As String
    NewFileLocation = mvarNewFileLocation
End Property

Public Property Get Pending() As enmPending
    Pending = mvarPending
End Property

Public Property Let Pending(ByVal vData As enmPending)
    mvarPending = vData
End Property

Public Property Let PhotoTypeRequired(vData As String)
    mvarPhotoTypeRequired = vData
End Property

Public Property Get PhotoTypeRequired() As String
    PhotoTypeRequired = mvarPhotoTypeRequired
End Property

Public Property Let SearchResultKey(vData As Long)
    mvarSearchResultKey = vData
End Property

Public Property Get SearchResultKey() As Long
    SearchResultKey = mvarSearchResultKey
End Property

Public Property Let DeliveryNoteNo(ByVal vData As Long)
    mvarDeliveryNoteNo = vData
End Property

Public Property Get DeliveryNoteNo() As Long
    DeliveryNoteNo = mvarDeliveryNoteNo
End Property

Public Property Let CustomerNo(ByVal vData As Integer)
    mvarCustomerNo = vData
End Property

Public Property Get CustomerNo() As Integer
    CustomerNo = mvarCustomerNo
End Property

Public Property Let CustomerName(ByVal vData As String)
    mvarCustomerName = vData
End Property

Public Property Get CustomerName() As String
    CustomerName = mvarCustomerName
End Property

Public Property Let AddressLine1(ByVal vData As String)
    mvarAddressLine1 = vData
End Property

Public Property Get AddressLine1() As String
    AddressLine1 = mvarAddressLine1
End Property

Public Property Let DigitalStatus(ByVal vData As String)
    mvarDigitalStatus = vData
End Property

Public Property Get DigitalStatus() As String
    DigitalStatus = mvarDigitalStatus
End Property


