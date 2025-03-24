Attribute VB_Name = "global_data"
Global glo_dbname As String
Global gADOConnection As String
Global ws As Workspace
Global db As Database
Global snap_customer As DAO.Recordset
'
'   Used for displaying the action being performed when the failure occurred.
'
Global Action As String
'
'   Used for determining the width of the scroll bar when resizing screens.
'
Global scroll_bar_width As Integer
'
'   Used for deciding whether to retry certain things e.g. entry of customer name.
'
Global retry As Boolean
'
'   Used for the response back from invoked methods.
'
Global abort As Boolean
'
'   Used for Yes or No type message boxes.
'
Global msg As String
Global style As String
Global response As Integer
Global msg_title
'
'   Global Objects
'
Global goSystemConfig   As New SystemConfig
Global goCompanyInfo    As CompanyInfo
'
'   ADO Database Object.
'
Global gdbADO   As ADODB.Connection
Global goLog    As New Log
'
'   Used for determining whether the error has been generated at a lower level or not.
'
Global gLowerLevelError As Boolean
'
Global Const CONST_APPLICATION = "Image Where"
Global gstrAppPath As String
Global gstrUpgradePath As String
Global gblnInhibitSubClassing As Boolean
'
'   Used to hold the Window Handles of windows that have been resized to the max.
'
Global gcolMaxedWindows As Collection
Global gblnResizeMaxedWindows As Boolean
'
Public Type typJobFilter
    CustomerNo      As Integer
    ShowAll         As Boolean
    ShowOpenOnly    As Boolean
End Type

Public Type typKeywordSynonym
    Keyword         As String
    KeywordType     As String
End Type

Public gintTransCount   As Integer
