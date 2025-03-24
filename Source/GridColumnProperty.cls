VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GridColumnProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Enum enmPropertyType
    smgPropTypeText
    smgPropTypeBooleanGraphic
    smgPropTypeDateShort
    smgPropTypeDateLong
End Enum

Private mvarProperyType As enmPropertyType

Public Property Get PropertyType() As enmPropertyType
    PropertyType = mvarProperyType
End Property

Public Property Let PropertyType(vData As enmPropertyType)
    mvarProperyType = vData
End Property

