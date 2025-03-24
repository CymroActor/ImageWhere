Attribute VB_Name = "CommonRoutines"
Option Explicit

Public Function fgIndex(msgGrid As MSFlexGrid, row As Integer, col As Integer) As Long
     fgIndex = row * msgGrid.Cols + col
End Function


