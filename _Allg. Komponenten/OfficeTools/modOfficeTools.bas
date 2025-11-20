Attribute VB_Name = "modOfficeTools"
Option Explicit
Private Const MODULNAME = "modOfficeTools"                            ' Modulname für Fehlerbehandlung

Public Type OfficeInfo
    AccessVersion As String
    AccessPath As String
    WordVersion As String
    WordPath As String
    ExcelVersion As String
    ExcelPath As String
    OutlookVersion As String
    OutlookPath As String
End Type


