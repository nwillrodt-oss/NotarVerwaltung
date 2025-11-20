VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const MODULNAME = "clsTools"
Private objObjectBag As clsObjectBag

Public Function InitTools(ObjectBag As clsObjectBag)

    Set objObjectBag = ObjectBag

End Function

Public Function CutLastChar(szStr As String, szChar As String) As String
    
    Dim intCharLen As Integer
    
On Error GoTo ErrorHandler
    intCharLen = Len(szChar)
    
    If Right(szStr, intCharLen) = szChar Then szStr = Left(szStr, Len(szStr) - intCharLen)
    CutLastChar = szStr
exithandler:

Exit Function
ErrorHandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objObjectBag.OBErrorhandler(MODULNAME, "CutLastChar", errNr, errDesc)
    Resume exithandler
End Function


Public Function CheckNull(ReturnValue) As Variant

On Error GoTo ErrorHandler

    If IsNull(ReturnValue) Then
        CheckNull = ""
    Else
        CheckNull = ReturnValue
    End If

Exit Function
ErrorHandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objObjectBag.OBErrorhandler(MODULNAME, "CheckNull", errNr, errDesc)
End Function
