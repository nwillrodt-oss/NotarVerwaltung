Attribute VB_Name = "modTools"
Option Explicit
Private Const MODULNAME = "modTools"

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
    Call OBErrorhandler(MODULNAME, "CutLastChar", errNr, errDesc)
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
    Call OBErrorhandler(MODULNAME, "CheckNull", errNr, errDesc)
End Function
